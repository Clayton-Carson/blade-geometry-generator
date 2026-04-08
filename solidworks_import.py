"""
SolidWorks 2021 COM Automation: S-64 Tail Rotor Blade Import
=============================================================
Imports blade section curves, creates spanwise guide curves, and generates
boundary surfaces via the SolidWorks API (win32com).

Creates a single part with two surface bodies:
  - Base OML surface (lofted through base_S##.txt cross-sections)
  - Nickel cap surface (lofted through cap_S##.txt cross-sections)

Each surface uses cross-section profile curves and spanwise guide curves
(LE, TE, and intermediate guides on upper/lower surfaces).

Requirements:
  - SolidWorks 2021 installed and licensed
  - pywin32: pip install pywin32
  - numpy
  - Output curves from blade_section_generator.py in output/curves/

Usage:
  python solidworks_import.py                        # full pipeline, both bodies
  python solidworks_import.py --phase curves          # just import curve files
  python solidworks_import.py --phase guides          # curves + guide curves
  python solidworks_import.py --phase surfaces        # curves + guides + surfaces
  python solidworks_import.py --body base             # only base OML
  python solidworks_import.py --body cap              # only nickel cap
  python solidworks_import.py --dry-run               # validate files, print plan
  python solidworks_import.py --save output/blade.sldprt
"""

import os
import sys
import time
import logging
import argparse
from pathlib import Path

import numpy as np

log = logging.getLogger(__name__)

# ─── Unit conversion ────────────────────────────────────────
INCHES_TO_METERS = 0.0254

# ─── SolidWorks API constants ───────────────────────────────
MARK_PROFILE = 1
MARK_GUIDE = 2

# Guide curve configuration
UPPER_GUIDE_INDICES = [6, 12, 18, 24]  # indices safe on all stations (0-29)
LOWER_GUIDE_XC = [0.15, 0.35, 0.55, 0.75]  # normalized chordwise positions
NUM_STATIONS = 8


# ─── Data loading ─────────���─────────────────────────────────

def read_curve_points(filepath):
    """Read tab-separated XYZ points from a curve file. Returns (N,3) array in inches."""
    pts = []
    with open(filepath) as f:
        for line in f:
            parts = line.strip().split('\t')
            if len(parts) == 3:
                pts.append([float(x) for x in parts])
    return np.array(pts)


def find_le_index(pts):
    """Find leading edge index (minimum X coordinate)."""
    return int(np.argmin(pts[:, 0]))


def get_lower_surface(pts):
    """Return lower surface points (LE to TE lower, inclusive)."""
    le = find_le_index(pts)
    return pts[le:]


def interpolate_lower_at_xc(lower_pts, xc_target):
    """
    Interpolate a 3D point on the lower surface at a target normalized
    chordwise position (0 = LE, 1 = TE).
    """
    x_le = lower_pts[0, 0]
    x_te = lower_pts[-1, 0]
    denom = x_te - x_le
    if abs(denom) < 1e-12:
        return lower_pts[0].copy()
    xc = (lower_pts[:, 0] - x_le) / denom
    for i in range(len(xc) - 1):
        if xc[i] <= xc_target <= xc[i + 1]:
            t = (xc_target - xc[i]) / (xc[i + 1] - xc[i] + 1e-15)
            return lower_pts[i] + t * (lower_pts[i + 1] - lower_pts[i])
    idx = int(np.argmin(np.abs(xc - xc_target)))
    return lower_pts[idx].copy()


def discover_curve_files(curves_dir, prefix, n_stations=NUM_STATIONS):
    """Find and validate curve files. Returns list of Paths."""
    files = []
    for i in range(1, n_stations + 1):
        p = curves_dir / f'{prefix}_S{i:02d}.txt'
        if not p.exists():
            raise FileNotFoundError(f"Missing curve file: {p}")
        files.append(p)
    return files


# ─── SolidWorks COM helpers ─────────────────────────────────

def connect_solidworks():
    """Connect to running SolidWorks. Returns (sw, model_getter) where
    model_getter() always returns a fresh reference to the active document."""
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()

    sw = None

    # Try attaching to a running instance first
    try:
        sw = win32com.client.GetObject(Class="SldWorks.Application")
        log.info("Attached to running SolidWorks instance")
    except Exception:
        pass

    # Launch new instance
    if sw is None:
        for progid in ["SldWorks.Application", "SldWorks.Application.29"]:
            try:
                sw = win32com.client.Dispatch(progid)
                log.info(f"Launched SolidWorks via {progid}")
                break
            except Exception:
                continue

    if sw is None:
        raise RuntimeError(
            "Cannot connect to SolidWorks. Ensure it is installed and licensed.\n"
            "Tried: GetObject, Dispatch('SldWorks.Application'), "
            "Dispatch('SldWorks.Application.29')"
        )

    sw.Visible = True
    # Wait for SW to be ready
    for _ in range(30):
        try:
            if sw.GetUserPreferenceIntegerValue(10) is not None:
                break
        except Exception:
            pass
        time.sleep(1)
    else:
        log.warning("SolidWorks may not be fully initialized yet")

    return sw


def create_part(sw):
    """Create a new empty part document. Returns sw (model accessed via sw.ActiveDoc)."""
    # Search common install locations for part template
    template = None
    fallback_dirs = [
        r"C:\ProgramData\SolidWorks\SOLIDWORKS 2021\templates",
        r"C:\ProgramData\SolidWorks\SOLIDWORKS 2022\templates",
        r"C:\ProgramData\SolidWorks\SOLIDWORKS 2023\templates",
        r"C:\ProgramData\SolidWorks\SOLIDWORKS 2024\templates",
        r"C:\ProgramData\SolidWorks\SOLIDWORKS 2025\templates",
    ]
    for d in fallback_dirs:
        candidate = os.path.join(d, "Part.PRTDOT")
        if os.path.exists(candidate):
            template = candidate
            log.info(f"Using template: {template}")
            break

    if not template:
        raise FileNotFoundError(
            "Cannot find SolidWorks part template. "
            "Set a default template in SolidWorks Options > Default Templates."
        )

    log.info(f"Creating new part from template: {template}")
    sw.NewDocument(template, 0, 0, 0)

    model = sw.ActiveDoc
    if model is None:
        raise RuntimeError("Failed to create new part document")

    # Set units to IPS (inches)
    try:
        model.Extension.SetUserPreferenceInteger(197, 0, 1)
        log.info("Set document units to IPS (inches)")
    except Exception as e:
        log.warning(f"Could not set units: {e}")

    log.info("Created new part document")
    return model


def create_3d_spline(sw, points_inches, sketch_name=None):
    """
    Create a 3D sketch containing a spline through the given points.
    points_inches: (N, 3) numpy array in inches.
    API expects meters, so we convert here.

    Returns the sketch_name used.
    """
    import win32com.client
    import pythoncom

    model = sw.ActiveDoc

    # Open a new 3D sketch
    model.Insert3DSketch2(True)

    # Convert inches to meters and flatten
    pts_m = (points_inches * INCHES_TO_METERS).flatten().tolist()

    # Build VARIANT SAFEARRAY of doubles
    pt_array = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, pts_m
    )

    # CreateSpline2 creates a spline through the given points
    sketch_mgr = model.SketchManager
    spline = sketch_mgr.CreateSpline2(pt_array, False)

    if spline is None:
        log.warning(f"CreateSpline2 returned None for {len(points_inches)} points")

    # Close the 3D sketch
    model.Insert3DSketch2(True)

    # Rename the sketch if requested — get it via selection
    if sketch_name:
        try:
            # After closing a 3D sketch, it's typically still selected
            sel_mgr = model.SelectionManager
            count = sel_mgr.GetSelectedObjectCount2(-1)
            if count > 0:
                feat = sel_mgr.GetSelectedObject6(1, -1)
                if feat is not None:
                    feat.Name = sketch_name
                    log.debug(f"  Renamed sketch to {sketch_name}")
        except Exception as e:
            log.debug(f"  Could not rename sketch: {e}")

    return sketch_name


def select_by_name(sw, feat_name, feat_type="SKETCH", mark=0, append=True):
    """Select a feature by name using SelectByID2 (reliable with late-bound COM)."""
    model = sw.ActiveDoc
    try:
        result = model.Extension.SelectByID2(
            feat_name, feat_type, 0, 0, 0, append, mark, None, 0
        )
        if result:
            log.debug(f"  Selected '{feat_name}' (mark={mark})")
        else:
            log.warning(f"  SelectByID2 failed for '{feat_name}'")
        return result
    except Exception as e:
        log.warning(f"  SelectByID2 error for '{feat_name}': {e}")
        return False


# ─── Phase: Import Curves ───────────────────────────────────

def import_curves(sw, curves_dir, body_type):
    """
    Import cross-section curves as 3D sketch splines.
    Returns list of (station_id, sketch_name) tuples.
    """
    prefix = body_type
    files = discover_curve_files(curves_dir, prefix)
    sketches = []

    log.info(f"Importing {len(files)} {body_type} curves as 3D sketch splines...")
    for i, fpath in enumerate(files):
        stn_id = i + 1
        name = f"{body_type}_S{stn_id:02d}"
        log.info(f"  Importing {fpath.name} -> {name}...")

        pts = read_curve_points(fpath)
        if len(pts) == 0:
            log.error(f"  No points in {fpath.name}")
            sketches.append((stn_id, None))
            continue

        create_3d_spline(sw, pts, sketch_name=name)
        sketches.append((stn_id, name))
        log.info(f"  Created {name} ({len(pts)} points)")

    return sketches


# ─── Phase: Guide Curves ─────────��──────────────────────────

def build_guide_points(curves_dir, body_type):
    """
    Compute guide curve point arrays from the curve files.
    Returns a dict of {guide_name: (N_stations, 3) array in inches}.
    """
    prefix = body_type
    files = discover_curve_files(curves_dir, prefix)
    all_pts = [read_curve_points(f) for f in files]

    guides = {}

    # TE Upper: index 0 on every station
    guides['TE_upper'] = np.array([pts[0] for pts in all_pts])

    # TE Lower: last index on every station
    guides['TE_lower'] = np.array([pts[-1] for pts in all_pts])

    # Leading Edge: min-X point per station
    guides['LE'] = np.array([pts[find_le_index(pts)] for pts in all_pts])

    # Upper surface intermediate guides (index-based, safe for 0-29)
    for idx in UPPER_GUIDE_INDICES:
        guides[f'upper_{idx:02d}'] = np.array([pts[idx] for pts in all_pts])

    # Lower surface intermediate guides (x/c-interpolated)
    for xc in LOWER_GUIDE_XC:
        pts_interp = []
        for pts in all_pts:
            lower = get_lower_surface(pts)
            pt = interpolate_lower_at_xc(lower, xc)
            pts_interp.append(pt)
        guides[f'lower_xc{xc:.2f}'] = np.array(pts_interp)

    return guides


def create_guide_curves(sw, curves_dir, body_type):
    """
    Create all guide curves as 3D sketch splines.
    Returns list of (guide_name, sketch_name) tuples.
    """
    guide_pts = build_guide_points(curves_dir, body_type)
    sketches = []

    log.info(f"Creating {len(guide_pts)} {body_type} guide curves...")
    for name, pts in guide_pts.items():
        sketch_name = f"{body_type}_guide_{name}"
        log.info(f"  Creating {sketch_name} ({len(pts)} points)...")
        create_3d_spline(sw, pts, sketch_name=sketch_name)
        sketches.append((name, sketch_name))

    return sketches


# ─── Phase: Surfaces ─────────────��──────────────────────────

def create_surface(sw, profile_sketches, guide_sketches, body_type):
    """
    Create a boundary surface (or loft fallback) from profiles and guides.
    profile_sketches: list of (stn_id, sketch_name) from import_curves
    guide_sketches: list of (name, sketch_name) from create_guide_curves
    """
    model = sw.ActiveDoc
    log.info(f"Creating {body_type} surface...")

    # Clear any existing selection
    model.ClearSelection2(True)

    # Select profiles with mark = 1
    profile_count = 0
    for stn_id, sketch_name in profile_sketches:
        if sketch_name is None:
            log.warning(f"  Skipping missing profile S{stn_id:02d}")
            continue
        if select_by_name(sw, sketch_name, "SKETCH", MARK_PROFILE, append=(profile_count > 0)):
            profile_count += 1

    # Select guides with mark = 2
    guide_count = 0
    for name, sketch_name in guide_sketches:
        if sketch_name is None:
            continue
        if select_by_name(sw, sketch_name, "SKETCH", MARK_GUIDE, append=True):
            guide_count += 1

    log.info(f"  Selected {profile_count} profiles, {guide_count} guides")

    if profile_count < 2:
        log.error("  Need at least 2 profiles for surface creation")
        return None

    feat_mgr = model.FeatureManager
    surface_feat = None

    # Try boundary surface
    try:
        log.info("  Attempting boundary surface...")
        surface_feat = feat_mgr.InsertBoundarySurface(False, False)
    except Exception as e:
        log.warning(f"  Boundary surface failed: {e}")

    # Fallback: loft with guides
    if surface_feat is None:
        log.info("  Falling back to loft surface...")
        model.ClearSelection2(True)
        for stn_id, sketch_name in profile_sketches:
            if sketch_name is not None:
                select_by_name(sw, sketch_name, "SKETCH", MARK_PROFILE, append=True)
        try:
            surface_feat = feat_mgr.InsertProtrusionBlend2(
                False, True, False, 1.0, 0, 0, 1, True, 0,
                0.0, 0.0, False, 0.0, 0.0, 0, True, True, 0, 0,
            )
        except Exception as e:
            log.warning(f"  Loft also failed: {e}")

    # Final fallback: simple loft
    if surface_feat is None:
        log.info("  Attempting simple loft (no guides)...")
        model.ClearSelection2(True)
        for stn_id, sketch_name in profile_sketches:
            if sketch_name is not None:
                select_by_name(sw, sketch_name, "SKETCH", MARK_PROFILE, append=True)
        try:
            surface_feat = model.InsertLoftSurface2(
                False, False, False, True, 0, 0, 0.0, 0.0, False, 0.0, 0.0, 0
            )
        except Exception as e:
            log.error(f"  All surface creation methods failed: {e}")
            log.error(
                "  You can create the surface manually: select the imported "
                "curves in order, then Insert > Surface > Boundary Surface."
            )
            return None

    if surface_feat is not None:
        log.info(f"  Created {body_type} surface successfully")

    return surface_feat


# ─── Orchestrator ──────────────────────────────────────────��─

def run(args):
    """Execute the requested phases."""
    curves_dir = args.curves_dir.resolve()

    # Determine which body types to process
    if args.body == 'both':
        body_types = ['base', 'cap']
    else:
        body_types = [args.body]

    # Determine phases to run
    phase = args.phase
    do_curves = phase in ('curves', 'guides', 'surfaces', 'all')
    do_guides = phase in ('guides', 'surfaces', 'all')
    do_surfaces = phase in ('surfaces', 'all')

    # ── Validate files ──
    log.info("Validating curve files...")
    for bt in body_types:
        files = discover_curve_files(curves_dir, bt)
        for f in files:
            pts = read_curve_points(f)
            log.info(f"  {f.name}: {len(pts)} points")

    if args.dry_run:
        log.info("\n── DRY RUN PLAN ──")
        log.info(f"Curves directory: {curves_dir}")
        log.info(f"Body types: {body_types}")
        log.info(f"Phases: curves={do_curves} guides={do_guides} surfaces={do_surfaces}")
        for bt in body_types:
            guides = build_guide_points(curves_dir, bt)
            log.info(f"\n{bt.upper()} guide curves ({len(guides)}):")
            for name, pts in guides.items():
                log.info(f"  {name}: {len(pts)} stations, "
                         f"Z range [{pts[0,2]:.3f} .. {pts[-1,2]:.3f}]")
        log.info("\nDry run complete. No SolidWorks operations performed.")
        return

    # ── Connect to SolidWorks ──
    sw = connect_solidworks()
    model = create_part(sw)

    # Storage for sketches per body type
    all_profiles = {}
    all_guides = {}

    for bt in body_types:
        log.info(f"\n{'='*50}")
        log.info(f"Processing {bt.upper()} body")
        log.info(f"{'='*50}")

        if do_curves:
            all_profiles[bt] = import_curves(sw, curves_dir, bt)

        if do_guides:
            all_guides[bt] = create_guide_curves(sw, curves_dir, bt)

        if do_surfaces:
            profiles = all_profiles.get(bt, [])
            guides = all_guides.get(bt, [])
            if profiles:
                create_surface(sw, profiles, guides, bt)
            else:
                log.warning(f"No profiles available for {bt} surface — skipping")

    # ── Final cleanup ──
    model = sw.ActiveDoc
    try:
        model.ForceRebuild3(False)
        model.ViewZoomtofit2()
    except Exception:
        pass

    if args.save:
        save_path = str(Path(args.save).resolve())
        log.info(f"Saving to {save_path}...")
        try:
            errors = model.Extension.SaveAs3(
                save_path, 0, 0, None, None, None
            )
            if errors == 0:
                log.info("Saved successfully")
            else:
                log.warning(f"Save returned error code: {errors}")
        except Exception as e:
            log.warning(f"Save failed: {e}")

    log.info("\nDone.")


# ─── CLI ─────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Import S-64 tail rotor blade curves into SolidWorks 2021"
    )
    parser.add_argument(
        '--phase',
        choices=['curves', 'guides', 'surfaces', 'all'],
        default='all',
        help="Which phases to execute (default: all)"
    )
    parser.add_argument(
        '--body',
        choices=['base', 'cap', 'both'],
        default='both',
        help="Which body types to process (default: both)"
    )
    parser.add_argument(
        '--curves-dir',
        type=Path,
        default=Path(__file__).parent / 'output' / 'curves',
        help="Path to curve files directory"
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help="Validate files and print plan without touching SolidWorks"
    )
    parser.add_argument(
        '--save',
        type=Path,
        default=None,
        help="Save part to this .sldprt path"
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help="Enable debug logging"
    )

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%H:%M:%S'
    )

    try:
        run(args)
    except FileNotFoundError as e:
        log.error(str(e))
        sys.exit(1)
    except RuntimeError as e:
        log.error(str(e))
        sys.exit(1)
    except Exception as e:
        log.error(f"Unexpected error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()
