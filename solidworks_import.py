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
# Document types
SW_DOC_PART = 1
# Unit system IDs
SW_UNITS_IPS = 3  # Inch-Pound-Second
# User preference integer IDs
SW_UNITS_LINEAR = 197
# Length units
SW_INCHES = 1
# Selection marks for boundary surface / loft
MARK_PROFILE = 1
MARK_GUIDE = 2
# Feature folder types
SW_FOLDER_EMPTY_BEFORE = 1

# Guide curve configuration
UPPER_GUIDE_INDICES = [6, 12, 18, 24]  # indices safe on all stations (0-29)
LOWER_GUIDE_XC = [0.15, 0.35, 0.55, 0.75]  # normalized chordwise positions
NUM_STATIONS = 8


# ─── Data loading ───────────────────────────────────────────

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
    """Connect to running SolidWorks or launch a new instance. Returns sw app object."""
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
        if sw.GetUserPreferenceIntegerValue(10) is not None:
            break
        time.sleep(1)
    else:
        log.warning("SolidWorks may not be fully initialized yet")

    return sw


def create_part(sw):
    """Create a new empty part document with IPS units. Returns model."""
    # Get default part template
    template = sw.GetUserPreferenceStringValue(7)  # swDefaultTemplatePart
    if not template or not os.path.exists(template):
        # Common fallback paths for SW 2021
        fallbacks = [
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2021\templates\Part.prtdot",
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2021\templates\part.prtdot",
        ]
        template = None
        for fb in fallbacks:
            if os.path.exists(fb):
                template = fb
                break
        if template is None:
            raise FileNotFoundError(
                "Cannot find SolidWorks part template. "
                "Set a default template in SolidWorks Options > Default Templates."
            )

    model = sw.NewDocument(template, 0, 0, 0)
    model = sw.ActiveDoc
    if model is None:
        raise RuntimeError("Failed to create new part document")

    log.info("Created new part document")
    return model


def set_units_ips(model):
    """Set document units to IPS (inches)."""
    ext = model.Extension
    # swUserPreferenceIntegerValue_e.swUnitsLinear = 197
    # swUserPreferenceOption_e.swDetailingNoOptionSpecified = 0
    # swLengthUnit_e.swINCHES = 1
    ext.SetUserPreferenceInteger(197, 0, 1)
    log.info("Set document units to IPS (inches)")


def get_last_feature(model):
    """Get the most recently added feature."""
    feat = model.Extension.GetLastFeatureAdded()
    if feat is not None:
        return feat
    # Fallback: walk to last feature
    feat = model.FirstFeature()
    last = feat
    while feat is not None:
        last = feat
        feat = feat.GetNextFeature()
    return last


def rename_feature(feat, name):
    """Rename a feature, appending a suffix if the name already exists."""
    if feat is None:
        return
    try:
        feat.Name = name
    except Exception:
        try:
            feat.Name = f"{name}_1"
        except Exception:
            log.warning(f"Could not rename feature to '{name}'")


def create_3d_spline(model, points_inches):
    """
    Create a 3D sketch containing a spline through the given points.
    points_inches: (N, 3) numpy array in inches.
    API expects meters, so we convert here.

    Returns the sketch feature.
    """
    import win32com.client
    import pythoncom

    sketch_mgr = model.SketchManager

    # Open a new 3D sketch
    model.Insert3DSketch2(True)

    # Convert inches to meters and flatten
    pts_m = (points_inches * INCHES_TO_METERS).flatten().tolist()
    n_pts = len(points_inches)

    # Build VARIANT SAFEARRAY of doubles
    pt_array = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, pts_m
    )

    # CreateSpline2(PointData) — creates a spline through the given points
    spline = sketch_mgr.CreateSpline2(pt_array, False)

    if spline is None:
        log.warning(f"CreateSpline2 returned None for {n_pts} points")

    # Close the 3D sketch
    model.Insert3DSketch2(True)

    return get_last_feature(model)


def select_feature(model, feat, mark, append=True):
    """Select a feature with a specific selection mark."""
    if feat is None:
        return False
    sel_mgr = model.SelectionManager
    sel_data = sel_mgr.CreateSelectData()
    sel_data.Mark = mark
    return feat.Select4(append, sel_data)


# ─── Phase: Import Curves ───────────────────────────────────

def import_curves(model, curves_dir, body_type):
    """
    Import cross-section curves via InsertCurveFilePoint.
    body_type: 'base' or 'cap'
    Returns list of (station_id, feature) tuples.
    """
    prefix = body_type
    files = discover_curve_files(curves_dir, prefix)
    features = []

    log.info(f"Importing {len(files)} {body_type} curves...")
    for i, fpath in enumerate(files):
        stn_id = i + 1
        abs_path = str(fpath.resolve())
        log.info(f"  Importing {fpath.name}...")

        success = model.InsertCurveFilePoint(abs_path)
        if not success:
            log.error(f"  InsertCurveFilePoint failed for {fpath.name}")
            features.append((stn_id, None))
            continue

        feat = get_last_feature(model)
        rename_feature(feat, f"{body_type}_S{stn_id:02d}")
        features.append((stn_id, feat))
        log.info(f"  Imported {fpath.name} -> {feat.Name if feat else '?'}")

    return features


# ─── Phase: Guide Curves ────────────────────────────────────

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


def create_guide_curves(model, curves_dir, body_type):
    """
    Create all guide curves as 3D sketch splines.
    Returns list of (guide_name, feature) tuples.
    """
    guide_pts = build_guide_points(curves_dir, body_type)
    features = []

    log.info(f"Creating {len(guide_pts)} {body_type} guide curves...")
    for name, pts in guide_pts.items():
        full_name = f"{body_type}_guide_{name}"
        log.info(f"  Creating {full_name} ({len(pts)} points)...")
        feat = create_3d_spline(model, pts)
        rename_feature(feat, full_name)
        features.append((name, feat))

    return features


# ─── Phase: Surfaces ────────────────────────────────────────

def create_surface(model, profile_features, guide_features, body_type):
    """
    Create a boundary surface (or loft fallback) from profiles and guides.
    profile_features: list of (stn_id, feat) from import_curves
    guide_features: list of (name, feat) from create_guide_curves
    """
    log.info(f"Creating {body_type} surface...")

    # Clear any existing selection
    model.ClearSelection2(True)

    # Select profiles with mark = 1
    profile_count = 0
    for stn_id, feat in profile_features:
        if feat is None:
            log.warning(f"  Skipping missing profile S{stn_id:02d}")
            continue
        if select_feature(model, feat, MARK_PROFILE, append=(profile_count > 0)):
            profile_count += 1
        else:
            log.warning(f"  Failed to select profile S{stn_id:02d}")

    # Select guides with mark = 2
    guide_count = 0
    for name, feat in guide_features:
        if feat is None:
            continue
        if select_feature(model, feat, MARK_GUIDE, append=True):
            guide_count += 1
        else:
            log.warning(f"  Failed to select guide {name}")

    log.info(f"  Selected {profile_count} profiles, {guide_count} guides")

    if profile_count < 2:
        log.error("  Need at least 2 profiles for surface creation")
        return None

    # Try boundary surface first
    feat_mgr = model.FeatureManager
    surface_feat = None

    try:
        log.info("  Attempting boundary surface...")
        # InsertBoundarySurface creates a surface from the current selection.
        # The profiles must be selected with Mark=1 and guides with Mark=2.
        surface_feat = feat_mgr.InsertBoundarySurface(
            False,  # bClosed - not a closed surface
            False,  # bMerge - don't merge with existing surfaces
        )
    except Exception as e:
        log.warning(f"  Boundary surface failed: {e}")

    if surface_feat is None:
        log.info("  Falling back to loft surface...")
        model.ClearSelection2(True)

        # Re-select profiles only for loft
        for stn_id, feat in profile_features:
            if feat is not None:
                select_feature(model, feat, MARK_PROFILE, append=True)

        try:
            surface_feat = feat_mgr.InsertProtrusionBlend2(
                False,  # Closed
                True,   # KeepTangent
                False,  # ForceNonRational
                1.0,    # TessToleranceFactor
                0, 0,   # StartMatchingType, EndMatchingType
                1,      # GuideTangentType
                True,   # MergeSmoothFaces
                0,      # NumGuides (not used in this overload)
                0.0, 0.0,  # StartTangentLength, EndTangentLength
                False,  # IsThinFeature
                0.0, 0.0,  # Thickness1, Thickness2
                0,      # ThinType
                True,   # UseFeatScope
                True,   # UseAutoSelect
                0, 0,   # StartTangentType, EndTangentType
            )
        except Exception as e:
            log.warning(f"  Loft also failed: {e}")

    if surface_feat is None:
        # Final fallback: simple loft without guide curves
        log.info("  Attempting simple loft (no guides)...")
        model.ClearSelection2(True)

        for stn_id, feat in profile_features:
            if feat is not None:
                select_feature(model, feat, MARK_PROFILE, append=True)

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
        feat = get_last_feature(model)
        rename_feature(feat, f"{body_type}_OML_Surface")
        log.info(f"  Created surface: {feat.Name if feat else body_type}")

    return surface_feat


# ─── Orchestrator ────────────────────────────────────────────

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
    set_units_ips(model)

    # Storage for features per body type
    all_profiles = {}
    all_guides = {}

    for bt in body_types:
        log.info(f"\n{'='*50}")
        log.info(f"Processing {bt.upper()} body")
        log.info(f"{'='*50}")

        # Phase: Curves
        if do_curves:
            all_profiles[bt] = import_curves(model, curves_dir, bt)

        # Phase: Guides
        if do_guides:
            all_guides[bt] = create_guide_curves(model, curves_dir, bt)

        # Phase: Surfaces
        if do_surfaces:
            profiles = all_profiles.get(bt, [])
            guides = all_guides.get(bt, [])
            if profiles:
                create_surface(model, profiles, guides, bt)
            else:
                log.warning(f"No profiles available for {bt} surface — skipping")

    # ── Final cleanup ──
    try:
        model.ForceRebuild3(False)
        model.ViewZoomtofit2()
    except Exception:
        pass

    if args.save:
        save_path = str(Path(args.save).resolve())
        log.info(f"Saving to {save_path}...")
        errors = model.Extension.SaveAs3(
            save_path, 0, 0, None, None, None
        )
        if errors == 0:
            log.info("Saved successfully")
        else:
            log.warning(f"Save returned error code: {errors}")

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
