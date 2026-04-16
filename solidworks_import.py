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

# sldworks.tlb GUID — stable across SW versions; the major version follows SW:
# SW 2021 = 29, SW 2022 = 30, etc.
SLDWORKS_TLB_GUID = "{83A33D31-27C5-11CE-BFD4-00400513BB57}"


# DISPIDs and arg-type signatures extracted from the generated
# win32com.gen_py wrapper for sldworks.tlb. We can't use the generated class
# wrappers directly because SW objects don't respond to QueryInterface for
# their declared interfaces, so the wrapper's _oleobj_ ends up being the raw
# CDispatch without the correct DISPID routing. Calling InvokeTypes directly
# on the underlying PyIDispatch (._oleobj_) sidesteps that entirely.
#
# Each tuple: (dispid, lcid, wFlags, return_type_tuple, arg_types_tuple)
_DISPID_InsertNetBlend2 = (
    251, 0, 1, (9, 0),
    ((2, 1), (2, 1), (2, 1), (11, 1), (5, 1), (11, 1), (11, 1), (11, 1),
     (11, 1), (11, 1), (5, 1), (5, 1), (11, 1), (2, 1), (11, 1), (11, 1),
     (5, 1), (11, 1), (5, 1), (11, 1), (11, 1)),
)
# IModelDoc2.InsertLoftRefSurface2 — lofted surface from current selection
# (profiles mark=1, guides mark=2). Returns VT_VOID in the typelib.
_DISPID_InsertLoftRefSurface2 = (
    65882, 0, 1, (24, 0),
    ((11, 1), (11, 1), (11, 1), (5, 1), (2, 1), (2, 1)),
)
# IModelDocExtension.SelectByID2 — authoritative name-based selection.
_DISPID_SelectByID2 = (
    68, 0, 1, (11, 0),
    ((8, 1), (8, 1), (5, 1), (5, 1), (5, 1), (11, 1), (3, 1), (9, 1), (3, 1)),
)
_DISPID_SaveAs3 = (
    66222, 0, 1, (3, 0),
    ((8, 1), (3, 1), (3, 1)),
)


def _invoke(obj, dispid_spec, *args):
    """Call a method by DISPID on the raw PyIDispatch underlying obj."""
    dispid, lcid, flags, rettype, argtypes = dispid_spec
    return obj._oleobj_.InvokeTypes(dispid, lcid, flags, rettype, argtypes, *args)


def connect_solidworks():
    """Connect to running SolidWorks, or launch a new instance."""
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()

    sw = None
    try:
        sw = win32com.client.GetObject(Class="SldWorks.Application")
        log.info("Attached to running SolidWorks instance")
    except Exception as e:
        log.debug(f"GetObject attach failed: {e}")

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
            "Cannot connect to SolidWorks. Ensure it is installed and licensed."
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


def _reset_sketch_counter():
    """No-op retained for backwards compatibility — current flow no longer
    needs a counter since it grabs the just-created feature directly."""
    pass


def _newest_feature(model):
    """Return the most recently added feature via FeatureManager.GetFeatures.

    We can't use model.FirstFeature/GetNextFeature in late-bound dispatch
    ("Member not found"), and selection-based approaches are unreliable after
    closing a 3D sketch. FeatureManager.GetFeatures(False) returns the full
    feature list in tree order, so the last entry is the newest.
    """
    try:
        features = model.FeatureManager.GetFeatures(False)
    except Exception as e:
        log.warning(f"  GetFeatures failed: {e}")
        return None
    if features is None:
        return None
    try:
        n = len(features)
    except Exception:
        return None
    if n == 0:
        return None
    return features[n - 1]


def _feature_count(model):
    """Return the current feature-tree count, or 0 if it can't be read."""
    try:
        features = model.FeatureManager.GetFeatures(False)
        return len(features) if features is not None else 0
    except Exception:
        return 0


def create_3d_spline(sw, points_inches, sketch_name=None):
    """
    Create a 3D sketch containing a spline through the given points.
    points_inches: (N, 3) numpy array in inches.
    API expects meters, so we convert here.

    Returns the Feature IDispatch handle for the newly-created sketch.
    Downstream code should hold this handle and call feat.Select2(append, mark)
    directly — by-name lookup (FeatureByName / SelectByID2) is unreliable
    across many sketches in late-bound dispatch.
    """
    import win32com.client
    import pythoncom

    model = sw.ActiveDoc
    sketch_mgr = model.SketchManager

    # Clear selection BEFORE opening. If the previously-closed sketch is still
    # selected, Insert3DSketch re-enters edit on that sketch (combining two
    # profiles into one). Clearing ensures a brand-new sketch.
    model.ClearSelection2(True)
    sketch_mgr.Insert3DSketch(True)  # open

    pts_m = (points_inches * INCHES_TO_METERS).flatten().tolist()
    pt_array = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, pts_m
    )
    spline = sketch_mgr.CreateSpline2(pt_array, False)
    if spline is None:
        log.debug(f"  CreateSpline2 returned None for {len(points_inches)} points")

    sketch_mgr.Insert3DSketch(True)  # close

    feat = _newest_feature(model)
    if feat is None:
        log.warning(f"  Could not locate newly-created sketch (rename to {sketch_name} skipped)")
        return None

    if sketch_name:
        try:
            feat.Name = sketch_name
            log.debug(f"  Named sketch {sketch_name!r}")
        except Exception as e:
            log.warning(f"  Could not rename sketch to {sketch_name!r}: {e}")

    return feat


def select_by_name(model, name, feat_type="SKETCH", mark=0, append=True):
    """Select a feature by name via IModelDocExtension.SelectByID2 (DISPID 68).

    Early-bound name-based selection is more reliable than Feature.Select2 on
    a captured handle — captured Feature IDispatches go stale once more
    features are added to the tree (Select2 returns True but the selection
    doesn't actually register, as seen in SelectionManager readback).
    """
    if not name:
        log.warning("  Cannot select: empty name")
        return False
    try:
        ok = _invoke(
            model.Extension, _DISPID_SelectByID2,
            name, feat_type, 0.0, 0.0, 0.0,
            bool(append), int(mark), None, 0,
        )
        if ok:
            log.debug(f"  Selected {name!r} (mark={mark})")
        else:
            log.warning(f"  SelectByID2 returned False for {name!r}")
        return ok
    except Exception as e:
        log.warning(f"  SelectByID2 exception for {name!r}: {e}")
        return False


# ─── Phase: Import Curves ───────────────────────────────────

def import_curves(sw, curves_dir, body_type):
    """
    Import cross-section curves as 3D sketch splines.
    Returns list of (station_id, label, feature_handle) tuples.
    """
    prefix = body_type
    files = discover_curve_files(curves_dir, prefix)
    sketches = []

    log.info(f"Importing {len(files)} {body_type} curves as 3D sketch splines...")
    for i, fpath in enumerate(files):
        stn_id = i + 1
        name = f"{body_type}_section_{stn_id:02d}"
        log.info(f"  Importing {fpath.name} -> {name}...")

        pts = read_curve_points(fpath)
        if len(pts) == 0:
            log.error(f"  No points in {fpath.name}")
            sketches.append((stn_id, name, None))
            continue

        feat = create_3d_spline(sw, pts, sketch_name=name)
        sketches.append((stn_id, name, feat))
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
    Returns list of (guide_name, label, feature_handle) tuples.
    """
    guide_pts = build_guide_points(curves_dir, body_type)
    sketches = []

    log.info(f"Creating {len(guide_pts)} {body_type} guide curves...")
    for name, pts in guide_pts.items():
        sketch_name = f"{body_type}_guide_{name}"
        log.info(f"  Creating {sketch_name} ({len(pts)} points)...")
        feat = create_3d_spline(sw, pts, sketch_name=sketch_name)
        sketches.append((name, sketch_name, feat))

    return sketches


# ─── Phase: Surfaces ─────────────��──────────────────────────

def create_surface(sw, profile_sketches, guide_sketches, body_type):
    """
    Create a boundary surface (or loft fallback) from profiles and guides.
    profile_sketches: list of (stn_id, label, feature_handle) from import_curves
    guide_sketches: list of (name, label, feature_handle) from create_guide_curves
    """
    model = sw.ActiveDoc
    log.info(f"Creating {body_type} surface...")

    # Capture feature count BEFORE touching selection. GetFeatures between
    # Select2 calls and the loft invocation can disturb selection state.
    before = _feature_count(model)

    model.ClearSelection2(True)

    profile_count = 0
    for stn_id, label, feat in profile_sketches:
        if not label:
            log.warning(f"  Skipping profile S{stn_id:02d} with no name")
            continue
        if select_by_name(model, label, "SKETCH", mark=MARK_PROFILE, append=(profile_count > 0)):
            profile_count += 1

    guide_count = 0
    for _name, label, feat in guide_sketches:
        if not label:
            continue
        if select_by_name(model, label, "SKETCH", mark=MARK_GUIDE, append=True):
            guide_count += 1

    log.info(f"  Selected {profile_count} profiles, {guide_count} guides")

    # Diagnostic: read back the SelectionManager to verify marks are set
    try:
        sel_mgr = model.SelectionManager
        sel_count = sel_mgr.GetSelectedObjectCount2(-1)
        mark_hist = {}
        for i in range(1, sel_count + 1):
            m = sel_mgr.GetSelectedObjectMark(i)
            mark_hist[m] = mark_hist.get(m, 0) + 1
        log.info(f"  Selection readback: {sel_count} items; marks = {mark_hist}")
    except Exception as e:
        log.warning(f"  Selection readback failed: {e}")

    if profile_count < 2:
        log.error("  Need at least 2 profiles for surface creation")
        return None

    # IModelDoc2.InsertLoftRefSurface2 lofts a surface through the selected
    # profiles (mark=1) using the selected guide curves (mark=2). The typelib
    # says it returns VT_VOID, so we verify success by counting features
    # before/after — if a new feature appeared, the loft worked.
    try:
        _invoke(
            model, _DISPID_InsertLoftRefSurface2,
            False,  # Closed
            False,  # KeepTangency
            False,  # ForceNonRational
            1.0,    # TessToleranceFactor
            0,      # StartMatchingType (0 = Default/None)
            0,      # EndMatchingType
        )
    except Exception as e:
        log.error(f"  InsertLoftRefSurface2 raised: {e}")
        return None

    after = _feature_count(model)
    if after <= before:
        log.error(
            f"  Loft surface not created (feature count {before} -> {after}). "
            "Check for sketch errors or intersection issues. Sketches and "
            "selections are in place — try Insert > Surface > Loft manually."
        )
        return None

    surface_feat = _newest_feature(model)
    log.info(f"  Created {body_type} surface (feature count {before} -> {after})")
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
    _reset_sketch_counter()

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
        Path(save_path).parent.mkdir(parents=True, exist_ok=True)
        try:
            model.ClearSelection2(True)
        except Exception:
            pass
        # Remove any existing file so SW doesn't show an overwrite dialog
        try:
            Path(save_path).unlink(missing_ok=True)
        except Exception:
            pass
        log.info(f"Saving to {save_path}...")
        try:
            # Options=4 = swSaveAsOptions_Silent (suppress dialogs)
            result = _invoke(model, _DISPID_SaveAs3, save_path, 0, 4)
            if result == 0:
                log.info("Saved successfully")
            else:
                log.warning(f"SaveAs3 returned error code {result}")
        except Exception as e:
            log.warning(f"SaveAs3 failed: {e}")

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
