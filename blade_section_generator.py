"""
S-64 Tail Rotor Blade Section Generator
========================================
Generates 3D point clouds for blade OML import into SolidWorks.

Features:
  - Linear taper with configurable start station
  - Linear twist with configurable start station
  - Sweep (X translation) with configurable start station
  - Trailing edge thickness enforcement (scale + truncate)
  - Nickel abrasion cap offset via outward unit normals
  - Multiple output formats: CSV per-station, combined, SolidWorks curve TXT

Coordinate system: X chordwise, Y thickness, Z spanwise
Airfoil point order: TE upper -> LE -> TE lower

Usage:
  python blade_section_generator.py                    # uses blade_config.yaml
  python blade_section_generator.py my_config.yaml     # custom config
  python blade_section_generator.py --plot              # show section plots
"""

import numpy as np
import yaml
import csv
import os
import sys
from pathlib import Path


# ─── Airfoil I/O ────────────────────────────────────────────

def load_airfoil(csv_path):
    """Load airfoil x/c, y/c from CSV. Returns (N,2) array."""
    pts = []
    with open(csv_path) as f:
        reader = csv.reader(f)
        header = next(reader)
        for row in reader:
            if len(row) >= 2:
                pts.append((float(row[0]), float(row[1])))
    return np.array(pts)


# ─── Geometry computations ───────────────────────────────────

def compute_chord(rR, cfg):
    """Compute chord at r/R using linear taper from taper_start_z."""
    g = cfg['global_geometry']
    span = g['span']
    z = rR * span
    taper_start = g['taper_start_z']
    if span <= taper_start or z <= taper_start:
        return g['root_chord']
    frac = (z - taper_start) / (span - taper_start)
    return g['root_chord'] + (g['tip_chord'] - g['root_chord']) * frac


def compute_twist(rR, cfg):
    """Compute twist in degrees at r/R."""
    g = cfg['global_geometry']
    span = g['span']
    z = rR * span
    twist_start = g['twist_start_z']
    if z <= twist_start or span == 0:
        return 0.0
    rate = g['twist_rate_deg_per_bs']  # deg per blade station (r/R)
    return rate * (z - twist_start) / span


def compute_sweep_x(rR, cfg):
    """Compute sweep X offset at r/R."""
    g = cfg['global_geometry']
    span = g['span']
    z = rR * span
    sweep_start = g['sweep_start_z']
    if z <= sweep_start:
        return 0.0
    return np.tan(np.radians(g['sweep_angle_deg'])) * (z - sweep_start)


def find_le_index(af_pts):
    """Find leading edge index (minimum x/c)."""
    return np.argmin(af_pts[:, 0])


def compute_normals(af_pts):
    """
    Compute outward unit normals at each airfoil point.

    Method (matching the Excel's AIRFOIL_NORMALS formula):
      - Forward difference tangent for all points except the last
        (last point uses backward difference)
      - Normal = perpendicular to tangent: (-dy, dx) / length
      - Outward direction determined by centroid sign test:
        if normal · (centroid - point) < 0, normal is already outward;
        otherwise flip it.
    """
    n = len(af_pts)
    normals = np.zeros_like(af_pts)

    # Centroid of all airfoil points
    cx = np.mean(af_pts[:, 0])
    cy = np.mean(af_pts[:, 1])

    for i in range(n):
        # Tangent: forward difference for all but last point
        if i < n - 1:
            dx = af_pts[i + 1, 0] - af_pts[i, 0]
            dy = af_pts[i + 1, 1] - af_pts[i, 1]
        else:
            dx = af_pts[i, 0] - af_pts[i - 1, 0]
            dy = af_pts[i, 1] - af_pts[i - 1, 1]

        length = np.sqrt(dx**2 + dy**2)
        if length < 1e-12:
            normals[i] = [0.0, 0.0]
            continue

        # Candidate normal: perpendicular to tangent
        nx, ny = -dy / length, dx / length

        # Centroid sign test: outward means pointing AWAY from centroid
        # dot(normal, centroid - point) should be negative for outward
        to_centroid_x = cx - af_pts[i, 0]
        to_centroid_y = cy - af_pts[i, 1]
        dot = nx * to_centroid_x + ny * to_centroid_y

        if dot >= 0:
            # Normal points toward centroid; flip to outward
            nx, ny = -nx, -ny

        normals[i] = [nx, ny]

    return normals


def apply_te_thickening(af_pts, chord, te_thickness_req):
    """
    Apply trailing edge thickness enforcement via scale + truncate.

    Algorithm:
      1. Compute current TE thickness = |y_upper(TE) - y_lower(TE)| * chord
      2. If current >= required, return airfoil unchanged.
      3. Compute te_ratio = te_req / chord.
      4. Build a thickness-vs-x/c distribution by pairing upper/lower surfaces.
      5. Search from TE inboard for u_cut where thickness(u_cut)/u_cut >= te_ratio.
      6. If found (u_cut < 1), scale x by 1/u_cut and truncate at x/c = 1.
      7. If no valid u_cut found, return airfoil unchanged.

    Returns modified (x/c, y/c) array.
    """
    if te_thickness_req <= 0 or chord <= 0:
        return af_pts.copy()

    le_idx = find_le_index(af_pts)
    te_thick_current = abs(af_pts[0, 1] - af_pts[-1, 1]) * chord

    if te_thick_current >= te_thickness_req:
        return af_pts.copy()

    te_ratio = te_thickness_req / chord

    # Build paired thickness distribution using upper/lower surface interpolation
    upper = af_pts[:le_idx + 1]  # TE -> LE (x/c decreasing)
    lower = af_pts[le_idx:]      # LE -> TE (x/c increasing)

    # Interpolate lower surface y/c as function of x/c
    lower_xc = lower[:, 0]
    lower_yc = lower[:, 1]

    # Build thickness ratio table matching Excel's AIRFOIL_THICKNESS:
    # ratio(i) = (y_upper(i) - y_lower_interp(x_upper(i))) / x_upper(i)
    # Then find smallest x/c where ratio >= te_ratio via interpolation.
    ratios = []
    xcs = []
    for i in range(len(upper)):
        xc = upper[i, 0]
        if xc <= 1e-9:
            continue
        y_up = upper[i, 1]
        j = np.searchsorted(lower_xc, xc)
        if j <= 0 or j >= len(lower_xc):
            continue
        frac = (xc - lower_xc[j - 1]) / (lower_xc[j] - lower_xc[j - 1] + 1e-15)
        y_lo = lower_yc[j - 1] + frac * (lower_yc[j] - lower_yc[j - 1])
        thickness = y_up - y_lo
        if thickness <= 0:
            continue
        ratios.append(thickness / xc)
        xcs.append(xc)

    # Search from TE inboard: find where ratio crosses te_ratio
    # Use linear interpolation between adjacent points for precise u_cut
    u_cut = None
    for k in range(len(xcs) - 1):
        # xcs[0] is closest to TE, decreasing toward LE
        if ratios[k] < te_ratio <= ratios[k + 1]:
            # Interpolate between k and k+1
            frac = (te_ratio - ratios[k]) / (ratios[k + 1] - ratios[k] + 1e-15)
            u_cut = xcs[k] + frac * (xcs[k + 1] - xcs[k])
            break
        elif ratios[k] >= te_ratio:
            u_cut = xcs[k]
            break

    if u_cut is None or u_cut >= 0.995:
        # No valid cut found — return airfoil unchanged
        return af_pts.copy()

    # Scale x coordinates and truncate
    new_pts = af_pts.copy()
    new_pts[:, 0] /= u_cut
    mask = new_pts[:, 0] <= 1.0 + 1e-9
    new_pts = new_pts[mask]
    new_pts[:, 0] = np.clip(new_pts[:, 0], 0.0, 1.0)
    return new_pts


def compute_cap_offset(af_pts, normals, cap_cfg, chord):
    """
    Compute nickel cap offset points.
    Forward of cap_te_xc: offset varies from offset_at_le (at LE) to
    offset_at_cap_te (at cap TE x/c) using power-law decay.
    Aft of cap_te_xc: zero offset (matches base).
    """
    cap_te_xc = cap_cfg['cap_te_xc']
    offset_le = cap_cfg['offset_at_le']
    offset_cap_te = cap_cfg['offset_at_cap_te']
    exponent = cap_cfg['exponent']
    sign = cap_cfg['normal_offset_sign']

    le_idx = find_le_index(af_pts)
    le_xc = af_pts[le_idx, 0]

    cap_pts = af_pts.copy()
    for i in range(len(af_pts)):
        xc = af_pts[i, 0]
        if xc > cap_te_xc:
            # Aft of cap TE: no offset
            continue

        # Normalized position: 0 at cap TE, 1 at LE
        if abs(cap_te_xc - le_xc) < 1e-12:
            t = 1.0
        else:
            t = (cap_te_xc - xc) / (cap_te_xc - le_xc)
        t = np.clip(t, 0.0, 1.0)

        # Power-law interpolation: offset ramps from cap_te value to LE value
        offset = offset_cap_te + (offset_le - offset_cap_te) * (t ** exponent)

        # Apply along outward normal, scaled by sign
        cap_pts[i, 0] += sign * normals[i, 0] * offset / chord
        cap_pts[i, 1] += sign * normals[i, 1] * offset / chord

    return cap_pts


def transform_section(af_xc_yc, chord, twist_deg, sweep_x, ref_axis_xc, z_pos):
    """
    Transform normalized airfoil (x/c, y/c) to 3D coordinates.
    1. Scale by chord
    2. Shift so ref axis is at X=0
    3. Rotate by twist about ref axis
    4. Apply sweep X offset
    5. Set Z position

    Returns (N, 3) array of [X, Y, Z].
    """
    n = len(af_xc_yc)
    pts_3d = np.zeros((n, 3))

    # Scale to physical coordinates
    x = af_xc_yc[:, 0] * chord
    y = af_xc_yc[:, 1] * chord

    # Shift so ref axis is at origin (for twist rotation)
    x_ref = ref_axis_xc * chord
    x_shifted = x - x_ref
    y_shifted = y  # ref axis is on the chord line (y=0)

    # Rotate by twist
    theta = np.radians(twist_deg)
    cos_t, sin_t = np.cos(theta), np.sin(theta)
    x_rot = x_shifted * cos_t - y_shifted * sin_t
    y_rot = x_shifted * sin_t + y_shifted * cos_t

    # Output with ref axis at X=0 (sections centered on pitch axis) + sweep
    pts_3d[:, 0] = x_rot + sweep_x
    pts_3d[:, 1] = y_rot
    pts_3d[:, 2] = z_pos

    return pts_3d


# ─── Output writers ──────────────────────────────────────────

def write_station_csv(filepath, pts_3d, station_id):
    """Write a single station's XYZ points to CSV."""
    with open(filepath, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['X_in', 'Y_in', 'Z_in'])
        for p in pts_3d:
            writer.writerow([f'{p[0]:.8f}', f'{p[1]:.8f}', f'{p[2]:.8f}'])


def write_sldcrv(filepath, pts_3d):
    """
    Write SolidWorks curve file (.sldcrv / .txt).
    Format: tab-separated X Y Z, one point per line.
    SolidWorks Import Curve expects this format.
    """
    with open(filepath, 'w') as f:
        for p in pts_3d:
            f.write(f'{p[0]:.8f}\t{p[1]:.8f}\t{p[2]:.8f}\n')


def write_combined_csv(filepath, all_stations):
    """Write all stations to a single CSV with station ID column."""
    with open(filepath, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Station', 'r/R', 'X_in', 'Y_in', 'Z_in'])
        for stn in all_stations:
            for p in stn['pts_3d']:
                writer.writerow([
                    stn['id'], f'{stn["rR"]:.6f}',
                    f'{p[0]:.8f}', f'{p[1]:.8f}', f'{p[2]:.8f}'
                ])


# ─── Visualization ───────────────────────────────────────────

def plot_sections(all_stations, all_stations_cap=None, save_dir=None):
    """Plot 2D sections and 3D blade overview using matplotlib."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from mpl_toolkits.mplot3d import Axes3D

    # 2D section overlay
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.set_title('Blade Sections (normalized x/c, y/c)')
    for stn in all_stations:
        af = stn['section_pts']
        ax.plot(af[:, 0], af[:, 1], '-', linewidth=0.8,
                label=f"S{stn['id']:02d} r/R={stn['rR']:.2f} c={stn['chord']:.3f}")
    ax.set_xlabel('x/c')
    ax.set_ylabel('y/c')
    ax.set_aspect('equal')
    ax.legend(fontsize=7, loc='upper right')
    ax.grid(True, alpha=0.3)
    if save_dir:
        fig.savefig(str(save_dir / 'sections_2d.png'), dpi=150, bbox_inches='tight')

    # 3D view
    fig3 = plt.figure(figsize=(14, 8))
    ax3 = fig3.add_subplot(111, projection='3d')
    for stn in all_stations:
        pts = stn['pts_3d']
        ax3.plot(pts[:, 2], pts[:, 0], pts[:, 1], 'b-', linewidth=0.8)
        if all_stations_cap:
            cap_stn = next((s for s in all_stations_cap if s['id'] == stn['id']), None)
            if cap_stn:
                cpts = cap_stn['pts_3d']
                ax3.plot(cpts[:, 2], cpts[:, 0], cpts[:, 1], 'r--', linewidth=0.5, alpha=0.6)
    ax3.set_xlabel('Z (span, in)')
    ax3.set_ylabel('X (chord, in)')
    ax3.set_zlabel('Y (thickness, in)')
    ax3.set_title('3D Blade OML — Blue: Base, Red: Nickel Cap')
    if save_dir:
        fig3.savefig(str(save_dir / 'blade_3d.png'), dpi=150, bbox_inches='tight')

    # Nickel cap detail for first and last station
    if all_stations_cap and len(all_stations) > 0:
        fig2, axes = plt.subplots(1, 2, figsize=(16, 6))
        for ax_idx, stn_idx in enumerate([0, -1]):
            ax2 = axes[ax_idx]
            stn = all_stations[stn_idx]
            cap_stn = all_stations_cap[stn_idx]
            base = stn['pts_3d']
            cap = cap_stn['pts_3d']
            ax2.plot(base[:, 0], base[:, 1], 'b-', linewidth=1.2, label='Base OML')
            ax2.plot(cap[:, 0], cap[:, 1], 'r--', linewidth=1.2, label='Nickel Cap')
            ax2.set_xlabel('X (in)')
            ax2.set_ylabel('Y (in)')
            ax2.set_title(f"S{stn['id']:02d} r/R={stn['rR']:.2f} — "
                          f"chord={stn['chord']:.3f}\", twist={stn['twist']:.1f}°")
            ax2.set_aspect('equal')
            ax2.legend(fontsize=9)
            ax2.grid(True, alpha=0.3)
        plt.tight_layout()
        if save_dir:
            fig2.savefig(str(save_dir / 'cap_detail.png'), dpi=150, bbox_inches='tight')

    if not save_dir:
        plt.show()
    else:
        plt.close('all')
        print(f"  Plots saved to: {save_dir}")


# ─── Main generator ──────────────────────────────────────────

def generate_blade(config_path, do_plot=False):
    """Main entry point: load config, compute all stations, write outputs."""

    config_dir = Path(config_path).parent
    with open(config_path) as f:
        cfg = yaml.safe_load(f)

    g = cfg['global_geometry']
    span = g['span']
    ref_xc = g['ref_axis_xc']
    te_cfg = cfg.get('trailing_edge', {})
    cap_cfg = cfg.get('nickel_cap', {})

    # Load airfoils
    airfoil_data = {}
    for name, csv_file in cfg['airfoils'].items():
        path = config_dir / csv_file
        airfoil_data[name] = load_airfoil(str(path))
        print(f"  Loaded airfoil {name}: {len(airfoil_data[name])} points")

    # Precompute normals per airfoil
    airfoil_normals = {}
    for name, pts in airfoil_data.items():
        airfoil_normals[name] = compute_normals(pts)

    # Process each station
    all_base = []
    all_cap = []
    output_dir = config_dir / 'output'
    output_dir.mkdir(exist_ok=True)
    curves_dir = output_dir / 'curves'
    curves_dir.mkdir(exist_ok=True)

    print(f"\n  Processing {len(cfg['stations'])} stations...")

    for i, stn_cfg in enumerate(cfg['stations']):
        stn_id = i + 1
        rR = stn_cfg['rR']
        z_pos = rR * span
        af_name = stn_cfg['airfoil']

        # Get base airfoil
        af_pts = airfoil_data[af_name].copy()

        # Compute distributions (with optional per-station overrides)
        chord = stn_cfg.get('chord', compute_chord(rR, cfg))
        twist = stn_cfg.get('twist_deg', compute_twist(rR, cfg))
        sweep = stn_cfg.get('sweep_x', compute_sweep_x(rR, cfg))

        # TE thickening
        te_req = stn_cfg.get('te_thickness', te_cfg.get('min_thickness', 0))
        if af_name == 'ROOT_CUTOUT':
            te_req = 0  # never apply to root cutout
        section_pts = apply_te_thickening(af_pts, chord, te_req)

        # Transform to 3D
        scaleY = chord  # default; could add override
        pts_3d_base = transform_section(section_pts, chord, twist, sweep, ref_xc, z_pos)

        stn_result = {
            'id': stn_id, 'rR': rR, 'z': z_pos,
            'chord': chord, 'twist': twist, 'sweep': sweep,
            'airfoil': af_name,
            'section_pts': section_pts,
            'pts_3d': pts_3d_base
        }
        all_base.append(stn_result)

        # Nickel cap offset
        normals = airfoil_normals[af_name]
        # If TE thickening changed the point count, recompute normals
        if len(section_pts) != len(normals):
            normals = compute_normals(section_pts)
        cap_section = compute_cap_offset(section_pts, normals, cap_cfg, chord)
        pts_3d_cap = transform_section(cap_section, chord, twist, sweep, ref_xc, z_pos)

        cap_result = {
            'id': stn_id, 'rR': rR, 'z': z_pos,
            'chord': chord, 'twist': twist, 'sweep': sweep,
            'airfoil': af_name,
            'section_pts': cap_section,
            'pts_3d': pts_3d_cap
        }
        all_cap.append(cap_result)

        # Write per-station files
        write_station_csv(output_dir / f'base_S{stn_id:02d}.csv', pts_3d_base, stn_id)
        write_station_csv(output_dir / f'cap_S{stn_id:02d}.csv', pts_3d_cap, stn_id)
        write_sldcrv(curves_dir / f'base_S{stn_id:02d}.txt', pts_3d_base)
        write_sldcrv(curves_dir / f'cap_S{stn_id:02d}.txt', pts_3d_cap)

        print(f"    S{stn_id:02d}: r/R={rR:.4f}, Z={z_pos:.4f}, "
              f"chord={chord:.4f}, twist={twist:.2f}°, "
              f"sweep={sweep:.4f}, {len(section_pts)} pts")

    # Combined output files
    write_combined_csv(output_dir / 'all_base_points.csv', all_base)
    write_combined_csv(output_dir / 'all_cap_points.csv', all_cap)

    # Summary
    print(f"\n  Output written to: {output_dir}")
    print(f"    Per-station CSVs:  base_S##.csv, cap_S##.csv")
    print(f"    SolidWorks curves: curves/base_S##.txt, curves/cap_S##.txt")
    print(f"    Combined:          all_base_points.csv, all_cap_points.csv")

    # Always save plots; show interactively if --plot
    plot_sections(all_base, all_cap, save_dir=output_dir)
    if do_plot:
        import matplotlib
        matplotlib.use('TkAgg')
        plot_sections(all_base, all_cap)

    return all_base, all_cap


# ─── CLI ─────────────────────────────────────────────────────

if __name__ == '__main__':
    config_file = 'blade_config.yaml'
    do_plot = False

    for arg in sys.argv[1:]:
        if arg == '--plot':
            do_plot = True
        elif arg.endswith(('.yaml', '.yml')):
            config_file = arg

    # Resolve relative to script directory if not absolute
    if not os.path.isabs(config_file):
        config_file = os.path.join(os.path.dirname(__file__) or '.', config_file)

    print(f"Blade Section Generator")
    print(f"  Config: {config_file}")
    generate_blade(config_file, do_plot)
