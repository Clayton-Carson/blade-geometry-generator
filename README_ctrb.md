# S-64 Tail Rotor Blade Section Generator

Python tool for generating 3D point clouds defining the outer mold line (OML) of the S-64 tail rotor blade, including nickel abrasion cap offset surfaces. Output is formatted for direct import into SolidWorks.

## Background

This tool replaces `tail_rotor_blade_section_generator.xlsx`, which attempted the same computation using ~235,000 Excel formula cells. The Excel approach had several issues: formulas exceeding 500 characters that were essentially undebuggable, a `#VALUE!` error in the nickel cap output at the leading edge of every station, stale cached values in the TE thickening logic (the AGGREGATE function wasn't evaluating properly), and painful recalculation times due to a fixed 50-station x 500-point grid — most of which sat empty. The Python replacement is ~400 lines, runs in under a second, and has been verified to match the Excel's base OML output to machine precision (0.00 error) and the cap output to ~1e-13 inches.

## Project Files

```
ctrb/
├── blade_section_generator.py    # Main generator script
├── blade_config.yaml             # All geometry inputs (edit this)
├── airfoil_RC310.csv             # RC-310 airfoil coordinates (x/c, y/c)
├── airfoil_RC410.csv             # RC-410 airfoil coordinates
├── airfoil_ROOT_CUTOUT.csv       # Root cutout shape coordinates
├── tail_rotor_blade_section_generator.xlsx  # Original Excel (reference only)
├── README.md                     # This file
└── output/                       # Generated output (created on run)
    ├── base_S01.csv ... base_S08.csv      # Base OML per station (X,Y,Z inches)
    ├── cap_S01.csv  ... cap_S08.csv       # Nickel cap OML per station
    ├── all_base_points.csv                # All stations combined
    ├── all_cap_points.csv                 # All cap stations combined
    ├── sections_2d.png                    # Airfoil section overlay plot
    ├── blade_3d.png                       # 3D blade visualization
    ├── cap_detail.png                     # Base vs cap comparison plot
    └── curves/
        ├── base_S01.txt ... base_S08.txt  # SolidWorks curve files (tab-sep XYZ)
        └── cap_S01.txt  ... cap_S08.txt
```

## Requirements

Python 3.8+ with the following packages:

- `numpy` (core math)
- `pyyaml` (config parsing)
- `matplotlib` (visualization — plots are always saved; interactive display is optional)

Install: `pip install numpy pyyaml matplotlib`

## Usage

```bash
# Basic run — reads blade_config.yaml, writes output/
python blade_section_generator.py

# Custom config file
python blade_section_generator.py my_config.yaml

# Interactive plot display (in addition to saved PNGs)
python blade_section_generator.py --plot
```

## Configuration (blade_config.yaml)

All geometry is defined in a single YAML file. Key sections:

### Global Geometry (inches, degrees)

| Parameter | Description | Current Value |
|-----------|-------------|---------------|
| `span` | Blade span | 39.370" (1000 mm) |
| `root_chord` | Chord at root | 4.724" (120 mm) |
| `tip_chord` | Chord at tip | 3.150" (80 mm) |
| `ref_axis_xc` | Reference axis as x/c (pitch axis) | 0.25 (quarter-chord) |
| `twist_start_z` | Z position where twist begins | 0.0" |
| `twist_rate_deg_per_bs` | Twist rate per blade station (r/R) | -10.0 deg |
| `sweep_angle_deg` | Sweep angle | 0.0 deg |
| `taper_start_z` | Z position where taper begins | 0.0" |
| `sweep_start_z` | Z position where sweep begins | 0.0" |

### Nickel Cap Parameters

| Parameter | Description | Current Value |
|-----------|-------------|---------------|
| `cap_te_xc` | Cap trailing edge in x/c | 0.225 |
| `offset_at_le` | Max offset at leading edge | 0.050" |
| `offset_at_cap_te` | Offset at cap trailing edge | 0.011" |
| `exponent` | Power-law decay exponent | 3.333 (10/3) |
| `normal_offset_sign` | +1 outward, -1 inward | -1 (inward) |

The cap offset interpolation uses a power law: offset ramps from `offset_at_cap_te` at x/c = `cap_te_xc` up to `offset_at_le` at the leading edge. The exponent controls how steeply it ramps — higher values concentrate the offset more tightly around the LE. Aft of `cap_te_xc`, there is zero offset (cap matches base OML).

### Trailing Edge Thickness Control

Set `min_thickness` under `trailing_edge:` to enforce a minimum TE thickness. The algorithm finds a cut point `u_cut` where the airfoil's thickness-to-x ratio matches the requirement, scales the x-coordinates by `1/u_cut`, and truncates — keeping the chord unchanged while thickening the TE.

Set to `0` to disable. Per-station overrides are supported via `te_thickness` in individual station definitions.

Note: With the current RC310 airfoil at these chord values, the TE thickening will reduce the point count from 66 to ~62-64 per station. The original Excel had a bug where the AGGREGATE function didn't evaluate correctly, so it never actually applied the thickening despite the formulas being present.

### Station Definitions

```yaml
stations:
  - { rR: 0.000, airfoil: RC310 }                          # root
  - { rR: 0.500, airfoil: RC410, chord: 4.0, twist_deg: -3.0 }  # override example
  - { rR: 1.000, airfoil: RC310 }                          # tip
```

Each station requires `rR` (radial position as fraction of span) and `airfoil` (name matching an entry in the `airfoils:` section). Optional overrides: `chord`, `twist_deg`, `sweep_x`, `te_thickness`.

### Adding a New Airfoil

1. Create a CSV file with columns `x/c,y/c`
2. Points must be ordered: **TE upper surface → LE → TE lower surface** (closed loop, starting and ending near x/c = 1.0)
3. Add the file reference to `blade_config.yaml` under `airfoils:`
4. Reference the airfoil name in your station definitions

## Coordinate System

- **X**: Chordwise (positive aft from leading edge)
- **Y**: Thickness (positive up)
- **Z**: Spanwise (positive outboard, root = 0)

Output sections are centered on the pitch axis (reference axis at X = 0). Twist rotates in the X-Y plane about this axis.

## SolidWorks Import

The `curves/` directory contains tab-separated XYZ files ready for SolidWorks:

1. **Insert → Curve → Through XYZ Points**
2. Click **Browse** and select a `base_S##.txt` or `cap_S##.txt` file
3. Repeat for each station
4. Use **Insert → Surface → Loft** to loft between the imported curves

For the nickel cap, import the `cap_S##.txt` curves and loft separately to define the cap IML (inner mold line, since `normal_offset_sign = -1`).

## Computation Pipeline

```
blade_config.yaml
    │
    ├── Airfoil CSVs (x/c, y/c)
    │       │
    │       ├── [Optional] TE Thickening (scale + truncate)
    │       │
    │       ├── Compute outward unit normals (forward-difference tangent, centroid sign test)
    │       │
    │       └── Compute cap offset (power-law interpolation along normals)
    │
    └── Per-station parameters (chord, twist, sweep from distributions or overrides)
            │
            ├── Scale by chord
            ├── Shift to ref axis at X=0
            ├── Rotate by twist
            ├── Apply sweep X-offset
            └── Set Z = r/R × span
                    │
                    ├── output/base_S##.csv    (base OML)
                    ├── output/cap_S##.csv     (nickel cap OML)
                    └── output/curves/*.txt    (SolidWorks format)
```

## Known Considerations

- **TE thickening behavior**: The algorithm correctly finds `u_cut` via interpolation, but for airfoils with very sparse point density near the trailing edge (like RC310 with only 66 points), the truncation can remove a few points. If you need the raw airfoil without TE modification, set `min_thickness: 0`.

- **Leading edge normal**: At the exact LE point (x/c = 0), the forward-difference tangent can be zero if two consecutive points share the same coordinates. The generator outputs a zero normal there. This matches the Excel's `#VALUE!` at the same point. In practice this affects one point per station and has no impact on the lofted surface.

- **ROOT_CUTOUT airfoil**: TE thickening is automatically skipped for any station using this airfoil. The root cutout is defined parametrically in the Excel (chord basis = 24", thickness = 4", TE radius = 1") and the exported CSV preserves those exact coordinates.

## Verification

The Python output was verified against the Excel's cached values:

| Surface | Max Error vs Excel | Notes |
|---------|-------------------|-------|
| Base OML | 0.00 inches | Perfect match across all 8 stations |
| Nickel Cap | ~1e-13 inches | Machine precision; Excel had #VALUE! at LE |

## SolidWorks Import Automation

`solidworks_import.py` automates the import of generated blade curves into SolidWorks 2021 via the COM API. It creates cross-section curves, spanwise guide curves, and boundary surfaces — replacing the manual Insert → Curve → Through XYZ Points workflow.

### Requirements

- SolidWorks 2021 (installed and licensed)
- `pywin32`: `pip install pywin32`
- `numpy` (already required by the section generator)
- Output curves from `blade_section_generator.py` in `output/curves/`

### Usage

```bash
# Validate files without touching SolidWorks
python solidworks_import.py --dry-run

# Just import the 16 curve files
python solidworks_import.py --phase curves

# Import curves + create spanwise guide curves
python solidworks_import.py --phase guides

# Full pipeline: curves + guides + boundary surfaces
python solidworks_import.py

# Process only base OML or only nickel cap
python solidworks_import.py --body base
python solidworks_import.py --body cap

# Save the part
python solidworks_import.py --save output/blade.sldprt

# Verbose logging
python solidworks_import.py -v
```

### What It Creates

The script produces a single `.sldprt` with two surface bodies (Base OML and Nickel Cap). For each body:

**Cross-section curves (8 per body)**
- Imported via `InsertCurveFilePoint` from the `output/curves/` directory
- Named `base_S01`..`base_S08` and `cap_S01`..`cap_S08`

**Spanwise guide curves (11 per body)**

| Guide | Method | Description |
|-------|--------|-------------|
| TE Upper | Index 0 on all stations | Trailing edge, upper side |
| TE Lower | Last index on all stations | Trailing edge, lower side |
| LE | Min-X point per station | Leading edge (handles 64→62 point shift) |
| Upper ×4 | Indices 6, 12, 18, 24 | Upper surface intermediates |
| Lower ×4 | x/c = 0.15, 0.35, 0.55, 0.75 | Lower surface intermediates (x/c-interpolated) |

The lower surface guides use x/c-based geometric interpolation rather than index matching because TE thickening changes the point count from 64 (S01–S04) to 62 (S05–S08), shifting the LE index from 31 to 30 and breaking direct index correspondence on the lower surface.

**Boundary surface**
- Created from the 8 profile curves and 11 guide curves
- Falls back to loft surface if boundary creation fails

### Recommended Workflow

Start with `--dry-run` to verify file discovery, then step through phases incrementally:

1. `--phase curves` — confirm 16 curves appear in the feature tree
2. `--phase guides` — confirm guide curves connect corresponding points spanwise
3. Full run (no `--phase` flag) — confirm smooth boundary surfaces
4. `--save` to write the `.sldprt`

### Troubleshooting

- **"Cannot connect to SolidWorks"**: Ensure SolidWorks 2021 is installed. The script tries `GetObject` (attach to running instance) then `Dispatch` (launch new instance).
- **Boundary surface fails**: The script falls back to loft surface automatically. If both fail, import curves and guides, then create the surface manually via Insert → Surface → Boundary Surface.
- **Unit mismatch**: The script sets document units to IPS (inches). Curve files are imported in document units; guide curve splines are created via the API in meters (converted internally).

## Version History

- **v1**: Original Excel workbook (`tail_rotor_blade_section_generator.xlsx`) — 235K formula cells, 8.4 MB
- **v2**: Python section generator — ~400 lines, < 1 second runtime, verified against Excel
- **v3** (current): Added SolidWorks COM automation (`solidworks_import.py`)
