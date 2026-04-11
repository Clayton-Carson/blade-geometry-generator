# blade_geometry_generator — S-64 tail rotor blade OML

Python replacement for `tail_rotor_blade_section_generator.xlsx`.
Generates 3D point clouds defining the outer mold line (OML) of the
S-64 tail rotor blade, including nickel abrasion cap offset surfaces.
Output is formatted for direct SolidWorks import.

## Stack
- Python 3.11+
- Deps: `numpy`, `matplotlib`, `pyyaml`, `pytest`, and `pywin32`
  (Windows only, for SolidWorks COM automation in `solidworks_import.py`)
- Lint: `ruff check` / Format: `ruff format` / Test: `pytest`
- No venv committed — uses the per-user Python on PATH. The background
  ruff hook probes `./venv/` first then falls through to PATH.

## Conventions
- All geometry inputs live in `blade_config.yaml`. Airfoil coordinate
  CSVs at the repo root (`airfoil_RC310.csv`, `airfoil_RC410.csv`,
  `airfoil_ROOT_CUTOUT.csv`) are calibrated airfoil definitions.
- Point clouds and SolidWorks macro files go to `./output/` (created
  at run time, not committed).
- Reference spreadsheet `tail_rotor_blade_section_generator.xlsx` is
  the ground truth for regression comparison. Do not edit it.
- Units are inches throughout (legacy S-64 geometry). If you introduce
  any SI math, label it explicitly and convert at the boundary, not
  silently mid-pipeline.

## Hard rules
- Never loosen the regression tolerances without a logged engineering
  reason. The Python version has been verified bit-identical to the
  Excel for base OML and ~1e-13 inches for the cap — regressions of
  that magnitude need justification.
- Never run `solidworks_import.py` against a live SolidWorks session
  without first confirming the target file is closed or backed up.
  COM automation will overwrite geometry.
- Never modify `airfoil_*.csv` files — those are calibrated airfoil
  definitions, not raw data.
- When refactoring, preserve the existing generator's numeric output
  until the new path is bit-identical or explicitly signed off.

## Working style
- When I ask "what does X do", read `blade_section_generator.py` —
  don't guess from variable names.
- New analyses go in new scripts, not edits to the main generator.
- Plots go to `./output/` with a timestamp prefix, not the repo root.
