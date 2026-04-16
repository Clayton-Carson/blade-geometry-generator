# SolidWorks Import — WIP Notes

Session on 2026-04-16. Trying to get `solidworks_import.py` to run end-to-end
against a live SW 2021 session and produce a saved .sldprt with two boundary
surface bodies (base OML, nickel cap).

## Current state

**Working:**
- Attaches to a running SolidWorks 2021 instance via `GetObject`.
- Creates a new part with IPS units.
- Generates **8 base + 8 cap cross-section sketches** (one 3D sketch per
  section, verified — no more "two sections per sketch" bug).
- Generates **11 base + 11 cap guide curves** as 3D sketches (TE upper/lower,
  LE, 4 upper index-based, 4 lower x/c-interpolated).
- Sketches rename correctly in the feature tree:
  `base_section_01..08`, `cap_section_01..08`,
  `base_guide_TE_upper`, `base_guide_LE`, etc.
  (Verified by reading `feat.Name` after assignment.)

**Not working:**
- **Boundary / loft surface never builds.** `InsertLoftRefSurface2` (DISPID
  65882) and `InsertNetBlend2` (DISPID 251) both return without raising but
  create no feature. `SelectionManager.GetSelectedObjectCount2(-1)` readback
  shows only 1 or 0 items selected despite ~19 Select2 calls returning True.
- **SaveAs3** returns error code 1 (`swGenericSaveError`). This happens after
  the failed surface creation — may be a side effect of the document's state
  after the failed loft attempts, or an independent issue.

## Root cause of the surface issue

**Selection marks aren't actually being applied when using captured Feature
IDispatches + `Feature.Select2(append, mark)`.** Verified via
`SelectionManager.GetSelectedObjectMark()` readback:

- 19 calls to `feat.Select2(True, mark)` all return True.
- Readback afterwards shows 1 item with mark=1 (or 0 items after switching
  to name-based SelectByID2).

Hypothesis: the Feature handles returned by
`FeatureManager.GetFeatures(False)[-1]` at sketch-creation time go stale once
more features are added to the tree. `Select2` on a stale handle appears to
succeed but the selection doesn't actually register in the
SelectionManager, which means the subsequent loft/boundary call finds no
profiles and silently fails.

Switching to `IModelDocExtension.SelectByID2` (DISPID 68) by sketch name
returned False for every name — suggests the names assigned via
`feat.Name = new_name` aren't visible to `SelectByID2`'s name resolution
either, even though `FeatureByName()` returned the renamed feature in
earlier tests. There is some staleness / name-index disagreement between
the different name-lookup paths in late-bound SW COM.

## COM landscape — what I learned the hard way

SolidWorks 2021's COM objects **do not expose typelib info via
`IDispatch::GetTypeInfo`** (it returns `(−2147352565, 'Invalid index.')`),
so pywin32's automatic early-bound wrapping fails. Consequences:

- Many documented methods are **not resolvable by name** via late-bound
  `GetIDsOfNames`: `InsertBoundarySurface2`, `InsertNetBlend2`, `SaveAs3`,
  `SelectByID2` all raise "Member not found" or "unknown" errors when called
  by attribute access on the CDispatch.
- `win32com.client.CastTo()` fails for the same reason
  (QueryInterface on the typelib IID is rejected).
- `model.FirstFeature()` raises `Member not found` — this method isn't in
  the late-bound dispatch table either.

**The workaround that works**: manually load `sldworks.tlb` via
`gencache.EnsureModule` to populate the generated wrapper module, scrape
DISPIDs and arg-type tuples out of it, and call
`obj._oleobj_.InvokeTypes(dispid, lcid, flags, rettype, argtypes, *args)`
on the raw PyIDispatch. This bypasses name lookup entirely and lets the COM
server route by DISPID, which it respects.

DISPIDs currently hardcoded in `solidworks_import.py`:
- `IFeatureManager.InsertNetBlend2` = 251
- `IModelDoc2.InsertLoftRefSurface2` = 65882
- `IModelDocExtension.SelectByID2` = 68
- `IModelDoc2.SaveAs3` = 66222

(Extracted from the generated wrapper at
`%LOCALAPPDATA%\Temp\gen_py\3.12\83A33D31-27C5-11CE-BFD4-00400513BB57x0x29x0.py`.)

## Other findings worth keeping

- `model.Insert3DSketch2(True)` **toggles**, so it can re-edit a
  previously-closed sketch if that sketch is still selected — this was the
  cause of "two profiles landing in one sketch". Fix: `ClearSelection2(True)`
  before each `SketchManager.Insert3DSketch(True)` call.
- `CreateSpline2` occasionally returns `None` on the cap profiles but the
  spline is nonetheless created in the sketch. Not fatal.
- `model.SaveAs(path)` (1-arg) returns False in states where SaveAs3 returns
  non-zero — not a workaround for the save issue.

## What to try next

1. **Test whether the sketches are structurally valid for loft** by manually
   selecting them in SW UI and running Insert → Surface → Loft. If that works,
   the API-side selection is the only problem.
2. **Try `SelectByID2` by sketch name** — but first prove the names are
   indexed. Possibly need a `ForceRebuild3(False)` + small delay between
   `feat.Name = ...` and subsequent `SelectByID2` calls so SW commits the
   rename to its name-resolution index.
3. **Write the surface-creation step as a .swp/.swb macro** that the user
   runs from within SolidWorks. Python creates all sketches; the macro
   (early-bound VBA) handles the Boundary Surface step where late-bound
   Python COM is unreliable.
4. **Try `gencache.EnsureDispatch("SldWorks.Application")`** to launch a
   fresh SW instance rather than attaching to a running one — the fresh
   instance may advertise typelib info correctly and unlock early-bound
   dispatch across the board. (Downside: user has to close their current
   session.)

## Files

- `solidworks_import.py` — the main driver; currently wired to call
  `SelectByID2` and `InsertLoftRefSurface2` via DISPID. Selection diagnostic
  logging is in place.
- `blade_section_generator.py` — untouched; produces the curve .txt files
  that `solidworks_import.py` consumes. Working correctly.
- `output/curves/*.txt` — 16 curve files (8 base + 8 cap), each a tab-
  separated XYZ point list. These are the ground-truth inputs.
