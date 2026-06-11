# waters-raw-sigmoid

Automate the MassLynx → Excel sigmoid workflow used by collaborators on Waters SYNAPT-class data.

## What this replaces

For each Waters `.raw` measurement (one per collision energy), the manual workflow was:

1. Open the file in MassLynx, integrate intensity in an m/z window by hand.
2. Copy the exported `(m/z, intensity)` pairs into a per-analyte Excel sheet.
3. Let Excel compute `fraction = sum(intensity in window) / sum(all intensity)` via `MATCH` + `SUM` formulas (cell `C6 = C5 / C8` in the template).
4. Read the resulting CE → fraction table off the "Summary Mass Intensity" tab and fit a sigmoid.

This script does steps 1–4 in one command:

```sh
python extract_intensities.py <folder-of-raw-dirs> --mass-lo <lo> --mass-hi <hi>
```

It walks the folder, opens every `*.raw`, integrates the same window across every scan with [`openwraw`](https://pypi.org/project/openwraw/), and writes a `summary.tsv` plus a copy-pasteable `summary.txt` with `CE\tfraction` lines.

## Setup (Windows)

One-time, copy-paste:

```
python -m pip install --user -r requirements.txt
```

That installs `openwraw`. No DLLs, no MassLynx SDK, no `regsvr32`, no 32-bit Python.

## Run

```
python extract_intensities.py Messdaten --mass-lo 1369 --mass-hi 1373.2
```

`Messdaten` is whatever folder contains the `.raw` directories. Filenames need a `CE_<number>` somewhere (e.g. `MR_010626_DM_BCD_SODIUM_MSMS_CE_10_01.raw` → CE=10). `10`, `10.5`, `CE 10`, `CE-10` all work.

Each run creates a `run-<timestamp>/` folder next to the script with:

- `run.log` — full timestamped log (send this back if anything looks wrong)
- `summary.tsv` — `ce, fraction, windowed, total, n_scans, n_peaks, func, filename`
- `summary.txt` — minimal `CE\tfraction` block to paste into Excel/Origin/etc.

## What the script computes

For each `.raw`:

1. Open with `openwraw.RawReader(path)`.
2. Pick the first non-lock-mass MS function (override with `--func N`).
3. For every scan, sum intensity into two accumulators:
   - `total` — every peak
   - `windowed` — only peaks where `mass_lo ≤ mz ≤ mass_hi`
4. Output `fraction = windowed / total`.

Identical to Excel cell `C6 = C5 / C8`. No baseline subtraction, smoothing, or peak picking — same as the manual workflow.

## The chemistry behind the number

A SYNAPT-style MS/MS experiment selects a precursor ion in the quadrupole, pushes it through a collision cell at a chosen **collision energy (CE)**, and records the resulting m/z spectrum at the TOF analyser. As CE rises, the precursor breaks apart into fragment ions, so the **fraction of the population that is still the precursor decreases monotonically** with CE — a sigmoid (or, for protein unfolding/CIU, a series of sigmoids).

The quantity to fit is exactly that fraction. For a given m/z window `[mass_lo, mass_hi]` chosen to bracket one species (precursor or a specific fragment / conformer), and for one `.raw` file acquired at one CE:

```
          Σ   I(m/z, scan)
         m/z∈[lo, hi], scan
fraction = ──────────────────────────────────────────
             Σ   I(m/z, scan)
            all m/z, scan
```

The denominator is the **total ion current (TIC)** integrated over the run — the same number MassLynx shows in its "Combine" view at the top right of the spectrum panel. The numerator is the **integrated intensity inside the m/z window**, the value MassLynx returns when you click-drag a window in Spectrum view and read off the displayed area-under-curve.

Why summing across scans (our approach) is equivalent to MassLynx's Combine-then-integrate (the manual approach):

- MassLynx's "Combine all scans" produces a single averaged/summed spectrum: for every m/z bin, the value is the sum (or mean, depending on the option) of intensities at that bin across all 58 scans.
- Integrating a window on that combined spectrum is `Σ_{m/z∈window} Σ_{scan} I(m/z, scan)`.
- That's the same as `Σ_{scan} Σ_{m/z∈window} I(m/z, scan)`, which is what we compute.
- Addition is associative; the order of summation is irrelevant.

So the only physical assumption baked into the script is the **same one made by hand**: a flat baseline (we sum all peak intensities as-reported by the reader; we do not subtract a chemical-noise baseline). MassLynx's display sometimes shows a baseline-subtracted view, but the integration tool returns area-under-curve from the raw peak list, matching what we do.

### How to choose the window

One `.raw` file = one CE = one row in the final CE→fraction table. The `--mass-lo / --mass-hi` window is **per-analyte**, not per-CE. Same window for all CE values of the same molecule — that's what makes the resulting fractions a sigmoid in CE rather than a moving target.

For a precursor-loss sigmoid, choose a window tight around the intact precursor m/z (e.g. ±2 Da around the [M+Na]⁺ peak). For a fragment-formation sigmoid, the window brackets a specific fragment instead, and the resulting curve is the inverted sigmoid. Either way, the script just integrates whatever window you give it.

## Reader: why openwraw?

Waters' `.raw` is a directory of binary files (`_FUNC*.DAT/IDX`, `_FUNCTNS.INF`, `_HEADER.TXT`, `_extern.inf`, …). Waters has never published a spec. Readers either link Waters DLLs (Windows-only, license-bound) or reverse-engineer the layout from public PRIDE datasets.

| Tool | Approach | Platform | Why not used here |
|---|---|---|---|
| **openwraw** | Rust core + PyO3 bindings, format reverse-engineered from PRIDE; pure file IO | Linux/macOS/Windows | **chosen** |
| rainbow-api | Pure-Python, also reverse-engineered | All | works for MS, IMS support less documented |
| multiplierz / mzAPI | Wraps Waters `MassLynxRaw.dll` | Windows + DLL | recent MassLynx 4.2 installs no longer ship `MassLynxRaw.dll` |
| masslynx_sdk_public | Python port of official Waters SDK 4.7 | Windows + DLLs | needs Waters registration + admin approval |
| Waters DACServer.dll (COM) | Free w/ MassLynx, accessed via `win32com` | Windows + 32-bit Python | DACServer is registered under WOW6432Node (32-bit COM); requires a separate 32-bit Python install on a 64-bit machine |
| Waters Databridge | Official `.raw → mzML` GUI converter | Windows GUI | manual, defeats the automation |
| pyopenms | None — Waters not supported | — | confirmed failing in `_test_reader*.py` |

`openwraw`'s killer feature: same `pip install` works on macOS for development and on Windows in production. No platform branching, no native deps, identical behavior.

## Validation

The reference Excel template (kept locally in `data/`, not tracked) carries only the formula scaffold per analyte sheet — no pasted data, no calculated values to bit-compare against. The math itself (`windowed / total`) is verifiably what the template does (cell `C6 = C5 / C8`), so the only open question is whether `openwraw` returns the same per-scan peak data that MassLynx exports.

Cheap sanity checks before running against a real dataset:

- Open one `.raw` in MassLynx, combine all scans, read the displayed TIC. It should equal our `total`.
- Run the script on a folder spanning a CE range. The fractions should trace a sigmoid (chemistry-dependent — that's an experiment-design question).
- If you can get a previously filled-out copy of the spreadsheet, drop the same `.raw` files in a folder and check that our `fraction` matches the value in row 6 of the corresponding block.

## Layout

```
extract_intensities.py    # the script
requirements.txt          # openwraw
data/                     # gitignored — local sample .raw + reference xlsx
run-<timestamp>/          # gitignored — produced by each run
```

## Credits

This project is essentially a thin orchestration layer around [**OpenWRaw**](https://github.com/Sigilweaver/OpenWRaw) by [@Sigilweaver](https://github.com/Sigilweaver). Without it we would be stuck on a Windows machine with a 32-bit Python install talking to MassLynx's `DACServer.dll` over COM, or waiting on a Waters SDK registration. Instead the entire reader half of this repo is one `pip install`.

OpenWRaw is part of the [**OpenProteo**](https://sigilweaver.app/openproteo/docs/) stack of clean-room, cross-platform readers for proprietary mass-spectrometry vendor formats. Sibling readers cover the other two majors:

- [OpenWRaw](https://github.com/Sigilweaver/OpenWRaw) — Waters MassLynx `.raw`
- [OpenTFRaw](https://github.com/Sigilweaver/OpenTFRaw) — Thermo Fisher Xcalibur `.raw`
- [OpenTimsTDF](https://github.com/Sigilweaver/OpenTimsTDF) — Bruker timsTOF `.d / .tdf`

All three are Apache-2.0 licensed, derived from binary analysis of public PRIDE datasets, and ship as Rust crates plus PyO3/maturin Python wheels. They publish their format specifications as docs in-repo — something the proprietary-vendor end of the proteomics tooling world has been missing for two decades. If you do anything else with raw vendor MS data, take a look.
