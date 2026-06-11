"""
extract_intensities.py — automate Maik's MassLynx → Excel sigmoid workflow.

Manual workflow being replaced
==============================
    1. Open each *.raw in MassLynx, integrate intensity inside an m/z window.
    2. Copy the exported (m/z, intensity) pairs into a per-analyte Excel sheet.
    3. Excel's MATCH + SUM formulas compute, per CE:
           fraction = sum(intensity where mass_lo <= mz <= mass_hi)
                    / sum(all intensity)
    4. Excel's "Summary Mass Intensity" tab pulls those fractions across all
       CE values and yields the CE -> fraction table that gets sigmoid-fit.

This script does steps 1-4 directly: walk a folder of *.raw directories,
read every scan with openwraw, integrate the same window in pure Python,
and emit one TSV with the same CE -> fraction table.

Usage
=====
    python extract_intensities.py <folder> --mass-lo 1369 --mass-hi 1373.2

Output
======
    A new directory `run-YYYYMMDD-HHMMSS/` next to the script containing:
        run.log       — full DEBUG log (timestamps, every scan, preflight)
        summary.tsv   — machine-readable: ce, fraction, windowed, total, ...
        summary.txt   — human-readable copy-pasteable into Excel/plotting

The console shows the same INFO-level summary; the *full* per-scan trace
only goes to run.log so the terminal stays readable.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import logging
import os
import platform
import re
import sys
import traceback
from pathlib import Path

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

log = logging.getLogger("sigmoid")


def setup_logging(out_dir: Path, console_level: int) -> Path:
    """Configure root logger: console + run.log file under out_dir."""
    out_dir.mkdir(parents=True, exist_ok=True)
    log_path = out_dir / "run.log"

    fmt_console = logging.Formatter("%(message)s")
    fmt_file = logging.Formatter(
        "%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    log.setLevel(logging.DEBUG)
    # Wipe any pre-existing handlers (e.g. when re-run interactively)
    for h in list(log.handlers):
        log.removeHandler(h)

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(console_level)
    ch.setFormatter(fmt_console)
    log.addHandler(ch)

    fh = logging.FileHandler(log_path, mode="w", encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt_file)
    log.addHandler(fh)

    return log_path


# ---------------------------------------------------------------------------
# Filename parsing
# ---------------------------------------------------------------------------

# Matches "CE_10", "CE 10", "CE10", "CE_10.5", "CE-10" — case-insensitive.
_CE_RE = re.compile(r"CE[_\s\-]*(\d+(?:\.\d+)?)", re.IGNORECASE)


def extract_ce(filename: str) -> float | None:
    m = _CE_RE.search(filename)
    return float(m.group(1)) if m else None


# ---------------------------------------------------------------------------
# Per-.raw integration
# ---------------------------------------------------------------------------

def pick_ms_function(reader, requested: int | None):
    """
    Pick the MS function to integrate over.

    If `requested` is given, use that 1-based index. Otherwise pick the
    first non-lock-mass function — that matches what MassLynx shows by
    default and what Maik would integrate by hand.
    """
    funcs = list(reader.functions)
    if not funcs:
        raise RuntimeError("No acquisition functions in this .raw")

    if requested is not None:
        for f in funcs:
            if f.index == requested:
                return f
        raise RuntimeError(
            f"Function {requested} not found. Available: "
            + ", ".join(str(f.index) for f in funcs)
        )

    for f in funcs:
        if not f.is_lock_mass:
            return f
    return funcs[0]


def integrate_raw(reader, func, mass_lo: float, mass_hi: float):
    """
    Sum intensity over every scan in `func`. Returns:
        windowed   — sum of intensities where mass_lo <= mz <= mass_hi
        total      — sum of all intensities
        n_scans    — scans iterated
        n_peaks    — total peak entries seen across all scans
        n_skipped  — scans that returned 0 peaks (for log diagnostics)
    """
    n_scans = reader.n_scans(func.index)
    log.info("  function %d: %d scans, acquisition m/z [%.2f, %.2f]",
             func.index, n_scans, func.mz_low, func.mz_high)

    windowed = 0.0
    total = 0.0
    n_peaks = 0
    n_skipped = 0
    n_failed = 0

    for s in range(n_scans):
        try:
            spec = reader.read_spectrum(func.index, s)
        except Exception as exc:  # noqa: BLE001
            n_failed += 1
            log.warning("  scan %d: read failed: %s", s, exc)
            continue

        if len(spec) == 0:
            n_skipped += 1
            continue

        # spec.mz / spec.intensity are list[float] from openwraw.
        for mz, inten in zip(spec.mz, spec.intensity):
            total += inten
            if mass_lo <= mz <= mass_hi:
                windowed += inten
        n_peaks += len(spec)

    if n_skipped:
        log.info("  %d/%d scans empty (skipped)", n_skipped, n_scans)
    if n_failed:
        log.warning("  %d/%d scans failed to read", n_failed, n_scans)

    return windowed, total, n_scans, n_peaks, n_skipped


# ---------------------------------------------------------------------------
# Preflight (everything we want to see in run.log when something is wrong)
# ---------------------------------------------------------------------------

def preflight(args, raw_dirs):
    log.info("=" * 70)
    log.info("waters-raw-sigmoid — extract_intensities.py")
    log.info("=" * 70)
    log.info("started:        %s", _dt.datetime.now().isoformat(timespec="seconds"))
    log.info("python:         %s  (%s)", sys.version.split()[0], sys.executable)
    log.info("platform:       %s  %s", platform.system(), platform.release())
    log.info("cwd:            %s", os.getcwd())
    log.info("script:         %s", Path(__file__).resolve())
    log.info("argv:           %s", " ".join(sys.argv))

    try:
        import openwraw  # noqa: WPS433
        v = getattr(openwraw, "__version__", "unknown")
        log.info("openwraw:       %s", v)
    except Exception as exc:  # noqa: BLE001
        log.error("openwraw:       NOT INSTALLED (%s)", exc)
        log.error("Install with:   python -m pip install --user openwraw")
        raise SystemExit(2)

    log.info("input folder:   %s", args.folder.resolve())
    log.info("mass window:    [%.4f, %.4f]", args.mass_lo, args.mass_hi)
    log.info("function:       %s", args.func if args.func is not None else "auto (first non-lock-mass)")
    log.info("found %d .raw director%s:", len(raw_dirs),
             "y" if len(raw_dirs) == 1 else "ies")
    for d in raw_dirs:
        ce = extract_ce(d.name)
        ce_label = f"CE={ce:g}" if ce is not None else "CE=??"
        log.info("    %-12s  %s", ce_label, d.name)
    log.info("-" * 70)


# ---------------------------------------------------------------------------
# Discovery
# ---------------------------------------------------------------------------

def find_raw_dirs(folder: Path) -> list[Path]:
    if not folder.exists():
        log.error("folder not found: %s", folder)
        raise SystemExit(2)
    if not folder.is_dir():
        log.error("not a directory: %s", folder)
        raise SystemExit(2)
    return sorted(
        p for p in folder.iterdir()
        if p.is_dir() and p.suffix.lower() == ".raw"
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    p = argparse.ArgumentParser(
        description=(
            "Compute fraction = sum(intensity in [mass_lo, mass_hi]) / "
            "sum(all intensity) for every .raw in a folder, indexed by the "
            "CE number embedded in each filename."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Example:\n"
            "  python extract_intensities.py Messdaten "
            "--mass-lo 1369 --mass-hi 1373.2\n"
        ),
    )
    p.add_argument("folder", type=Path,
                   help="Folder containing one or more *.raw directories.")
    p.add_argument("--mass-lo", type=float, required=True,
                   help="Lower m/z bound for the integration window.")
    p.add_argument("--mass-hi", type=float, required=True,
                   help="Upper m/z bound for the integration window.")
    p.add_argument("--func", type=int, default=None,
                   help="1-based MS function index (default: first non-lock-mass).")
    p.add_argument("--out-dir", type=Path, default=None,
                   help="Output directory (default: run-<timestamp> next to the script).")
    p.add_argument("-v", "--verbose", action="count", default=0,
                   help="-v: show DEBUG on console too.")
    args = p.parse_args()

    if args.mass_hi <= args.mass_lo:
        print(f"ERROR: --mass-hi ({args.mass_hi}) must be > --mass-lo ({args.mass_lo})",
              file=sys.stderr)
        return 2

    # Output directory next to the script (so it doesn't pollute the data folder)
    if args.out_dir is None:
        ts = _dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        args.out_dir = Path(__file__).resolve().parent / f"run-{ts}"

    console_level = logging.DEBUG if args.verbose else logging.INFO
    log_path = setup_logging(args.out_dir, console_level)

    raw_dirs = find_raw_dirs(args.folder)
    preflight(args, raw_dirs)

    if not raw_dirs:
        log.error("No *.raw directories found in %s", args.folder)
        return 2

    # Per-file processing -----------------------------------------------------

    import openwraw  # noqa: WPS433  — checked in preflight

    rows = []  # list[dict] for TSV
    for raw_dir in raw_dirs:
        log.info("")
        log.info("[%s]", raw_dir.name)
        ce = extract_ce(raw_dir.name)
        if ce is None:
            log.warning("  no 'CE_<n>' found in filename — skipping")
            continue

        try:
            reader = openwraw.RawReader(str(raw_dir))
        except Exception as exc:  # noqa: BLE001
            log.error("  openwraw failed to open: %s", exc)
            log.debug("  traceback:\n%s", traceback.format_exc())
            continue

        log.info("  CE=%g  %d function(s)", ce, len(reader.functions))
        for f in reader.functions:
            log.debug("    %r", f)

        try:
            func = pick_ms_function(reader, args.func)
        except Exception as exc:  # noqa: BLE001
            log.error("  %s", exc)
            continue

        windowed, total, n_scans, n_peaks, n_skipped = integrate_raw(
            reader, func, args.mass_lo, args.mass_hi,
        )
        if total == 0.0:
            log.warning("  total intensity is 0 — fraction undefined, set to NaN")
            fraction = float("nan")
        else:
            fraction = windowed / total

        log.info("  -> total=%.1f  windowed=%.1f  fraction=%.6f",
                 total, windowed, fraction)

        rows.append({
            "ce": ce,
            "fraction": fraction,
            "windowed": windowed,
            "total": total,
            "n_scans": n_scans,
            "n_peaks": n_peaks,
            "func": func.index,
            "filename": raw_dir.name,
        })

    if not rows:
        log.error("No usable .raw files. See %s for details.", log_path)
        return 1

    rows.sort(key=lambda r: r["ce"])

    # ------------------------------------------------------------------ output
    tsv_path = args.out_dir / "summary.tsv"
    txt_path = args.out_dir / "summary.txt"

    with tsv_path.open("w", encoding="utf-8") as fh:
        fh.write("ce\tfraction\twindowed\ttotal\tn_scans\tn_peaks\tfunc\tfilename\n")
        for r in rows:
            fh.write(
                f"{r['ce']:g}\t{r['fraction']:.8f}\t{r['windowed']:.3f}\t"
                f"{r['total']:.3f}\t{r['n_scans']}\t{r['n_peaks']}\t"
                f"{r['func']}\t{r['filename']}\n"
            )

    with txt_path.open("w", encoding="utf-8") as fh:
        fh.write(f"# mass window: [{args.mass_lo}, {args.mass_hi}]\n")
        fh.write("# CE\tfraction\n")
        for r in rows:
            fh.write(f"{r['ce']:g}\t{r['fraction']:.6f}\n")

    log.info("")
    log.info("=" * 70)
    log.info("RESULTS  (mass window [%.4f, %.4f])", args.mass_lo, args.mass_hi)
    log.info("=" * 70)
    log.info("%6s  %12s  %14s  %14s  %s",
             "CE", "fraction", "windowed", "total", "filename")
    for r in rows:
        log.info("%6g  %12.6f  %14.1f  %14.1f  %s",
                 r["ce"], r["fraction"], r["windowed"], r["total"], r["filename"])
    log.info("")
    log.info("wrote: %s", tsv_path)
    log.info("wrote: %s", txt_path)
    log.info("wrote: %s", log_path)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
