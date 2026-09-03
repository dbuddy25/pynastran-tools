"""Profile an OP2 read: how long it takes, and what is actually in the file.

Answers two questions a slow load raises:
  1. Where is the size?  If stress/strain/energy dominate an 800 MB file that
     you only plot accelerations from, the deck is asking for too much.
  2. Would a filtered read help?  --filtered times a set_results() read against
     the full one, which is also the check that gates the
     perf/filtered-op2-read branch.

Usage:
    py -3 op2_profile.py model.op2
    py -3 op2_profile.py model.op2 --filtered

Needs only pyNastran + numpy — no GUI, no customtkinter.
"""

import argparse
import os
import sys
import time


# The four families the ASD tools actually read.
ASD_RESULT_NAMES = [
    'displacements', 'accelerations', 'spc_forces',
    'force.cbush_force', 'cbush_force',
    'psd.displacements', 'psd.accelerations', 'psd.spc_forces',
    'psd.cbush_force',
    'rms.displacements', 'rms.accelerations', 'rms.spc_forces',
    'rms.cbush_force',
]

ASD_ATTRS = ('displacements', 'accelerations', 'spc_forces', 'cbush_force')

# Only these containers are read by the tools. ato.* (autocorrelation, from
# ATOC) and crm.* (from CRMS) are parsed and thrown away — RALL requests them.
ASD_PREFIXES = ('', 'psd.', 'rms.')


def _is_used(label):
    """True if the ASD tools actually read this result."""
    head, _, tail = label.rpartition(".")
    prefix = f"{head}." if head else ""
    return prefix in ASD_PREFIXES and tail in ASD_ATTRS


def _human(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:,.1f} {unit}"
        n /= 1024.0


def _table_bytes(tbl):
    """Bytes held by one result table's arrays.

    Duck-typed on .nbytes so this module needs no numpy import of its own —
    it should be able to start and explain itself on a bare interpreter.
    """
    total = 0
    for attr in ("data", "_times", "node_gridtype", "element", "element_node"):
        nbytes = getattr(getattr(tbl, attr, None), "nbytes", None)
        if isinstance(nbytes, int):
            total += nbytes
    return total


def _check_deps():
    """Explain a missing pyNastran instead of dying on the import."""
    try:
        import pyNastran  # noqa: F401
        return True
    except ImportError as exc:
        exe = sys.executable
        print(f"Cannot import pyNastran ({exc}).\n")
        print(f"This interpreter is:\n  {exe}\n")
        print("It is probably not the one the Structures Tools launcher set up.")
        print("On Windows 'py -3' and 'python' are often different installs —")
        print("try 'python' rather than 'py -3':\n")
        print(f"  python {os.path.basename(__file__)} model.op2\n")
        print("Or point at the launcher's environment directly, e.g.")
        print("  %LOCALAPPDATA%\\StructuresTools\\.venv\\Scripts\\python.exe "
              f"{os.path.basename(__file__)} model.op2\n")
        print("Or install into this one:")
        print(f"  \"{exe}\" -m pip install pyNastran numpy")
        return False


def _walk(op2):
    """Yield (label, n_subcases, bytes) for every populated result dict."""
    def _scan(container, prefix):
        # seen is per-container: op2.accelerations and op2_results.psd.
        # accelerations share a name but are different results.
        seen = set()
        for name in dir(container):
            if name.startswith("_") or name in seen:
                continue
            try:
                val = getattr(container, name)
            except Exception:
                continue
            if not isinstance(val, dict) or not val:
                continue
            nbytes = 0
            ok = False
            for tbl in val.values():
                if hasattr(tbl, "data"):
                    nbytes += _table_bytes(tbl)
                    ok = True
            if ok:
                seen.add(name)
                yield f"{prefix}{name}", len(val), nbytes

    yield from _scan(op2, "")
    results = getattr(op2, "op2_results", None)
    if results is not None:
        for sub in ("psd", "rms", "ato", "crm", "no"):
            container = getattr(results, sub, None)
            if container is not None:
                yield from _scan(container, f"{sub}.")


def _read(path, filtered):
    """Read the OP2. Returns (op2, seconds, rejected_reason_or_None).

    A rejected set_results() is fatal to the comparison, not a warning: the
    exception leaves the result set half-configured with most results
    disabled, so reading anyway looks dramatically faster while silently
    dropping nearly everything. Bail instead of reporting a fake speedup.
    """
    from pyNastran.op2.op2 import OP2
    op2 = OP2(mode='nx', debug=False)
    if filtered:
        try:
            op2.set_results(ASD_RESULT_NAMES)
        except Exception as exc:
            first = str(exc).strip().splitlines()[0]
            return None, 0.0, first
    t0 = time.perf_counter()
    op2.read_op2(path)
    return op2, time.perf_counter() - t0, None


def _has_family(op2, attr):
    results = getattr(op2, "op2_results", None)
    for sub in ("psd", "rms"):
        container = getattr(results, sub, None)
        if getattr(container, attr, None):
            return True
    return bool(getattr(op2, attr, None))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("op2")
    ap.add_argument("--filtered", action="store_true",
                    help="also time a set_results() read and compare families")
    ap.add_argument("--top", type=int, default=15,
                    help="how many result types to list (default 15)")
    args = ap.parse_args(argv)

    if not _check_deps():
        return 2

    if not os.path.isfile(args.op2):
        print(f"No such file: {args.op2}")
        return 1

    on_disk = os.path.getsize(args.op2)
    print(f"File: {args.op2}\nSize on disk: {_human(on_disk)}")
    print("(tip: redirect to a file — the table is long)\n")

    print("Full read…")
    op2, secs, _ = _read(args.op2, filtered=False)
    print(f"  {secs:.1f} s  ({_human(on_disk / secs)}/s)\n")

    rows = sorted(_walk(op2), key=lambda r: -r[2])
    in_mem = sum(r[2] for r in rows)
    asd_bytes = sum(r[2] for r in rows if _is_used(r[0]))
    waste = {}
    for label, _n, nbytes in rows:
        head = label.rpartition(".")[0]
        if head in ("ato", "crm", "no"):
            waste[head] = waste.get(head, 0) + nbytes

    print(f"{'result type':<34}{'subcases':>9}{'arrays':>14}")
    print("-" * 57)
    for label, n, nbytes in rows[:args.top]:
        mark = " *" if _is_used(label) else ""
        print(f"{label:<34}{n:>9}{_human(nbytes):>14}{mark}")
    if len(rows) > args.top:
        rest = sum(r[2] for r in rows[args.top:])
        print(f"{f'... {len(rows) - args.top} more':<34}{'':>9}{_human(rest):>14}")
    print("-" * 57)
    print(f"{'total in memory':<34}{'':>9}{_human(in_mem):>14}")
    print(f"{'* used by the ASD tools':<34}{'':>9}{_human(asd_bytes):>14}")
    if in_mem:
        pct = 100.0 * asd_bytes / in_mem
        print(f"\nThe ASD tools use {pct:.1f}% of what was parsed.")
        if pct < 50:
            print("Most of the read is results these tools never touch — trim the\n"
                  "deck's output requests, or filter the read.")
    if waste:
        total = sum(waste.values())
        names = {"ato": "ATOC (autocorrelation)", "crm": "CRMS (cumulative RMS)",
                 "no": "NO (zero crossings)"}
        print(f"\nRALL-only output, never read by these tools: {_human(total)}")
        for k, v in sorted(waste.items(), key=lambda kv: -kv[1]):
            print(f"  {names.get(k, k):<28}{_human(v):>12}")
        print("  Drop these from the request (PSDF instead of RALL) to shrink\n"
              "  both the .op2 and the read.")

    if args.filtered:
        print("\nFiltered read (set_results)…")
        f_op2, f_secs, rejected = _read(args.op2, filtered=True)
        if rejected is not None:
            print(f"  set_results REJECTED the result names:\n    {rejected}\n")
            print("  No comparison is possible — a rejected set_results leaves most\n"
                  "  results disabled, so timing it would report a large but\n"
                  "  meaningless speedup. The names in ASD_RESULT_NAMES are wrong\n"
                  "  for this pyNastran; do not merge perf/filtered-op2-read.")
            return 0
        print(f"  {f_secs:.1f} s")
        if f_secs > 0:
            print(f"  {secs / f_secs:.1f}x faster" if f_secs < secs else "  no faster")
        print("\n  family            full   filtered")
        bad = []
        for attr in ASD_ATTRS:
            a, b = _has_family(op2, attr), _has_family(f_op2, attr)
            flag = "" if a == b else "   <-- LOST"
            if a and not b:
                bad.append(attr)
            print(f"  {attr:<16}{str(a):>6}{str(b):>11}{flag}")
        print("\n  " + ("SAFE: filtered read kept every family the full read found."
                        if not bad else
                        f"UNSAFE: filtered read lost {', '.join(bad)}. "
                        "Do not merge perf/filtered-op2-read."))
    return 0


if __name__ == "__main__":
    sys.exit(main())
