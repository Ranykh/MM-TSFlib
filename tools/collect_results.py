#!/usr/bin/env python3
"""Collect MM-TSFlib runs into one tidy CSV, and detect the silent data loss first.

    python tools/collect_results.py                        # scan ./results
    python tools/collect_results.py --txt result_*.txt     # also parse the text logs
    python tools/collect_results.py --pivot                # add a uni-vs-multi table
    python tools/collect_results.py --expect 24            # assert the sweep's run count
    python tools/collect_results.py --check                # report only, write nothing

READ THIS BEFORE TRUSTING ANY NUMBER FROM THIS REPO
----------------------------------------------------
MM-TSFlib's `setting` string (run.py, around lines 200-218) is built from
task / model_id / model / data / features / seq_len / label_len / pred_len /
d_model / ... It does NOT contain llm_model, text_len, prompt_weight or
pool_type.

Consequence: two runs that differ ONLY by which LLM they used land in the same
results/ folder, and the second silently overwrites the first. There is no
error, no warning, and nothing in the output to tell you it happened. You end up
with a table of numbers where several rows are the same run.

The remedy is to encode those fields in --model_id, following the convention the
repro scripts use:

    <Domain>_<uni|multi>_<llm>_tl<text_len>_pw<prompt_weight>_sl<seq_len>_s<seed>

This script checks that convention is being followed and refuses to pretend
everything is fine when it is not. A run whose model_id does not follow it still
parses; the missing fields come back empty and are counted as at-risk.

metrics.npy is np.array([mae, mse, rmse, mape, mspe]) -- MAE FIRST. Reading
index 0 as MSE is a classic and silent error.

Two rules about MM-TSFlib numbers
---------------------------------
1. Never average MSE across domains. Their scales differ by three orders of
   magnitude (Agriculture ~0.06 vs Security ~100-130). Report per-domain, or
   report win rates.
2. Time-MMD publishes MSE only. Do not invent MAE comparisons against it.
"""
import argparse
import csv
import datetime as _dt
import glob
import os
import re
import subprocess
import sys

try:
    import numpy as np
except ImportError:
    sys.exit("needs numpy: pip install numpy")


# model_id convention from scripts/repro/*.sh
MODEL_ID_RE = re.compile(
    r"^(?P<domain>.+?)_(?P<tag>uni|multi)_(?P<llm>[A-Za-z0-9]+)"
    r"_tl(?P<text_len>\d+)_pw(?P<prompt_weight>[0-9.]+)"
    r"_sl(?P<seq_len>\d+)_s(?P<seed>\d+)$"
)

SETTING_RE = re.compile(
    r"^(?P<task>long_term_forecast)_(?P<model_id>.+?)_(?P<model>[A-Za-z_]+?)_"
    r"(?P<data>custom)_ft(?P<features>\w+)_sl(?P<seq_len>\d+)_ll(?P<label_len>\d+)_"
    r"pl(?P<pred_len>\d+)_dm(?P<d_model>\d+)_nh(?P<n_heads>\d+)_el(?P<e_layers>\d+)_"
    r"dl(?P<d_layers>\d+)_df(?P<d_ff>\d+)_"
)

FIELDS = ["repo", "git_sha", "git_dirty", "domain", "tag", "model", "llm",
          "pred_len", "seq_len", "label_len", "text_len", "prompt_weight", "seed",
          "mae", "mse", "rmse", "mape", "mspe",
          "model_id_ok", "setting", "run_dir", "mtime", "source"]


def git_info(repo_root):
    def _run(a):
        return subprocess.run(a, cwd=repo_root, capture_output=True,
                              text=True, timeout=10)
    try:
        sha = _run(["git", "rev-parse", "HEAD"])
        if sha.returncode != 0:
            return "NO-GIT", ""
        st = _run(["git", "status", "--porcelain"])
        return sha.stdout.strip(), ("yes" if st.stdout.strip() else "no")
    except (OSError, subprocess.SubprocessError):
        return "NO-GIT", ""


def parse_setting(setting):
    m = SETTING_RE.match(setting)
    if not m:
        return None
    d = m.groupdict()
    out = {
        "model": d["model"], "seq_len": int(d["seq_len"]),
        "label_len": int(d["label_len"]), "pred_len": int(d["pred_len"]),
        "domain": "", "tag": "", "llm": "", "text_len": "",
        "prompt_weight": "", "seed": "", "model_id_ok": "NO",
    }
    mid = MODEL_ID_RE.match(d["model_id"])
    if mid:
        g = mid.groupdict()
        out.update({
            "domain": g["domain"], "tag": g["tag"], "llm": g["llm"],
            "text_len": int(g["text_len"]),
            "prompt_weight": float(g["prompt_weight"]),
            "seed": int(g["seed"]), "model_id_ok": "yes",
        })
    else:
        # Not the convention. The run is usable but cannot be told apart from
        # another that differs only by LLM / text_len / prompt_weight.
        out["domain"] = d["model_id"]
    return out


def from_results_dir(root, sha, dirty):
    rows, skipped = [], []
    if not os.path.isdir(root):
        return rows, skipped
    for name in sorted(os.listdir(root)):
        run_dir = os.path.join(root, name)
        metrics = os.path.join(run_dir, "metrics.npy")
        if not os.path.exists(metrics):
            continue
        parsed = parse_setting(name)
        if parsed is None:
            skipped.append(name)
            continue
        try:
            arr = np.load(metrics)
            mae, mse, rmse, mape, mspe = (float(x) for x in arr[:5])
        except (OSError, ValueError) as e:
            skipped.append(f"{name}  (bad metrics.npy: {e})")
            continue
        parsed.update({
            "repo": "MM-TSFlib", "git_sha": sha, "git_dirty": dirty,
            "mae": mae, "mse": mse, "rmse": rmse, "mape": mape, "mspe": mspe,
            "setting": name, "run_dir": run_dir, "source": "results",
            "mtime": _dt.datetime.fromtimestamp(
                os.path.getmtime(run_dir)).isoformat(timespec="seconds"),
        })
        rows.append(parsed)
    return rows, skipped


def from_txt(patterns, sha, dirty):
    """Parse --save_name text logs: <setting> on one line, metrics on the next.

    Useful when results/ was cleared but the append-only log survived. Note the
    log is append-only, so it can hold BOTH sides of an overwrite -- which makes
    it the only place a clobbered run can still be recovered from.
    """
    rows = []
    num = r"[-+0-9.eE]+"
    metric_re = re.compile(
        r"mse:({0}), mae:({0}), rmse:({0}), mape:({0}), mspe:({0})".format(num))
    for pat in patterns:
        for path in sorted(glob.glob(pat)):
            with open(path) as fh:
                lines = [ln.strip() for ln in fh if ln.strip()]
            for i, ln in enumerate(lines):
                m = metric_re.search(ln)
                if not m or i == 0:
                    continue
                parsed = parse_setting(lines[i - 1])
                if parsed is None:
                    continue
                mse, mae, rmse, mape, mspe = (float(x) for x in m.groups())
                parsed.update({
                    "repo": "MM-TSFlib", "git_sha": sha, "git_dirty": dirty,
                    "mae": mae, "mse": mse, "rmse": rmse, "mape": mape,
                    "mspe": mspe, "setting": lines[i - 1], "run_dir": path,
                    "mtime": "", "source": "txt",
                })
                rows.append(parsed)
    return rows


def dedupe(rows):
    """Prefer results/ rows over txt rows for the same setting."""
    best = {}
    for r in rows:
        k = r["setting"]
        if k not in best or (best[k]["source"] == "txt" and r["source"] == "results"):
            best[k] = r
    return list(best.values())


def audit(rows, skipped, expect, results_root):
    """The overwrite check, run BEFORE anything is reported as a result."""
    print("=" * 68)
    print("AUDIT -- silent data loss check")
    print("=" * 68)
    print(f"results dir:      {results_root}")
    print(f"runs parsed:      {len(rows)}")
    if skipped:
        print(f"folders skipped:  {len(skipped)}")
        for s in skipped[:8]:
            print(f"                    {s}")

    bad_id = [r for r in rows if r["model_id_ok"] != "yes"]
    fatal = False

    if bad_id:
        fatal = True
        print()
        print(f"!! {len(bad_id)} of {len(rows)} runs have a model_id that does NOT encode")
        print("!! llm / text_len / prompt_weight. MM-TSFlib's setting string omits those")
        print("!! fields, so any two such runs differing only by LLM WROTE TO THE SAME")
        print("!! FOLDER and the later one won.")
        print("!!")
        print("!! THESE RESULTS ARE ALREADY CORRUPTED. Re-running with a --model_id of")
        print("!! the form <Domain>_<uni|multi>_<llm>_tl<N>_pw<F>_sl<N>_s<N> is the only")
        print("!! remedy. Do not put these numbers in the workbook.")
        print("!!")
        for r in bad_id[:8]:
            print(f"!!   {r['setting'][:88]}")
        if len(bad_id) > 8:
            print(f"!!   ... and {len(bad_id) - 8} more")

    if expect is not None:
        print()
        if len(rows) == expect:
            print(f"run count:        {len(rows)} == {expect} expected. OK.")
        else:
            fatal = True
            missing = expect - len(rows)
            print(f"!! run count:     {len(rows)} found, {expect} expected "
                  f"({'MISSING ' + str(missing) if missing > 0 else 'EXTRA ' + str(-missing)}).")
            if missing > 0:
                print("!! Either runs crashed, or they overwrote each other. Check the")
                print("!! sweep log for failures before assuming the former.")

    # Distinct-key collision check: two different settings that describe the
    # same experiment mean something is being distinguished only by chance.
    keys = {}
    for r in rows:
        k = (r["domain"], r["tag"], r["model"], r["llm"], r["pred_len"],
             r["text_len"], r["prompt_weight"], r["seed"])
        keys.setdefault(k, []).append(r)
    dupes = {k: v for k, v in keys.items() if len(v) > 1}
    if dupes:
        print()
        print(f"!  {len(dupes)} experiment key(s) appear more than once:")
        for k, v in list(dupes.items())[:5]:
            print(f"     {k} x{len(v)}")

    if rows and rows[0]["git_dirty"] == "yes":
        print()
        print("!! git_dirty=yes -- the working tree had uncommitted changes, so these")
        print("!! results are not tied to the recorded SHA.")

    # Stale-folder check: results much older than the rest of the sweep are
    # probably left over from a previous run.
    stamped = [r for r in rows if r["mtime"]]
    if len(stamped) > 2:
        times = sorted(r["mtime"] for r in stamped)
        if (_dt.datetime.fromisoformat(times[-1])
                - _dt.datetime.fromisoformat(times[0])).days >= 1:
            print()
            print(f"!  results span more than a day ({times[0]} .. {times[-1]}).")
            print("!  Older folders may be cached runs from a previous sweep. A number")
            print("!  that matches the paper suspiciously well is often one of these.")

    print("=" * 68)
    return fatal


def write_pivot(rows, out):
    def key(r):
        return (r["domain"], r["model"], str(r["llm"]), r["pred_len"])
    uni = {key(r): r["mse"] for r in rows if r["tag"] == "uni"}
    multi = {key(r): r["mse"] for r in rows if r["tag"] == "multi"}
    both = sorted(set(uni) & set(multi))
    pivot = out.replace(".csv", "_pivot.csv")
    with open(pivot, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["domain", "model", "llm", "pred_len",
                    "uni_mse", "multi_mse", "delta", "multi_better"])
        for k in both:
            d = multi[k] - uni[k]
            w.writerow(list(k) + [f"{uni[k]:.6f}", f"{multi[k]:.6f}",
                                  f"{d:.6f}", "yes" if d < 0 else "NO"])
    print(f"wrote {pivot} ({len(both)} paired configs)")
    if not both:
        print("no uni/multi pairs yet. Remember: --prompt_weight 0 IS the unimodal")
        print("baseline -- MM-TSFlib's weight1/weight2 are dead code, never used in")
        print("any forward pass, so fusion is the constant --prompt_weight.")
        return
    losses = sum(1 for k in both if multi[k] >= uni[k])
    print(f"multimodal loses to unimodal in {losses}/{len(both)} "
          f"({losses / len(both):.1%}).")
    print("Time-MMD reports multimodal winning ~95% of configs. A much worse rate")
    print("means prompt_weight is mistuned for this domain -- it is a hand-set")
    print("constant, tuned per domain in the original work.")


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    here = os.path.dirname(os.path.abspath(__file__))
    repo_root = os.path.abspath(os.path.join(here, ".."))
    ap.add_argument("--results", default=os.path.join(repo_root, "results"))
    ap.add_argument("--repo-root", default=repo_root)
    ap.add_argument("--txt", nargs="*", default=[],
                    help="also parse these --save_name logs")
    ap.add_argument("--out", default="mmtsflib_results.csv")
    ap.add_argument("--expect", type=int, default=None,
                    help="number of runs the sweep should have produced")
    ap.add_argument("--pivot", action="store_true",
                    help="also write a uni-vs-multi pivot with the delta")
    ap.add_argument("--check", action="store_true",
                    help="audit only, write nothing")
    ap.add_argument("--force", action="store_true",
                    help="write the CSV even if the audit found corruption")
    args = ap.parse_args()

    sha, dirty = git_info(args.repo_root)
    rows, skipped = from_results_dir(args.results, sha, dirty)
    rows = dedupe(rows + from_txt(args.txt, sha, dirty))

    if not rows:
        sys.exit(f"no runs found. checked {args.results} "
                 f"and {args.txt or 'no logs'}")

    rows.sort(key=lambda r: (r["domain"], r["model"], str(r["llm"]),
                             r["pred_len"], r["tag"], str(r["seed"])))

    fatal = audit(rows, skipped, args.expect, args.results)

    if args.check:
        print("\n--check: nothing written.")
        return 1 if fatal else 0

    if fatal and not args.force:
        print("\nREFUSING to write a CSV from results the audit flagged as corrupted.")
        print("Fix the model_id convention and re-run the sweep, or pass --force if")
        print("you genuinely intend to keep these numbers anyway.")
        return 1

    with open(args.out, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})
    print(f"\nwrote {args.out} ({len(rows)} runs)")

    if args.pivot:
        write_pivot(rows, args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
