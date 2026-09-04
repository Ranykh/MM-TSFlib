#!/usr/bin/env python3
"""Merge reproduction CSVs into the MM-MoGU comparison workbook.

    python tools/build_workbook.py \
        --base "MMMoGU_Results_Comparison_v2.xlsx" \
        --out  "MMMoGU_Results_Comparison_v3.xlsx" \
        --gmmts   gmmts_g3.csv gmmts_g4.csv \
        --timemmd timemmd_Economy.csv \
        --mogu    mogu_t1.csv

    python tools/build_workbook.py --base v2.xlsx --out v3.xlsx \
        --gmmts gmmts_g4.csv --audit-only

Why the workbook needs changing at all
--------------------------------------
v2's own PROTOCOL says "Reproduce the published row FIRST and confirm it matches
before filling any of your own columns." But every blank column in v2 is
G4 IV / G5 SNIV / G6 UCA / G7 TUG -- those are Phase 2, our new gates. There was
nowhere to record a reproduction. This script adds that place.

What it writes
--------------
1.  `Raw_Runs` -- one row per run, from any of the three collectors
    (moe_unc_tsf/collect_mogu.py, MM-TSFlib/collect_results.py,
    gmmts_lib/collect_gmmts.py), normalised into one schema.

2.  `Pairwise_Baselines` -- THE DELIVERABLE. Beside the locked published
    G1 FIXED (TimeMMD) and G3 ATTN (GMM-TS) columns it adds, as live formulas:
        R: G1 FIXED (ours)      our TimeMMD reproduction
        R: G3 ATTN  (ours)      our GMM-TS reproduction -- the harness check
        R: G4 IV    (MM-MoGU)   our method
        Δ vs published for each, then G4 − our G3, then a verdict.
    The G4 vs OUR G3 column is the one that matters: it is the comparison made
    inside a single codebase, under one training recipe, rather than against
    someone else's printed numbers.

3.  `MoGU_ETT` -- the same treatment for MoGU Tables 1, 2 and 4.

4.  An audit: every filter written into a formula is resolved against the source
    rows in Python, and what each formula WILL show is printed.

The published (teal) columns are never touched.

Standard deviations without STDEVIFS
------------------------------------
Excel has AVERAGEIFS and COUNTIFS but no STDEVIFS, and array formulas are
fragile across versions. So `Raw_Runs` carries live mse^2 helper columns and the
std is the algebraic identity

    s = sqrt( ( SUM(x^2) - n * mean^2 ) / (n - 1) )

using only SUMIFS / COUNTIFS / AVERAGEIFS. Works in every Excel and in Sheets.

Ranges are bounded (row 2 to MAX_RAW_ROW) rather than whole-column: with several
thousand AVERAGEIFS on one sheet, whole-column references make Excel crawl.

There is no LibreOffice on the Mac, so formula cells carry no cached values
until Excel or Sheets opens the file. The --audit output stands in for that.
"""
import argparse
import csv
import os
import sys

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as CL
except ImportError:
    sys.exit("needs openpyxl: pip install openpyxl")


# --------------------------------------------------------------------------
# Raw_Runs layout -- a union schema covering all three collectors.
# These positions are load-bearing: every formula references them by letter,
# and audit_layout() asserts they still hold.
# --------------------------------------------------------------------------
RAW_COLUMNS = [
    "run_id", "repo", "git_sha", "git_dirty",
    "method",          # unified verb: single/MoE/MoGE/MoGU/G3_ATTN_direct/G4_IV/G1_FIXED_multi/...
    "config",          # moe_unc_tsf's own label
    "dataset",         # ETTh1 etc.  (moe_unc_tsf)
    "domain_sheet",    # "Economy", "Social Good" ... as the workbook spells it
    "model",           # numeric expert / TSF-N
    "tsf_t",           # text expert / TSF-T / llm
    "num_experts", "pred_len", "seq_len", "seed", "learning_rate",
    "agg_type", "inv_var_norm", "prob_expert", "unc_gating",
    "prompt_weight", "text_len", "tag",
    "mse", "mae", "rmse",
    "gate_w_std", "gate_collapsed",
    "model_id", "run_dir", "mtime", "source_csv",
    "mse_sq", "mae_sq",
]
COL = {name: CL(i + 1) for i, name in enumerate(RAW_COLUMNS)}
RAW = "Raw_Runs"

FIRST_DATA_ROW = 2
MAX_RAW_ROW = 5000        # bounded ranges keep Excel responsive

REPRO_HDR_FILL = PatternFill("solid", fgColor="FF1E5B3A")
REPRO_HDR_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
REPRO_VAL_FILL = PatternFill("solid", fgColor="FFE8F4EC")
REPRO_VAL_FONT = Font(name="Arial", size=10, color="FF11402A")
CLAIM_HDR_FILL = PatternFill("solid", fgColor="FF7A4B12")   # the G4-vs-G3 block
CLAIM_VAL_FILL = PatternFill("solid", fgColor="FFFBF0DF")
DELTA_FONT = Font(name="Arial", size=10)
STATUS_FONT = Font(name="Arial", size=10, bold=True)

MSE_FMT = "0.0000"
DELTA_FMT = '0.0000;[Red]-0.0000'
PCT_FMT = '0.0%;[Red]-0.0%'

TOL_MATCH = 0.02
TOL_CLOSE = 0.05

DEFAULT_NE = 3

DATASET_ALIAS = {
    "ETTh1": "ETTh1", "ETTh2": "ETTh2", "ETTm1": "ETTm1", "ETTm2": "ETTm2",
    "ILI": "national-illness", "Weather": "weather",
    "Electricity": "electricity", "Exchange": "exchange-rate",
}

# MM-TSFlib data-folder name -> the workbook's domain label.
DOMAIN_TO_SHEET = {
    "Algriculture": "Agriculture", "Agriculture": "Agriculture",
    "Climate": "Climate", "Economy": "Economy", "Energy": "Energy",
    "Environment": "Environment", "Public_Health": "Public Health",
    "Security": "Security", "SocialGood": "Social Good", "Traffic": "Traffic",
}
# The workbook writes GPT-3.5; the code calls the closed model ClosedLLM.
TSFT_ALIAS = {"ClosedLLM": "GPT3.5", "GPT3.5": "GPT3.5"}


# ==========================================================================
# Loading and normalisation
# ==========================================================================
def load_runs(csv_paths):
    rows = []
    for path in csv_paths:
        if not os.path.exists(path):
            print(f"  ! missing CSV, skipped: {path}")
            continue
        with open(path, newline="") as fh:
            n = 0
            for r in csv.DictReader(fh):
                r["source_csv"] = os.path.basename(path)
                rows.append(r)
                n += 1
        print(f"  read {n:>4} runs from {path}")
    return rows


def _f(row, key):
    v = row.get(key, "")
    if v in ("", None):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _i(row, key):
    v = _f(row, key)
    return None if v is None else int(v)


def normalise(r):
    """Map any collector's row onto the union schema.

    The three collectors describe different experiments and cannot share a
    vocabulary by accident, so the mapping is explicit and per-repo.
    """
    repo = r.get("repo", "")
    out = {k: "" for k in RAW_COLUMNS}
    out.update({
        "repo": repo,
        "git_sha": r.get("git_sha", ""),
        "git_dirty": r.get("git_dirty", ""),
        "pred_len": _i(r, "pred_len"),
        "seq_len": _i(r, "seq_len"),
        "seed": _i(r, "seed"),
        "mse": _f(r, "mse"),
        "mae": _f(r, "mae"),
        "rmse": _f(r, "rmse"),
        "gate_w_std": _f(r, "gate_w_std"),
        "gate_collapsed": r.get("gate_collapsed", ""),
        "model_id": r.get("model_id", ""),
        "run_dir": r.get("run_dir", ""),
        "mtime": r.get("mtime", ""),
        "source_csv": r.get("source_csv", ""),
    })

    if repo == "gmmts_lib":
        # Already carries method / domain_sheet / tsf_n / tsf_t.
        out.update({
            "method": r.get("method", ""),
            "domain_sheet": r.get("domain_sheet", ""),
            "model": r.get("tsf_n", ""),
            "tsf_t": TSFT_ALIAS.get(r.get("tsf_t", ""), r.get("tsf_t", "")),
            "agg_type": r.get("agg_type", ""),
            "inv_var_norm": r.get("inv_var_norm", ""),
            "prob_expert": _i(r, "prob_expert"),
        })
    elif repo == "MM-TSFlib":
        # TimeMMD: the fixed-weight baseline. tag is uni or multi, and
        # prompt_weight 0 IS the unimodal baseline (weight1/weight2 are dead
        # code, so fusion is that constant).
        tag = r.get("tag", "")
        out.update({
            "method": f"G1_FIXED_{tag}" if tag else "G1_FIXED",
            "domain_sheet": DOMAIN_TO_SHEET.get(r.get("domain", ""), r.get("domain", "")),
            "model": r.get("model", ""),
            "tsf_t": TSFT_ALIAS.get(r.get("llm", ""), r.get("llm", "")),
            "prompt_weight": _f(r, "prompt_weight"),
            "text_len": _i(r, "text_len"),
            "tag": tag,
        })
    elif repo == "moe_unc_tsf":
        out.update({
            "method": r.get("config", ""),
            "config": r.get("config", ""),
            "dataset": r.get("dataset", ""),
            "model": r.get("model", ""),
            "num_experts": _i(r, "num_experts"),
            "learning_rate": _f(r, "learning_rate"),
            "prob_expert": _i(r, "prob_expert"),
            "unc_gating": _i(r, "unc_gating"),
        })
    else:
        out["method"] = r.get("method", r.get("config", ""))
        out["model"] = r.get("model", "")
    return out


# ==========================================================================
# Raw_Runs sheet
# ==========================================================================
def write_raw_runs(wb, runs):
    if RAW in wb.sheetnames:
        del wb[RAW]
    ws = wb.create_sheet(RAW)

    # Row 1 is the header and nothing else: every formula below assumes the
    # data starts at row 2.
    for i, name in enumerate(RAW_COLUMNS, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.font = REPRO_HDR_FONT
        c.fill = REPRO_HDR_FILL
        c.alignment = Alignment(horizontal="center")

    rows = [normalise(r) for r in runs]
    rows.sort(key=lambda r: (str(r["repo"]), str(r["domain_sheet"]), str(r["dataset"]),
                             r["pred_len"] or 0, str(r["model"]), str(r["tsf_t"]),
                             str(r["method"]), r["num_experts"] or 0, r["seed"] or 0))

    if len(rows) > MAX_RAW_ROW - FIRST_DATA_ROW:
        print(f"  !! {len(rows)} runs exceeds the bounded formula range "
              f"({MAX_RAW_ROW}). Raise MAX_RAW_ROW and rebuild, or rows past "
              f"{MAX_RAW_ROW} will be invisible to every summary formula.")

    for n, r in enumerate(rows):
        row = FIRST_DATA_ROW + n
        for name in RAW_COLUMNS:
            if name in ("run_id", "mse_sq", "mae_sq"):
                continue
            v = r.get(name, "")
            c = ws.cell(row=row, column=RAW_COLUMNS.index(name) + 1,
                        value="" if v is None else v)
            c.font = Font(name="Arial", size=10)
            if name in ("mse", "mae", "rmse", "gate_w_std"):
                c.number_format = MSE_FMT
        ws.cell(row=row, column=1, value=n + 1)
        ws.cell(row=row, column=RAW_COLUMNS.index("mse_sq") + 1,
                value=f"={COL['mse']}{row}^2").number_format = MSE_FMT
        ws.cell(row=row, column=RAW_COLUMNS.index("mae_sq") + 1,
                value=f"={COL['mae']}{row}^2").number_format = MSE_FMT

    ws.freeze_panes = "A2"
    for name, w in {"git_sha": 42, "run_dir": 58, "model_id": 24, "mtime": 20,
                    "domain_sheet": 14, "method": 17, "model": 13, "tsf_t": 10,
                    "source_csv": 20, "inv_var_norm": 13}.items():
        ws.column_dimensions[COL[name]].width = w
    print(f"  Raw_Runs: {len(rows)} rows, {len(RAW_COLUMNS)} columns")
    return rows


def audit_layout(wb):
    ws = wb[RAW]
    bad = [f"col {CL(i)}: expected {n!r}, found {ws.cell(row=1, column=i).value!r}"
           for i, n in enumerate(RAW_COLUMNS, start=1)
           if ws.cell(row=1, column=i).value != n]
    if bad:
        raise SystemExit("Raw_Runs layout drifted:\n  " + "\n  ".join(bad))


# ==========================================================================
# Formulas
# ==========================================================================
def _rng(col):
    return f"{RAW}!${COL[col]}${FIRST_DATA_ROW}:${COL[col]}${MAX_RAW_ROW}"


def _crit(filters):
    parts = []
    for col, val in filters:
        lit = val if isinstance(val, int) else f'"{val}"'
        parts.append(f"{_rng(col)},{lit}")
    return ",".join(parts)


def f_mean(metric, filters):
    return f'=IFERROR(AVERAGEIFS({_rng(metric)},{_crit(filters)}),"")'


def f_count(filters):
    return f"=COUNTIFS({_crit(filters)})"


def f_std(metric, filters):
    crit = _crit(filters)
    n = f"COUNTIFS({crit})"
    mean = f"AVERAGEIFS({_rng(metric)},{crit})"
    ssq = f"SUMIFS({_rng(metric + '_sq')},{crit})"
    return f'=IF({n}<2,"",IFERROR(SQRT(MAX(0,({ssq}-{n}*({mean})^2))/({n}-1)),""))'


def f_delta(a, b):
    return f'=IF(OR({a}="",{b}=""),"",{a}-{b})'


def resolve(runs, metric, filters):
    """Python mirror of the Excel filter -- this is what makes the audit real."""
    sel = []
    for r in runs:
        ok = True
        for col, val in filters:
            got = r.get(col, "")
            if isinstance(val, int):
                if got != val:
                    ok = False
                    break
            elif str(got) != str(val):
                ok = False
                break
        if ok and r.get(metric) is not None and r.get(metric) != "":
            sel.append(float(r[metric]))
    return sel


def _hdr(ws, row, col, text, width=None, claim=False):
    c = ws.cell(row=row, column=col, value=text)
    c.font = REPRO_HDR_FONT
    c.fill = CLAIM_HDR_FILL if claim else REPRO_HDR_FILL
    c.alignment = Alignment(horizontal="center", wrap_text=True)
    if width:
        ws.column_dimensions[CL(col)].width = width


def _put(ws, row, col, formula, fmt=MSE_FMT, fill=True, claim=False):
    c = ws.cell(row=row, column=col, value=formula)
    c.font = REPRO_VAL_FONT if fill else DELTA_FONT
    if fill:
        c.fill = CLAIM_VAL_FILL if claim else REPRO_VAL_FILL
    c.number_format = fmt
    return c


def _banner(ws, row, col, text, color="FF1E5B3A"):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name="Arial", size=9, italic=True, color=color)


# ==========================================================================
# Pairwise_Baselines -- THE DELIVERABLE
# ==========================================================================
def build_pairwise(wb, runs, audit):
    """Published: A Domain, B Freq, C Horizon, D TSF-N, E TSF-T,
                  F G1 FIXED (TimeMMD), G G3 ATTN (GMM-TS), H..K our Phase-2 gates."""
    if "Pairwise_Baselines" not in wb.sheetnames:
        print("  ! Pairwise_Baselines sheet not found, skipped")
        return
    ws = wb["Pairwise_Baselines"]

    T = 20   # first reproduced column
    cols = {
        "g1": T, "g3": T + 1, "g4": T + 2,
        "d_g1": T + 4, "d_g3": T + 5,
        "claim": T + 7, "vs_g1": T + 8, "wins": T + 9,
        "harness": T + 11, "n": T + 12,
    }

    _banner(ws, 4, T,
            "REPRODUCED BY US — live formulas over Raw_Runs. The published G1/G3 "
            "columns to the left are untouched.")
    _hdr(ws, 5, cols["g1"], "R: G1 FIXED\n(TimeMMD, ours)", width=14)
    _hdr(ws, 5, cols["g3"], "R: G3 ATTN\n(GMM-TS, ours)", width=14)
    _hdr(ws, 5, cols["g4"], "R: G4 IV\n(MM-MoGU, ours)", width=14)
    _hdr(ws, 5, cols["d_g1"], "Δ G1\nours − pub", width=11)
    _hdr(ws, 5, cols["d_g3"], "Δ G3\nours − pub", width=11)
    _hdr(ws, 5, cols["claim"], "G4 − our G3\n(THE CLAIM)", width=14, claim=True)
    _hdr(ws, 5, cols["vs_g1"], "G4 − pub G1", width=12, claim=True)
    _hdr(ws, 5, cols["wins"], "G4 beats\nour G3?", width=11, claim=True)
    _hdr(ws, 5, cols["harness"], "harness check\n(G3 vs pub)", width=13)
    _hdr(ws, 5, cols["n"], "n runs", width=8)

    written = 0
    for r in range(6, ws.max_row + 1):
        domain, horizon = ws[f"A{r}"].value, ws[f"C{r}"].value
        tsf_n, tsf_t = ws[f"D{r}"].value, ws[f"E{r}"].value
        # Data rows only: the sheet also holds roll-ups and prose notes.
        if not isinstance(horizon, (int, float)) or not domain or not tsf_n or not tsf_t:
            continue
        domain, tsf_n, tsf_t = str(domain).strip(), str(tsf_n).strip(), str(tsf_t).strip()
        pl = int(horizon)

        base = [("domain_sheet", domain), ("model", tsf_n),
                ("tsf_t", tsf_t), ("pred_len", pl)]
        specs = [
            ("g1", base + [("method", "G1_FIXED_multi")], "F"),
            ("g3", base + [("method", "G3_ATTN_direct")], "G"),
            ("g4", base + [("method", "G4_IV")], None),
        ]
        for key, filters, pub in specs:
            _put(ws, r, cols[key], f_mean("mse", filters),
                 claim=(key == "g4"))
            if audit is not None:
                audit.append((f"Pairwise!{CL(cols[key])}{r}",
                              f"{domain} h={pl} {tsf_n}x{tsf_t} {key.upper()}",
                              resolve(runs, "mse", filters),
                              ws[f"{pub}{r}"].value if pub else None))

        g1c, g3c, g4c = (f"{CL(cols['g1'])}{r}", f"{CL(cols['g3'])}{r}",
                         f"{CL(cols['g4'])}{r}")
        _put(ws, r, cols["d_g1"], f_delta(g1c, f"F{r}"), fmt=DELTA_FMT, fill=False)
        _put(ws, r, cols["d_g3"], f_delta(g3c, f"G{r}"), fmt=DELTA_FMT, fill=False)

        # The apples-to-apples claim: our gate against our own reproduction of
        # the learned gate, in one codebase under one training recipe. Negative
        # is better (lower MSE).
        _put(ws, r, cols["claim"], f_delta(g4c, g3c), fmt=DELTA_FMT, claim=True)
        _put(ws, r, cols["vs_g1"], f_delta(g4c, f"F{r}"), fmt=DELTA_FMT, claim=True)
        ws.cell(row=r, column=cols["wins"], value=(
            f'=IF(OR({g4c}="",{g3c}=""),"",IF({g4c}<{g3c},"yes","no"))')
        ).font = STATUS_FONT

        # Harness check: if our G3 does not reproduce the published G3, the
        # G4-vs-G3 comparison is still internally valid but the absolute numbers
        # are not comparable to the paper. Say which.
        ws.cell(row=r, column=cols["harness"], value=(
            f'=IF({g3c}="","not run",IF(G{r}="","no target",'
            f'IF(ABS(({g3c}-G{r})/G{r})<={TOL_MATCH},"match",'
            f'IF(ABS(({g3c}-G{r})/G{r})<={TOL_CLOSE},"close","FAIL"))))')
        ).font = STATUS_FONT
        counts = [f"COUNTIFS({_crit(base + [('method', m)])})"
                  for m in ("G1_FIXED_multi", "G3_ATTN_direct", "G4_IV")]
        ws.cell(row=r, column=cols["n"], value="=" + "+".join(counts))
        written += 1

    _banner(ws, 3, T,
            "Never average MSE across domains — scales differ by three orders of "
            "magnitude (Agriculture ~0.06 vs Security ~100-130). Compare within "
            "domain, per horizon, per expert pair.", color="FF8A2A2A")
    print(f"  Pairwise_Baselines: {written} data rows wired up")


# ==========================================================================
# MoGU_ETT
# ==========================================================================
def build_mogu_ett(wb, runs, audit):
    if "MoGU_ETT" not in wb.sheetnames:
        print("  ! MoGU_ETT sheet not found, skipped")
        return
    ws = wb["MoGU_ETT"]
    ws["A2"] = ("Source: arXiv:2510.07459v2, Tables 1, 2, 4. Lookback 96. "
                "VERIFIED against the paper PDF (Tables 1, 2, 4 match cell for cell) "
                "— confidence upgraded from Medium to High.")
    ws["A2"].font = Font(name="Arial", size=9, italic=True, color="FF1E5B3A")
    _t1(ws, runs, audit)
    _t2(ws, runs, audit)
    _t4(ws, runs, audit)


def _t1(ws, runs, audit):
    rows = {"ETTh1": 7, "ETTh2": 8, "ETTm1": 9, "ETTm2": 10}
    specs = [("Single", "B", "single", 1),
             ("MoE-2", "C", "MoE", 2), ("MoE-3", "D", "MoE", 3),
             ("MoE-4", "E", "MoE", 4), ("MoE-5", "F", "MoE", 5),
             ("MoGU-2", "G", "MoGU", 2), ("MoGU-3", "H", "MoGU", 3),
             ("MoGU-4", "I", "MoGU", 4), ("MoGU-5", "J", "MoGU", 5)]
    R0 = 16
    D0 = R0 + len(specs) + 1
    EX = D0 + len(specs) + 1

    _banner(ws, 5, R0, "REPRODUCED — live over Raw_Runs. Δ = reproduced − published. "
                       "Tolerance |Δ|/published: ≤2% match, ≤5% close, else FAIL.")
    for i, (label, _p, _c, _n) in enumerate(specs):
        _hdr(ws, 6, R0 + i, f"R: {label}", width=10)
        _hdr(ws, 6, D0 + i, f"Δ {label}", width=9)
    for j, (label, w) in enumerate([("|%Δ| MoGU-3", 11), ("STATUS", 11), ("n runs", 8),
                                    ("min gate std", 12), ("MoE worse than single?", 20)]):
        _hdr(ws, 6, EX + j, label, width=w)

    for dset, r in rows.items():
        counts = []
        for i, (label, pub, cfg, ne) in enumerate(specs):
            filters = [("dataset", dset), ("model", "iTransformer"),
                       ("pred_len", 96), ("method", cfg), ("num_experts", ne)]
            rc, dc = CL(R0 + i), CL(D0 + i)
            _put(ws, r, R0 + i, f_mean("mse", filters))
            _put(ws, r, D0 + i, f_delta(f"{rc}{r}", f"{pub}{r}"), fmt=DELTA_FMT, fill=False)
            counts.append(f"COUNTIFS({_crit(filters)})")
            if audit is not None:
                audit.append((f"MoGU_ETT!{rc}{r}", f"T1 {dset} {label}",
                              resolve(runs, "mse", filters), ws[f"{pub}{r}"].value))

        # Anchored on MoGU-3: Excel's MAX raises #VALUE! on a cell holding "",
        # so a max across all nine would blank the verdict during a partial sweep.
        m3r, m3d = f"{CL(R0 + 6)}{r}", f"{CL(D0 + 6)}{r}"
        ws.cell(row=r, column=EX,
                value=f'=IF(OR({m3r}="",H{r}=""),"",ABS({m3d}/H{r}))').number_format = PCT_FMT
        g = CL(EX)
        ws.cell(row=r, column=EX + 1, value=(
            f'=IF({m3r}="","not run",IF({g}{r}<={TOL_MATCH},"match",'
            f'IF({g}{r}<={TOL_CLOSE},"close","FAIL")))')).font = STATUS_FONT
        ws.cell(row=r, column=EX + 2, value=f'={"+".join(counts)}')
        mogu = [("dataset", dset), ("model", "iTransformer"),
                ("pred_len", 96), ("method", "MoGU")]
        ws.cell(row=r, column=EX + 3,
                value=f'=IFERROR(MINIFS({_rng("gate_w_std")},{_crit(mogu)}),"")'
                ).number_format = MSE_FMT
        exp = "TRUE" if dset in ("ETTh2", "ETTm1") else "either"
        moe3, single = f"{CL(R0 + 2)}{r}", f"{CL(R0)}{r}"
        ws.cell(row=r, column=EX + 4, value=(
            f'=IF(OR({moe3}="",{single}=""),"not run",'
            f'IF({moe3}>{single},"MoE worse (expected: {exp})",'
            f'"MoE better (expected: {exp})"))'))

    _banner(ws, 11, R0, "Sanity beats point values: MoE MUST be worse than a single "
                        "expert on ETTh2 and ETTm1. THE GATE: ETTh1 MoGU-3 ≈ 0.380.")


def _t2(ws, runs, audit):
    specs = [("iT MoE MAE", "C", "iTransformer", "MoE", "mae"),
             ("iT MoE MSE", "D", "iTransformer", "MoE", "mse"),
             ("iT MoGU MAE", "E", "iTransformer", "MoGU", "mae"),
             ("iT MoGU MSE", "F", "iTransformer", "MoGU", "mse"),
             ("PT MoE MAE", "G", "PatchTST", "MoE", "mae"),
             ("PT MoE MSE", "H", "PatchTST", "MoE", "mse"),
             ("PT MoGU MAE", "I", "PatchTST", "MoGU", "mae"),
             ("PT MoGU MSE", "J", "PatchTST", "MoGU", "mse")]
    R0 = 16
    D0 = R0 + len(specs) + 1
    EX = D0 + len(specs) + 1

    _banner(ws, 13, R0, f"REPRODUCED (num_experts={DEFAULT_NE}). The paper does not "
                        "state Table 2's expert count; 3 is what its Table 1 "
                        "highlights and the repo defaults to.")
    for i, (label, *_rest) in enumerate(specs):
        _hdr(ws, 14, R0 + i, f"R: {label}", width=11)
        _hdr(ws, 14, D0 + i, f"Δ {label}", width=10)
    _hdr(ws, 14, EX, "|%Δ| iT MoGU MSE", width=15)
    _hdr(ws, 14, EX + 1, "STATUS", width=11)
    _hdr(ws, 14, EX + 2, "n runs", width=8)

    for r in range(15, 43):
        label, horizon = ws[f"A{r}"].value, ws[f"B{r}"].value
        if not label or horizon is None:
            continue
        dset = DATASET_ALIAS.get(str(label).strip())
        if dset is None:
            continue
        pl, counts = int(horizon), []
        for i, (name, pub, model, cfg, metric) in enumerate(specs):
            filters = [("dataset", dset), ("model", model), ("pred_len", pl),
                       ("method", cfg), ("num_experts", DEFAULT_NE)]
            rc, dc = CL(R0 + i), CL(D0 + i)
            _put(ws, r, R0 + i, f_mean(metric, filters))
            _put(ws, r, D0 + i, f_delta(f"{rc}{r}", f"{pub}{r}"), fmt=DELTA_FMT, fill=False)
            counts.append(f"COUNTIFS({_crit(filters)})")
            if audit is not None:
                audit.append((f"MoGU_ETT!{rc}{r}", f"T2 {label} h={pl} {name}",
                              resolve(runs, metric, filters), ws[f"{pub}{r}"].value))
        ar, ad = f"{CL(R0 + 3)}{r}", f"{CL(D0 + 3)}{r}"
        ws.cell(row=r, column=EX,
                value=f'=IF(OR({ar}="",F{r}=""),"",ABS({ad}/F{r}))').number_format = PCT_FMT
        g = CL(EX)
        ws.cell(row=r, column=EX + 1, value=(
            f'=IF({ar}="","not run",IF({g}{r}<={TOL_MATCH},"match",'
            f'IF({g}{r}<={TOL_CLOSE},"close","FAIL")))')).font = STATUS_FONT
        ws.cell(row=r, column=EX + 2, value=f'=({"+".join(counts)})/2')

    _banner(ws, 43, R0, "MoGU wins 18 of 32 MSE settings on iTransformer and 19 of 32 "
                        "on PatchTST — not all. (The often-quoted 21/32 is the MAE "
                        "count; Num. Wins reads 4 9 21 18 | 5 8 21 19.) A correct "
                        "reproduction reproduces the losses too.")


def _t4(ws, runs, audit):
    rows = {"ETTh1": 47, "ETTh2": 48, "ETTm1": 49, "ETTm2": 50}
    R0 = 16
    _banner(ws, 45, R0, "REPRODUCED — std is the live identity "
                        "sqrt((SUM(x²)−n·mean²)/(n−1)); Excel has no STDEVIFS. "
                        "Seeds 2351-2355.")
    for i, label in enumerate(["R: MoE MSE", "R: MoE MSE std",
                               "R: MoGU MSE", "R: MoGU MSE std"]):
        _hdr(ws, 46, R0 + i, label, width=13)
    for j, (label, w) in enumerate([("n seeds MoE", 11), ("n seeds MoGU", 12)]):
        _hdr(ws, 46, R0 + 4 + j, label, width=w)
    for j, (label, w) in enumerate([("Δ MoGU MSE", 11), ("Δ MoGU std", 11),
                                    ("std ratio MoE/MoGU", 17), ("STATUS", 11)]):
        _hdr(ws, 46, R0 + 7 + j, label, width=w)

    for dset, r in rows.items():
        moe = [("dataset", dset), ("model", "iTransformer"), ("pred_len", 96),
               ("method", "MoE"), ("num_experts", DEFAULT_NE)]
        mogu = [("dataset", dset), ("model", "iTransformer"), ("pred_len", 96),
                ("method", "MoGU"), ("num_experts", DEFAULT_NE)]
        _put(ws, r, R0 + 0, f_mean("mse", moe))
        _put(ws, r, R0 + 1, f_std("mse", moe))
        _put(ws, r, R0 + 2, f_mean("mse", mogu))
        _put(ws, r, R0 + 3, f_std("mse", mogu))
        ws.cell(row=r, column=R0 + 4, value=f_count(moe))
        ws.cell(row=r, column=R0 + 5, value=f_count(mogu))
        mean_c, std_c = f"{CL(R0 + 2)}{r}", f"{CL(R0 + 3)}{r}"
        _put(ws, r, R0 + 7, f_delta(mean_c, f"E{r}"), fmt=DELTA_FMT, fill=False)
        _put(ws, r, R0 + 8, f_delta(std_c, f"G{r}"), fmt=DELTA_FMT, fill=False)
        ws.cell(row=r, column=R0 + 9, value=(
            f'=IF(OR({CL(R0 + 1)}{r}="",{std_c}="",{std_c}=0),"",'
            f'{CL(R0 + 1)}{r}/{std_c})')).number_format = "0.00"
        d = f"{CL(R0 + 7)}{r}"
        ws.cell(row=r, column=R0 + 10, value=(
            f'=IF({mean_c}="","not run",IF(E{r}="","no target",'
            f'IF(ABS({d}/E{r})<={TOL_MATCH},"match",'
            f'IF(ABS({d}/E{r})<={TOL_CLOSE},"close","FAIL"))))')).font = STATUS_FONT
        if audit is not None:
            audit.append((f"MoGU_ETT!{CL(R0)}{r}", f"T4 {dset} MoE mean",
                          resolve(runs, "mse", moe), ws[f"C{r}"].value))
            audit.append((f"MoGU_ETT!{mean_c}", f"T4 {dset} MoGU mean",
                          resolve(runs, "mse", mogu), ws[f"E{r}"].value))

    _banner(ws, 51, R0, "MoGU's advantage is as much about variance as mean: its seed "
                        "std should be 2-4x tighter than MoE's. The 'std ratio' column "
                        "checks that claim for free.")


# ==========================================================================
def report_audit(audit, n_runs):
    print()
    print("=" * 74)
    print("FORMULA AUDIT")
    print("=" * 74)
    print("No LibreOffice here, so formula cells carry no cached values until Excel")
    print("opens the file. Every filter written into a formula was instead resolved")
    print("against the source rows in Python. This is what each formula WILL show.")
    print()
    resolved = [a for a in audit if a[2]]
    empty = [a for a in audit if not a[2]]
    print(f"formulas written and checked: {len(audit)}")
    print(f"  resolve to data:            {len(resolved)}")
    print(f"  resolve to empty (not run): {len(empty)}")
    print(f"  source rows available:      {n_runs}")

    if resolved:
        print()
        print(f"{'cell':<22}{'what':<38}{'n':>3}  {'repro':>9}  {'published':>10}  "
              f"{'Δ':>9}  verdict")
        print("-" * 110)
        fails = 0
        for cell, what, vals, pub in sorted(resolved, key=lambda a: a[1]):
            mean = sum(vals) / len(vals)
            if isinstance(pub, (int, float)) and pub:
                d = mean - pub
                rel = abs(d / pub)
                verdict = ("match" if rel <= TOL_MATCH
                           else "close" if rel <= TOL_CLOSE else "FAIL")
                fails += verdict == "FAIL"
                print(f"{cell:<22}{what:<38}{len(vals):>3}  {mean:>9.4f}  "
                      f"{pub:>10.4f}  {d:>+9.4f}  {verdict}")
            else:
                print(f"{cell:<22}{what:<38}{len(vals):>3}  {mean:>9.4f}  "
                      f"{'—':>10}  {'—':>9}  no target")
        if fails:
            print()
            print(f"!! {fails} cell(s) would read FAIL (>{TOL_CLOSE:.0%} from published).")
            print("!! Report the honest delta. Do not widen the tolerance to make it pass.")
    if empty and len(empty) < len(audit):
        print()
        print(f"{len(empty)} formula(s) have no matching runs yet — expected while the")
        print("sweep is partial. They fill in as rows are added to Raw_Runs.")
    print("=" * 74)


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--base", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--mogu", nargs="*", default=[],
                    help="CSVs from moe_unc_tsf/scripts/repro/collect_mogu.py")
    ap.add_argument("--timemmd", nargs="*", default=[],
                    help="CSVs from MM-TSFlib/tools/collect_results.py")
    ap.add_argument("--gmmts", nargs="*", default=[],
                    help="CSVs from gmmts_lib/scripts/repro/collect_gmmts.py")
    ap.add_argument("--audit-only", action="store_true")
    args = ap.parse_args()

    if os.path.abspath(args.base) == os.path.abspath(args.out):
        sys.exit("refusing to overwrite the base workbook — give --out a new path. "
                 "v2 is the locked published reference.")

    print("loading run CSVs")
    raw = load_runs(args.mogu + args.timemmd + args.gmmts)
    if not raw:
        print("  (none yet — the workbook is built with empty repro columns)")

    print(f"\nopening {args.base}")
    wb = openpyxl.load_workbook(args.base)
    print(f"  sheets: {', '.join(wb.sheetnames)}")

    print("\nwriting Raw_Runs")
    runs = write_raw_runs(wb, raw)
    audit_layout(wb)
    print("  layout check: column letters match headers")

    audit = []
    print("\nwiring Pairwise_Baselines (the deliverable)")
    build_pairwise(wb, runs, audit)
    print("\nwiring MoGU_ETT")
    build_mogu_ett(wb, runs, audit)

    report_audit(audit, len(runs))

    if args.audit_only:
        print("\n--audit-only: no file written.")
        return 0
    wb.save(args.out)
    print(f"\nwrote {args.out}")
    print("\nPublished (teal) columns were not modified. Open the file in Excel or")
    print("Sheets once to populate the formula values.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
