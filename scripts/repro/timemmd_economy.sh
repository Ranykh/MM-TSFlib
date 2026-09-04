#!/usr/bin/env bash
# =============================================================================
# TimeMMD (G1 FIXED) reproduction -- one important domain, the three expert pairs
#
#   bash scripts/repro/timemmd_economy.sh <gpu_id> [multi|uni|both]
#
# Scope, deliberately small: Economy only, 3 expert pairs, 4 horizons.
# Economy is the domain where the gate matters most -- removing dynamic gating
# costs 12.728 MSE there, the largest degradation in the entire GMM-TS ablation
# ("text is informative and volatile; the gate carries the whole method here").
# If TimeMMD's fixed weighting is going to look weak anywhere, it is here, which
# makes it the honest place to reproduce it rather than the flattering one.
#
# 3 pairs x 4 horizons = 12 multimodal runs, or 24 with the unimodal floor.
#
# -----------------------------------------------------------------------------
# WHAT "TimeMMD G1 FIXED" ACTUALLY IS
# -----------------------------------------------------------------------------
# MM-TSFlib's "learnable" fusion weights are DEAD CODE. exp_long_term_forecasting.py
# builds weight1/weight2 (lines 344-356) and hands them to an optimizer (386-387)
# but never references them in any forward pass. Fusion is therefore the constant
# --prompt_weight, and that constant IS the G1 FIXED baseline.
#
# The corollary is useful: --prompt_weight 0 gives the pure unimodal numeric
# baseline. There is no --unimodal flag and none is needed.
#
# -----------------------------------------------------------------------------
# THE OVERWRITE TRAP -- the reason every run carries a hand-built --model_id
# -----------------------------------------------------------------------------
# run.py's `setting` string omits llm_model, text_len, prompt_weight and
# pool_type. Two runs differing only by LLM land in the SAME results/ folder and
# the later one wins, silently, with nothing in the output to show it happened.
# Every run below therefore encodes those fields in --model_id, in the shape
# tools/collect_results.py parses:
#
#     <Domain>_<uni|multi>_<llm>_tl<text_len>_pw<prompt_weight>_sl<seq_len>_s<seed>
#
# -----------------------------------------------------------------------------
# OTHER TRAPS ALREADY HANDLED HERE
# -----------------------------------------------------------------------------
#   --text_len must be 2, 4 or 6. The columns are Final_Search_{2,4,6} and
#     run.py's default of 3 raises KeyError.
#   --llm_model ClosedLLM needs a Final_Output column. Economy does NOT have one
#     (only Energy, Public_Health and Traffic do), so GPT-3.5 is unavailable here
#     and the text experts are GPT2 and LLAMA2 -- which is exactly the pair set
#     the gmmts sweep uses, so the two tables line up.
#   run.py forces features='S': every Time-MMD task is univariate.
#   exp_basic.py uses the raw --gpu index, so we pin externally and pass nothing.
# =============================================================================
set -uo pipefail

GPU_ID=${1:?usage: timemmd_economy.sh <gpu_id> [multi|uni|both]}
MODE=${2:-both}

export CUDA_VISIBLE_DEVICES="$GPU_ID"

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$REPO_ROOT"

DOMAIN="${DOMAIN:-Economy}"
DATA_FILE="${DATA_FILE:-US_TradeBalance_Month.csv}"
ROOT_PATH="./data/$DOMAIN"

# Same three pairs as the gmmts sweep, so the two tables are directly comparable.
PAIRS="${PAIRS:-DLinear:GPT2 PatchTST:GPT2 DLinear:LLAMA2}"
HORIZONS="${HORIZONS:-6 8 10 12}"
SEEDS="${SEEDS:-2021}"

# Monthly protocol, matching the gmmts monthly example.
SEQ_LEN=8
LABEL_LEN=4
TEXT_LEN=4
POOL_TYPE="avg"
TYPE_TAG="#F#"
# TimeMMD's fixed fusion constant. It is hand-tuned per domain in the original
# work, so treat this as a starting point: if multimodal loses to unimodal on
# most configs, this is the first thing to vary, not the gate.
PROMPT_WEIGHT="${PROMPT_WEIGHT:-0.01}"

LOG_DIR="${LOG_DIR:-$HOME/msc/logs/timemmd/$DOMAIN}"
SKIP_DONE="${SKIP_DONE:-0}"
mkdir -p "$LOG_DIR"

GIT_SHA=$(git rev-parse HEAD 2>/dev/null || echo "NO-GIT")
GIT_BRANCH=$(git rev-parse --abbrev-ref HEAD 2>/dev/null || echo "?")
GIT_DIRTY=$(git status --porcelain 2>/dev/null | head -c1)
MANIFEST="$LOG_DIR/manifest_$(date +%F_%H%M%S).txt"

{
  echo "domain:        $DOMAIN ($DATA_FILE)"
  echo "started:       $(date -Is)"
  echo "gpu:           physical card $GPU_ID"
  echo "git_sha:       $GIT_SHA  (branch $GIT_BRANCH)"
  echo "git_dirty:     $([ -n "$GIT_DIRTY" ] && echo YES || echo no)"
  echo "pairs:         $PAIRS"
  echo "horizons:      $HORIZONS"
  echo "prompt_weight: $PROMPT_WEIGHT"
} | tee "$MANIFEST"

# The mm-mogu branch cannot run this: its run.py has no --prob_expert, but its
# models/iTransformer.py and models/PatchTST.py read configs.prob_expert, so a
# plain run.py dies with AttributeError. Its PatchTST is also modified
# (BatchNorm for LayerNorm, patch_len clamp deleted), so even a working run
# would not be comparable to the published Time-MMD numbers.
if [ "$GIT_BRANCH" = "mm-mogu" ]; then
  echo ""
  echo "!! You are on branch mm-mogu. This will crash with AttributeError, and"
  echo "!! its PatchTST is not upstream PatchTST. Use a branch cut from e789ce7."
  exit 1
fi
if [ -n "$GIT_DIRTY" ]; then
  echo ""
  echo "!! working tree is DIRTY -- results will not be tied to $GIT_SHA."
fi

N_OK=0; N_FAIL=0; N_SKIP=0; FAILED_RUNS=()

# =============================================================================
# run_one <tsf_n> <tsf_t> <pred_len> <seed> <prompt_weight> <uni|multi>
# =============================================================================
run_one () {
  local model=$1 llm=$2 pl=$3 seed=$4 pw=$5 tag=$6
  local rc elapsed start mse

  local model_id="${DOMAIN}_${tag}_${llm}_tl${TEXT_LEN}_pw${pw}_sl${SEQ_LEN}_s${seed}"
  local name="${DOMAIN}_${tag}_${model}_${llm}_pl${pl}_pw${pw}_s${seed}"
  local log="$LOG_DIR/${name}.log"

  if [ "$SKIP_DONE" = "1" ] && [ -f "$log" ] && grep -q "^mse:" "$log"; then
    echo "--- SKIP (already done): $name"; N_SKIP=$((N_SKIP + 1)); return 0
  fi

  echo ""
  echo "=== $name"
  echo "    model_id -> $model_id"
  start=$(date +%s)

  {
    python -u run.py \
      --task_name long_term_forecast \
      --is_training 1 \
      --root_path "$ROOT_PATH" \
      --data_path "$DATA_FILE" \
      --model_id "$model_id" \
      --model "$model" \
      --data custom \
      --seq_len $SEQ_LEN \
      --label_len $LABEL_LEN \
      --pred_len "$pl" \
      --des 'Exp' \
      --seed "$seed" \
      --type_tag "$TYPE_TAG" \
      --text_len $TEXT_LEN \
      --pool_type "$POOL_TYPE" \
      --llm_model "$llm" \
      --prompt_weight "$pw" \
      --huggingface_token "${HUGGINGFACE_TOKEN:-}" \
      --save_name "result_timemmd_${DOMAIN}.txt"
  } 2>&1 | tee "$log"
  rc=${PIPESTATUS[0]}
  elapsed=$(( $(date +%s) - start ))

  if [ "$rc" -eq 0 ] && grep -q "^mse:" "$log"; then
    mse=$(grep "^mse:" "$log" | tail -1)
    echo "    OK  (${elapsed}s)  $mse"
    echo "$name | ${elapsed}s | $mse" >> "$MANIFEST"
    N_OK=$((N_OK + 1))
  else
    echo "    FAILED (rc=$rc, ${elapsed}s) -- see $log"
    echo "$name | ${elapsed}s | FAILED rc=$rc" >> "$MANIFEST"
    FAILED_RUNS+=("$name"); N_FAIL=$((N_FAIL + 1))
  fi
}

sweep () {   # sweep <prompt_weight> <uni|multi>
  local pw=$1 tag=$2
  for seed in $SEEDS; do
    for pair in $PAIRS; do
      local model="${pair%%:*}" llm="${pair##*:}"
      for pl in $HORIZONS; do
        run_one "$model" "$llm" "$pl" "$seed" "$pw" "$tag"
      done
    done
  done
}

case "$MODE" in
  multi) sweep "$PROMPT_WEIGHT" multi ;;
  uni)   sweep 0 uni ;;
  both)
    # The unimodal floor first: it is the cheaper half and it tells you straight
    # away whether text is helping at all in this domain.
    sweep 0 uni
    sweep "$PROMPT_WEIGHT" multi
    ;;
  *) echo "unknown mode '$MODE'. use: multi | uni | both" >&2; exit 1 ;;
esac

echo ""
echo "============================================================"
echo "finished at $(date -Is)"
echo "  ok: $N_OK   skipped: $N_SKIP   failed: $N_FAIL"
for f in "${FAILED_RUNS[@]:-}"; do [ -n "$f" ] && echo "    - $f"; done
echo ""
echo "  manifest: $MANIFEST"
echo "  git_sha:  $GIT_SHA"
echo ""
echo "collect with:"
echo "  python tools/collect_results.py --expect $((N_OK + N_FAIL)) --pivot \\"
echo "      --out ~/msc/logs/timemmd_${DOMAIN}.csv"
echo ""
echo "Time-MMD reports multimodal winning ~95% of configs. A much worse rate in"
echo "the pivot means PROMPT_WEIGHT is mistuned for this domain -- vary that"
echo "before concluding anything about text being uninformative."
echo "============================================================"
