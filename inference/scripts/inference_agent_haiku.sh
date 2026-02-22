#!/bin/bash
# Run agent-style inference with claude-code + Haiku 4.5 on verified_400.
# This is Experiment 3 in the parity plan.
#
# Prerequisites:
#   - claude-code CLI installed (npm install -g @anthropic-ai/claude-code)
#   - ANTHROPIC_API_KEY environment variable set
#   - verified_400 dataset extracted in ../data/spreadsheetbench_verified_400/
#
# Usage:
#   cd inference/
#   bash scripts/inference_agent_haiku.sh           # Trial 1
#   bash scripts/inference_agent_haiku.sh 2         # Trial 2
#   bash scripts/inference_agent_haiku.sh 3         # Trial 3

set -euo pipefail

TRIAL_ID="${1:-1}"

if [ -z "${ANTHROPIC_API_KEY:-}" ]; then
    echo "ERROR: ANTHROPIC_API_KEY environment variable not set"
    exit 1
fi

if ! command -v claude &> /dev/null; then
    echo "ERROR: claude-code CLI not found. Install with: npm install -g @anthropic-ai/claude-code"
    exit 1
fi

echo "=== Agent Inference Trial $TRIAL_ID ==="

python inference_agent.py \
    --dataset spreadsheetbench_verified_400 \
    --model claude-haiku-4-5-20251001 \
    --max-turns 10 \
    --timeout 300 \
    --trial-id "$TRIAL_ID"

echo ""
echo "=== Post-processing: Recalculate formulas ==="
DATASET_PATH=$(cd .. && pwd)/data/spreadsheetbench_verified_400
OUTPUT_DIR="$DATASET_PATH/outputs/agent_claude-haiku-4-5-20251001_trial${TRIAL_ID}"
bash ../evaluation/recalculate_libreoffice.sh "$OUTPUT_DIR"

echo ""
echo "=== Evaluation ==="
cd ../evaluation
python evaluation.py \
    --model "claude-haiku-4-5-20251001_trial${TRIAL_ID}" \
    --setting agent \
    --dataset spreadsheetbench_verified_400 \
    --num-test-cases 1

echo ""
echo "=== Trial $TRIAL_ID Complete ==="
echo "Results: ../outputs/eval_agent_claude-haiku-4-5-20251001_trial${TRIAL_ID}.json"
