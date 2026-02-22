#!/bin/bash
# Run single-round inference with Haiku 4.5 on verified_400 dataset.
# Uses Anthropic's OpenAI-compatible endpoint (no changes to llm_api.py needed).
#
# Prerequisites:
#   - ANTHROPIC_API_KEY env var set
#   - Code execution Docker containers running (see code_exec_docker/README.md)
#   - verified_400 dataset extracted in ../data/spreadsheetbench_verified_400/
#
# Usage:
#   cd inference/
#   bash scripts/inference_single_haiku.sh

set -euo pipefail

if [ -z "${ANTHROPIC_API_KEY:-}" ]; then
    echo "ERROR: ANTHROPIC_API_KEY environment variable not set"
    exit 1
fi

python inference_single.py \
    --model claude-haiku-4-5-20251001 \
    --api_key "$ANTHROPIC_API_KEY" \
    --base_url "https://api.anthropic.com/v1/" \
    --dataset spreadsheetbench_verified_400 \
    --num-test-cases 1
