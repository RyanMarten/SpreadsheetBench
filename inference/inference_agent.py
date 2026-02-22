"""
Agent-style inference using claude-code CLI for SpreadsheetBench parity experiments.

For each task in the dataset, this script:
1. Creates a temp working directory with the input spreadsheet
2. Runs claude-code with the task instruction
3. Collects the output spreadsheet
4. After all tasks: run LibreOffice recalculation + evaluation.py

This is "Experiment 3" in the parity plan — agent-style inference in the fork
environment, to compare against LLM-style (Exp 1/2) and Harbor (Exp 4).

Usage:
    cd inference/
    python inference_agent.py --dataset spreadsheetbench_verified_400

    # Limit to first N tasks for testing
    python inference_agent.py --dataset spreadsheetbench_verified_400 --limit 5

    # Resume from a specific task
    python inference_agent.py --dataset spreadsheetbench_verified_400 --resume-from 42930

Prerequisites:
    - claude-code CLI installed (npm install -g @anthropic-ai/claude-code)
    - ANTHROPIC_API_KEY environment variable set
    - openpyxl, pandas installed
"""

import os
import json
import shutil
import argparse
import subprocess
import tempfile
from pathlib import Path
from tqdm import tqdm


# Same naming constants as inference_single.py
BARE_NAMING_IDS = {"13284", "32023", "32789", "56274", "58109"}
MISMATCHED_IDS = {"42930": "43930"}


def get_input_filename(task_id, dataset):
    """Get the input filename for a task, handling naming variations."""
    if dataset.startswith("spreadsheetbench_verified_400"):
        if str(task_id) in BARE_NAMING_IDS:
            return "initial.xlsx"
        return f"1_{task_id}_init.xlsx"
    return f"1_{task_id}_input.xlsx"


def build_agent_instruction(data, input_filename, task_id):
    """Build the instruction prompt for the claude-code agent.

    Mirrors the Harbor adapter's instruction.md template to ensure
    Exp 3 and Exp 4 use equivalent instructions.
    """
    instruction = data['instruction']
    answer_position = data['answer_position']

    return f"""You are given a spreadsheet file in the current directory. Your task is to write and execute Python code that manipulates the spreadsheet according to the instructions below.

## Instructions

{instruction}

## Input Files

The following input spreadsheet file is available in the current directory:

- `{input_filename}`

## Output Requirements

Produce an output file named `1_{task_id}_output.xlsx` in the current directory.

## Important Notes

- Use Python with `openpyxl` or `pandas` to manipulate the spreadsheet file. Both libraries are available.
- The answer should be written to the cell(s) at position: `{answer_position}`
- Make sure to save the output file before finishing.
- Write your code and execute it directly. Do not ask for confirmation.
"""


def run_agent_on_task(data, dataset_path, output_dir, opt):
    """Run claude-code on a single task and collect the output."""
    task_id = data['id']
    spreadsheet_dir = f"{dataset_path}/{data['spreadsheet_path']}"
    input_filename = get_input_filename(task_id, opt.dataset)
    input_path = f"{spreadsheet_dir}/{input_filename}"

    if not os.path.exists(input_path):
        print(f"  WARNING: Input file not found: {input_path}")
        return {"id": task_id, "status": "missing_input"}

    output_filename = f"1_{task_id}_output.xlsx"
    output_dest = f"{output_dir}/{output_filename}"

    # Skip if output already exists (for resume support)
    if os.path.exists(output_dest):
        return {"id": task_id, "status": "skipped_existing"}

    # Create a temp working directory
    with tempfile.TemporaryDirectory(prefix=f"ssb_{task_id}_") as tmpdir:
        # Copy input spreadsheet
        shutil.copy2(input_path, f"{tmpdir}/{input_filename}")

        # Build instruction
        instruction = build_agent_instruction(data, input_filename, task_id)

        # Run claude-code
        try:
            result = subprocess.run(
                [
                    "claude",
                    "--print",
                    "--model", opt.model,
                    "--max-turns", str(opt.max_turns),
                    "--dangerously-skip-permissions",
                    instruction,
                ],
                cwd=tmpdir,
                capture_output=True,
                text=True,
                timeout=opt.timeout,
                env={
                    k: v for k, v in os.environ.items()
                    if k not in ("CLAUDECODE", "CLAUDE_CODE")
                },
            )

            agent_output = result.stdout
            agent_stderr = result.stderr
            returncode = result.returncode

        except subprocess.TimeoutExpired:
            return {"id": task_id, "status": "timeout"}
        except Exception as e:
            return {"id": task_id, "status": "error", "error": str(e)}

        # Collect output spreadsheet
        output_in_tmp = f"{tmpdir}/{output_filename}"
        if os.path.exists(output_in_tmp):
            shutil.copy2(output_in_tmp, output_dest)
            return {
                "id": task_id,
                "status": "success",
                "returncode": returncode,
            }
        else:
            # Check if agent wrote to a different location
            xlsx_files = list(Path(tmpdir).glob("*.xlsx"))
            non_input = [f for f in xlsx_files if f.name != input_filename]
            if non_input:
                # Use the first non-input xlsx as output
                shutil.copy2(str(non_input[0]), output_dest)
                return {
                    "id": task_id,
                    "status": "success_alt_name",
                    "actual_name": non_input[0].name,
                    "returncode": returncode,
                }
            return {
                "id": task_id,
                "status": "no_output",
                "returncode": returncode,
                "stdout_tail": agent_output[-500:] if agent_output else "",
                "stderr_tail": agent_stderr[-500:] if agent_stderr else "",
            }


def main():
    parser = argparse.ArgumentParser("Agent-style inference using claude-code CLI")
    parser.add_argument('--dataset', type=str, default='spreadsheetbench_verified_400',
                        help='dataset name')
    parser.add_argument('--model', type=str, default='claude-haiku-4-5-20251001',
                        help='claude-code model')
    parser.add_argument('--max-turns', type=int, default=10,
                        help='max claude-code turns per task')
    parser.add_argument('--timeout', type=int, default=300,
                        help='timeout per task in seconds')
    parser.add_argument('--limit', type=int, default=0,
                        help='limit number of tasks (0 = all)')
    parser.add_argument('--resume-from', type=str, default=None,
                        help='resume from a specific task ID')
    parser.add_argument('--trial-id', type=str, default='1',
                        help='trial identifier (for multiple runs)')
    opt = parser.parse_args()

    dataset_path = os.path.abspath(f'../data/{opt.dataset}')
    with open(f'{dataset_path}/dataset.json', 'r') as fp:
        dataset = json.load(fp)

    # Output directory
    output_dir = f"{dataset_path}/outputs/agent_{opt.model}_trial{opt.trial_id}"
    os.makedirs(output_dir, exist_ok=True)

    # Log file
    log_dir = "outputs"
    os.makedirs(log_dir, exist_ok=True)
    log_path = f"{log_dir}/agent_{opt.model}_trial{opt.trial_id}.jsonl"

    # Apply limits
    if opt.limit > 0:
        dataset = dataset[:opt.limit]

    # Resume support
    resume_active = opt.resume_from is not None
    if resume_active:
        print(f"Will resume from task ID: {opt.resume_from}")

    print(f"Running agent inference on {len(dataset)} tasks")
    print(f"Model: {opt.model}")
    print(f"Output: {output_dir}")
    print(f"Trial: {opt.trial_id}")

    results_summary = {"success": 0, "no_output": 0, "error": 0, "timeout": 0, "skipped": 0}

    for data in tqdm(dataset):
        if resume_active:
            if str(data['id']) == str(opt.resume_from):
                resume_active = False
            else:
                continue

        result = run_agent_on_task(data, dataset_path, output_dir, opt)

        # Track results
        status = result.get("status", "error")
        if status in ("success", "success_alt_name"):
            results_summary["success"] += 1
        elif status == "skipped_existing":
            results_summary["skipped"] += 1
        elif status == "timeout":
            results_summary["timeout"] += 1
        elif status == "no_output":
            results_summary["no_output"] += 1
        else:
            results_summary["error"] += 1

        # Log result
        with open(log_path, 'a') as f:
            f.write(json.dumps(result, ensure_ascii=False) + '\n')

    print(f"\nResults: {json.dumps(results_summary, indent=2)}")
    print(f"\nOutput files saved to: {output_dir}")
    print(f"Log saved to: {log_path}")
    print(f"\nNext steps:")
    print(f"  1. Recalculate formulas:")
    print(f"     bash ../evaluation/recalculate_libreoffice.sh {output_dir}")
    print(f"  2. Evaluate:")
    print(f"     cd ../evaluation && python evaluation.py --model {opt.model}_trial{opt.trial_id} --setting agent --dataset {opt.dataset} --num-test-cases 1")


if __name__ == '__main__':
    main()
