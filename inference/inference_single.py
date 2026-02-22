import os
import json
import argparse
import pandas as pd
from tqdm import tqdm

from llm_api import get_llm_response
from prompt_format import PROMPT_FORMAT_SINGLE
from code_exec import get_exec_client, extract_code, exec_code


# Tasks in verified_400 with bare naming (initial.xlsx / golden.xlsx, no ID prefix)
BARE_NAMING_IDS = {"13284", "32023", "32789", "56274", "58109"}

# Tasks in verified_400 with mismatched IDs in golden filenames
# Maps task_id -> actual golden file ID
MISMATCHED_IDS = {"42930": "43930"}


def get_input_filename(task_id, dataset):
    """Get the input filename for a task, handling naming variations."""
    if dataset.startswith("spreadsheetbench_verified_400"):
        if str(task_id) in BARE_NAMING_IDS:
            return "initial.xlsx"
        return f"1_{task_id}_init.xlsx"
    return f"1_{task_id}_input.xlsx"


def gen_file_content(input_file, row_count):
    excel_file = pd.ExcelFile(input_file)
    sheet_names = excel_file.sheet_names
    excel_data = {}

    for sheet_name in sheet_names:
        df = excel_file.parse(sheet_name)
        n = row_count if df.shape[0] > row_count else df.shape[0]
        excel_data[sheet_name] = df.head(n).to_string()

    final_str = ""
    for sheet_name, sheet_str in excel_data.items():
        final_str += f"Sheet Name: {sheet_name}\n"
        final_str += sheet_str + "\n"
        final_str += "-" * 50 + "\n"

    return final_str


def gen_solution(opt):
    dataset_path = os.path.abspath(f'../data/{opt.dataset}')
    with open(f'{dataset_path}/dataset.json', 'r') as fp:
        dataset = json.load(fp)

    # check if output file folder exists
    output_file_path = f'{dataset_path}/outputs'
    if not os.path.exists(output_file_path):
        os.makedirs(output_file_path)
        os.chmod(output_file_path, 0o777)

    # check if output file folder of the model exists
    output_file_path = f'{output_file_path}/single_{opt.model}'
    if not os.path.exists(output_file_path):
        os.makedirs(output_file_path)
        os.chmod(output_file_path, 0o777)

    # create code execution client
    client = get_exec_client(opt.code_exec_url, opt.conv_id)

    for data in tqdm(dataset):
        try:
            task_id = data['spreadsheet_path'].lstrip('spreadsheet/')
            file_name = get_input_filename(task_id, opt.dataset)

            input_path = f"/mnt/data/{data['spreadsheet_path']}/{file_name}"
            # Output always uses consistent _output.xlsx naming
            output_path = f"/mnt/data/outputs/single_{opt.model}/1_{task_id}_output.xlsx"

            find_input_path = f"{dataset_path}/{data['spreadsheet_path']}/{file_name}"
            file_content = gen_file_content(find_input_path, opt.row)
            prompt = ""
            prompt = PROMPT_FORMAT_SINGLE.format_map({
                'instruction': data['instruction'],
                'spreadsheet_path': input_path,
                'spreadsheet_content' : file_content,
                'instruction_type': data['instruction_type'],
                'answer_position': data['answer_position'],
                'output_path': output_path
            })
            messages = [prompt]
            response = get_llm_response(messages, opt)
            messages.append(response)
            try:
                exec_result = exec_code(client, extract_code(response))
            except Exception as e:
                exec_result = 'Error occur when running code.'
            messages.append(exec_result)
            conv_result = {
                'id': data['id'],
                'instruction_type': data['instruction_type'],
                'conversation': messages,
                'solution': extract_code(response)
            }
        except Exception as e:
            print(str(e))
            conv_result = {
                'id': data['id'],
                'instruction_type': data['instruction_type'],
                'conversation': "",
                'solution': ""
            }
            with open(f'log/single_{opt.model}.jsonl', 'a+') as f:
                f.write(json.dumps(data, ensure_ascii=False) + '\n')
        with open(f'outputs/conv_single_{opt.model}.jsonl', 'a+') as fp:
            fp.write(json.dumps(conv_result, ensure_ascii=False) + '\n')


def run_solution(opt):
    client = get_exec_client(opt.code_exec_url, opt.conv_id)
    dataset_path = os.path.abspath(f'../data/{opt.dataset}')
    with open(f'{dataset_path}/outputs/conv_single_{opt.model}.jsonl', 'r') as fp:
        conv_records = [json.loads(line) for line in fp.readlines()]
    for conv in tqdm(conv_records):
        try:
            for idx in range(2, opt.num_test_cases + 1):
                input_file = f"{idx}_{conv['id']}_input.xlsx"
                output_file = f"{idx}_{conv['id']}_output.xlsx"
                solution = conv['solution'].replace(f"1_{conv['id']}_input.xlsx", input_file)
                solution = solution.replace(f"1_{conv['id']}_output.xlsx", output_file)
                exec_result = exec_code(client, solution)
        except Exception as e:
            print(e)


def parse_option():
    parser = argparse.ArgumentParser("command line arguments for generation.")

    parser.add_argument('--model', type=str, help='model name')
    parser.add_argument('--api_key', type=str, default="", help='the api key of model')
    parser.add_argument('--base_url', type=str, default="", help='the base url of model')
    parser.add_argument('--dataset', type=str, default="sample_data_200", help='dataset name')
    parser.add_argument('--code_exec_url', type=str, default="http://localhost:8081/execute", help='code execution docker url')
    parser.add_argument('--conv_id', type=str, default="EVAL", help='code execution conversation id')
    parser.add_argument('--row', type=int, default=5, help='the number of rows provided in the prompt')
    parser.add_argument('--num-test-cases', type=int, default=3,
                        help='number of test cases per task (3 for sample_data_200, 1 for verified_400)')
    opt = parser.parse_args()

    return opt


if __name__ == '__main__':
    opt = parse_option()
    print(opt)

    gen_solution(opt)
    # Only run additional test cases if num_test_cases > 1
    # verified_400 has only 1 test case, so run_solution is unnecessary
    if opt.num_test_cases > 1:
        run_solution(opt)
