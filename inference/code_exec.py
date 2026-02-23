import re

from jupyter_kernel_cli import ClientJupyterKernel

def get_exec_client(url, conv_id):
    client = ClientJupyterKernel(url, conv_id)
    return client

def extract_code(response):
    # 1. Try markdown code blocks (original behavior)
    if '```python' in response:
        code = response[response.find('```python') + len('```python'):]
        code = code[:code.find('```')].lstrip('\n').rstrip('\n')
        return code
    # 2. Try XML function_calls format (Anthropic models via OpenAI-compat API)
    match = re.search(r'<parameter\s+name="code">\s*(.*?)\s*</parameter>', response, re.DOTALL)
    if match:
        return match.group(1).strip()
    # 3. Fallback: treat entire response as code
    return response

def exec_code(client, code):
    res = client.execute(code)
    if res.find('-----') != -1:
        tracebacks = res.split('\n\n\n\n')
        error_feedback = ''
        for t in tracebacks:
            if t.find('Error') != -1:
                error_feedback += t + '\n'
                break
        for t in tracebacks:
            if len(t) >= len('Cell') and t[:len('Cell')] == 'Cell':
                error_feedback += t
                break
        error_feedback += tracebacks[-1]
        return error_feedback
    else:
        return res
