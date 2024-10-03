chcp 65001

call "%~dp0net_venv\Scripts\activate"

pip install -r "%~dp0requirements.txt"

py "%~dp0_main.py"

