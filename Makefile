PYTHON ?= python

.PHONY: setup run test check

setup:
	$(PYTHON) -m pip install --upgrade pip
	$(PYTHON) -m pip install -r requirements.txt
	$(PYTHON) -m pip install pytest

run:
	$(PYTHON) -m streamlit run app/streamlit_app.py

test:
	$(PYTHON) -m pytest tests -q

check:
	$(PYTHON) -c "import ast,pathlib; root=pathlib.Path('.'); files=[root/'app'/'streamlit_app.py']+list((root/'src'/'page_indexing_rag').glob('*.py')); [ast.parse(p.read_text(encoding='utf-8', errors='replace')) for p in files]; print('AST_OK')"

