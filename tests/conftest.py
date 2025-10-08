import importlib.util
from pathlib import Path

import pytest


@pytest.fixture(scope="session")
def main_code():
    module_path = Path(__file__).resolve().parents[1] / "main-code.py"
    spec = importlib.util.spec_from_file_location("main_code", module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)  # type: ignore
    return module
