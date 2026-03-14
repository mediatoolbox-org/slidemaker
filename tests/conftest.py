from __future__ import annotations

import os
from pathlib import Path

import pytest
from dotenv import load_dotenv

TESTS_DIR = Path(__file__).resolve().parent
load_dotenv(TESTS_DIR / ".env")


@pytest.fixture(scope="session")
def openai_api_key() -> str | None:
    return os.environ.get("OPENAI_API_KEY")
