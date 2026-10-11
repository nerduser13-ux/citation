"""Shared pytest fixtures: generated .docx scenarios + the real fixtures."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def _load_generator_module():
    path = REPO_ROOT / "scripts" / "generate_fixtures.py"
    spec = importlib.util.spec_from_file_location("generate_fixtures", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="session")
def gen_fixtures(tmp_path_factory) -> dict[str, Path]:
    """All programmatically generated fixture docx files, name -> path."""
    outdir = tmp_path_factory.mktemp("generated_fixtures")
    return _load_generator_module().generate_all(outdir)


@pytest.fixture(scope="session")
def real_docx() -> Path:
    return REPO_ROOT / "tests" / "fixtures" / "real_vancouver.docx"


@pytest.fixture(scope="session")
def sample_docx() -> Path:
    return REPO_ROOT / "tests" / "fixtures" / "zotero_sample.docx"
