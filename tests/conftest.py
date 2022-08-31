import pytest
from pathlib import Path


@pytest.fixture(scope="session")
def _root():
    # root directory
    return Path(__file__).parent.resolve()

