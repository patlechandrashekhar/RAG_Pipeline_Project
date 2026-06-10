from pathlib import Path


def test_expected_directories_exist() -> None:
    root = Path(__file__).resolve().parents[1]
    for rel in ("app", "src", "data", "docs", "scripts", "tests"):
        assert (root / rel).exists(), f"Missing directory: {rel}"

