from pathlib import Path


def test_core_modules_exist() -> None:
    root = Path(__file__).resolve().parents[1] / "src" / "page_indexing_rag"
    expected = {"config.py", "ingestion.py", "retrieval.py", "generation.py"}
    actual = {p.name for p in root.glob("*.py")}
    missing = expected - actual
    assert not missing, f"Missing modules: {sorted(missing)}"

