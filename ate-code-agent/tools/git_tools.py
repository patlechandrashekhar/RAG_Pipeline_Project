"""
Git backup tools for the ATE Code Agent.
"""

import os
import ssl
import httpx
from datetime import datetime
from pathlib import Path

import git

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context


def git_commit_backup(file_path: str, message: str) -> str:
    """Commit current state of a file before editing."""
    path = Path(file_path).resolve()

    try:
        repo = git.Repo(str(path.parent), search_parent_directories=True)
    except git.InvalidGitRepositoryError:
        repo = git.Repo.init(str(path.parent))
        gitignore = path.parent / ".gitignore"
        if not gitignore.exists():
            gitignore.write_text("*.pyc\n__pycache__/\n.env\n", encoding="utf-8")
        repo.index.add([".gitignore"])

    try:
        repo.index.add([str(path)])
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit = repo.index.commit(f"[ATE Agent Backup] {timestamp} - {message}")
        return commit.hexsha
    except Exception as exc:
        print(f"Warning: Git backup failed: {exc}")
        return "backup-failed"
