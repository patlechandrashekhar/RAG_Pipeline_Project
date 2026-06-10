"""Entry point for the ATE Code Agent."""

import os
import ssl
import httpx
import sys
from pathlib import Path

import anyio
from dotenv import load_dotenv

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

from agent import run_interactive_agent


def main() -> None:
    load_dotenv()

    if len(sys.argv) < 2:
        print("Usage: python main.py <path-to-igxl-program.xlsx>")
        print("Example: python main.py C:/programs/MyDevice/MyDevice_TP.xlsx")
        sys.exit(1)

    program_path = Path(sys.argv[1])
    if not program_path.exists():
        print(f"Error: Program file not found: {program_path}")
        sys.exit(1)

    anyio.run(run_interactive_agent, str(program_path.parent))


if __name__ == "__main__":
    main()
