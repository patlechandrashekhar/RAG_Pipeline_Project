"""Pydantic models for API requests and responses."""

import os
import ssl
import httpx
from typing import Optional

from pydantic import BaseModel, Field

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context


class QueryRequest(BaseModel):
    program_path: str = Field(..., description="Folder containing IG-XL program")
    task: str = Field(..., description="Engineer task prompt")
    approved: bool = Field(False, description="Set true after engineer approval for write/run actions")


class QueryResponse(BaseModel):
    response: str
    error: Optional[str] = None
