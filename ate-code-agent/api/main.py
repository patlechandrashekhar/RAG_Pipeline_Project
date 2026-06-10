"""FastAPI app for ATE Code Agent single-shot queries."""

import os
import ssl
import httpx

from fastapi import FastAPI, HTTPException

from agent import run_single_query
from api.models import QueryRequest, QueryResponse

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context


app = FastAPI(title="ATE Code Agent API", version="0.1.0")


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/query", response_model=QueryResponse)
async def query(req: QueryRequest) -> QueryResponse:
    try:
        response = await run_single_query(req.program_path, req.task, approved=req.approved)
        return QueryResponse(response=response)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))
