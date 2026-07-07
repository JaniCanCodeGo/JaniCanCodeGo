"""Grant Writing Studio — web server.

FastAPI backend + static accessible frontend. Run with:
    python3 -m uvicorn server:app --reload --port 8000
from the grant_site/ directory, or `python3 server.py`.

Privacy by design: run data lives in memory and generated files on local
disk only; no analytics, no third-party trackers, no cookies. See
static/privacy.html and docs/COMPLIANCE.md.
"""

import os
import threading
import uuid

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from agent import ein, pipeline, trademark

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
STATIC_DIR = os.path.join(BASE_DIR, "static")

app = FastAPI(title="Grant Writing Studio", version="1.0.0",
              description="Agent-driven grant writing: website review, EIN "
                          "lookup/application prep, business plan & "
                          "prospectus, trademark search & TEAS prep.")

RUNS = {}
RUNS_LOCK = threading.Lock()


class ForecastAssumptions(BaseModel):
    base_revenue: float | None = None
    revenue_growth_pct: float | None = None
    base_expenses: float | None = None
    expense_growth_pct: float | None = None
    grant_target: float | None = None
    merch_units: float | None = None
    merch_price: float | None = None
    merch_unit_cost: float | None = None
    merch_growth_pct: float | None = None
    participants: float | None = None
    participant_growth_pct: float | None = None


class RunRequest(BaseModel):
    org_name: str = Field(min_length=1, max_length=200)
    website_url: str = ""
    ein: str = ""
    state: str = ""
    entity_type: str = ""
    contact_email: str = ""
    mission: str = ""
    programs: str = ""
    merchandise: str = ""
    focus_areas: str = ""
    trademark_name: str = ""
    mark_in_use: bool = False
    forecast: ForecastAssumptions = ForecastAssumptions()


@app.post("/api/runs")
def create_run(req: RunRequest):
    run_id = uuid.uuid4().hex[:12]
    out_dir = os.path.join(OUTPUT_DIR, run_id)
    state = {"id": run_id, "status": "running", "steps": [], "result": None}
    with RUNS_LOCK:
        RUNS[run_id] = state

    def progress(step, status, detail=""):
        with RUNS_LOCK:
            state["steps"].append({"step": step, "status": status,
                                   "detail": detail})

    def work():
        payload = req.model_dump()
        payload["forecast"] = {k: v for k, v in payload["forecast"].items()
                               if v is not None}
        outcome = pipeline.run_safe(payload, out_dir, progress)
        with RUNS_LOCK:
            state["status"] = "done" if outcome["ok"] else "error"
            state["result"] = outcome.get("results")
            state["error"] = outcome.get("error")
            if outcome["ok"]:
                state["result"]["files"] = [
                    os.path.basename(f) for f in state["result"]["files"]]

    threading.Thread(target=work, daemon=True).start()
    return {"run_id": run_id}


@app.get("/api/runs/{run_id}")
def get_run(run_id: str):
    with RUNS_LOCK:
        state = RUNS.get(run_id)
        if not state:
            raise HTTPException(404, "Run not found")
        return state


@app.get("/api/runs/{run_id}/files/{filename}")
def download_file(run_id: str, filename: str):
    if "/" in filename or "\\" in filename or ".." in filename:
        raise HTTPException(400, "Invalid filename")
    path = os.path.join(OUTPUT_DIR, run_id, filename)
    if not os.path.isfile(path):
        raise HTTPException(404, "File not found")
    return FileResponse(path, filename=filename)


@app.delete("/api/runs/{run_id}")
def delete_run(run_id: str):
    """Data-deletion endpoint (CCPA right to delete)."""
    import shutil
    with RUNS_LOCK:
        RUNS.pop(run_id, None)
    run_dir = os.path.join(OUTPUT_DIR, run_id)
    if os.path.isdir(run_dir):
        shutil.rmtree(run_dir)
    return {"deleted": run_id}


@app.get("/api/ein/lookup")
def ein_lookup(name: str, state: str = ""):
    return ein.lookup_ein(name, state or None)


@app.get("/api/trademark/search")
def trademark_search(q: str, goods: str = ""):
    return trademark.search_trademark(q, goods)


@app.get("/api/health")
def health():
    return {"ok": True}


app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
