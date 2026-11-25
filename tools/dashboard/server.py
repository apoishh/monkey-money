"""Local dashboard for Monkey Money - http://localhost:8765 (only reachable from this PC)."""
import os
import socket
import sys
import threading
import traceback
import webbrowser

import uvicorn
from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import data  # noqa: E402

HOST, PORT = "127.0.0.1", 8765
URL = f"http://localhost:{PORT}"

app = FastAPI(title="Monkey Money", docs_url=None, redoc_url=None)


@app.middleware("http")
async def no_cache(request, call_next):
    # always serve the latest page code (and data) - nothing is cached by the browser
    resp = await call_next(request)
    resp.headers["Cache-Control"] = "no-store"
    return resp


app.mount("/static", StaticFiles(directory=os.path.join(HERE, "static")), name="static")


@app.get("/")
def index():
    return FileResponse(os.path.join(HERE, "static", "index.html"), headers={"Cache-Control": "no-store"})


@app.get("/api/data")
def api_data():
    try:
        return JSONResponse(data.load(), headers={"Cache-Control": "no-store"})
    except Exception as e:  # show the reason on the page instead of a blank screen
        traceback.print_exc()
        return JSONResponse({"error": f"{type(e).__name__}: {e}"}, status_code=500)


def already_running():
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex((HOST, PORT)) == 0


if __name__ == "__main__":
    if already_running():
        print("Dashboard is already running - opening it.")
        webbrowser.open(URL)
        sys.exit(0)
    if "--no-browser" not in sys.argv:
        threading.Timer(1.5, lambda: webbrowser.open(URL)).start()
    print(f"Dashboard running at {URL}  (close this window to stop it)")
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")
