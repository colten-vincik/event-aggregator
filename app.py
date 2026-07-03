"""
app.py — Flask web interface for the US Event Aggregator

Routes
------
GET  /                   → city picker form
POST /run                → start a scraper job, return job_id (JSON)
GET  /stream/<job_id>    → Server-Sent Events stream of progress lines
GET  /status/<job_id>    → job status + result paths (JSON)
GET  /download/<job_id>/<kind>  → download xlsx | map | report
GET  /health             → uptime check for Railway / Render
"""

import os, uuid, time, queue, threading
from pathlib import Path
from flask import (Flask, render_template, request, jsonify,
                   Response, send_file, abort)

import scraper

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-change-me-in-prod")

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

# ── In-memory job store (fine for single-worker deployment) ───────────────────
# job_id -> {"status": "running"|"done"|"error",
#             "queue":  Queue of log lines,
#             "result": dict | None,
#             "error":  str | None}
_JOBS: dict[str, dict] = {}
_JOBS_LOCK = threading.Lock()


def _build_city_label(mode, value):
    """Turn form input into a (cities list, label string) pair."""
    if mode == "region":
        region = value.lower()
        city_keys = scraper.REGIONS.get(region, [])
        cities = []
        for ck in city_keys:
            cfg   = scraper.CITY_CONFIG.get(ck, {})
            state = cfg.get("state", "")
            cities.append(f"{ck.title()}, {state}" if state else ck.title())
        label = f"{region.title()} Region"
        return cities, label

    if mode == "multi":
        cities = [c.strip() for c in value.split(";") if c.strip()]
        label  = "; ".join(cities)
        return cities, label

    # single city
    city_key = value.lower().strip()
    cfg   = scraper.CITY_CONFIG.get(city_key, {})
    state = cfg.get("state", "")
    city_label = f"{city_key.title()}, {state}" if state else city_key.title()
    return [city_label], city_label


# ─────────────────────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    cities  = [c.title() for c in scraper.CITY_OPTIONS]
    regions = [r.title() for r in scraper.REGION_OPTIONS]
    return render_template("index.html", cities=cities, regions=regions)


@app.route("/run", methods=["POST"])
def start_run():
    mode             = request.form.get("mode", "single")
    value            = request.form.get("value", "").strip()
    skip_attractions = request.form.get("skip_attractions") == "1"

    if not value:
        return jsonify({"error": "No city / region selected."}), 400

    try:
        cities, label = _build_city_label(mode, value)
    except Exception as e:
        return jsonify({"error": str(e)}), 400

    job_id = str(uuid.uuid4())
    q: queue.Queue[str | None] = queue.Queue()

    with _JOBS_LOCK:
        _JOBS[job_id] = {"status": "running", "queue": q,
                         "result": None, "error": None}

    def worker():
        def cb(msg):
            q.put(msg)

        try:
            result = scraper.run(
                cities=cities,
                label=label,
                skip_attractions=skip_attractions,
                output_dir=str(OUTPUT_DIR),
                progress_cb=cb,
            )
            with _JOBS_LOCK:
                _JOBS[job_id]["status"] = "done"
                _JOBS[job_id]["result"] = result
        except Exception as e:
            with _JOBS_LOCK:
                _JOBS[job_id]["status"] = "error"
                _JOBS[job_id]["error"]  = str(e)
        finally:
            q.put(None)  # sentinel — tells the SSE stream to close

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/stream/<job_id>")
def stream(job_id):
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job:
        abort(404)

    def generate():
        q = job["queue"]
        while True:
            try:
                msg = q.get(timeout=30)
            except queue.Empty:
                yield "data: \n\n"   # keepalive ping
                continue
            if msg is None:
                yield "data: __DONE__\n\n"
                break
            # Escape newlines inside the SSE data field
            safe = msg.replace("\n", " ").replace("\r", "")
            yield f"data: {safe}\n\n"

    return Response(generate(),
                    mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache",
                             "X-Accel-Buffering": "no"})


@app.route("/status/<job_id>")
def status(job_id):
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "status":      job["status"],
        "result":      job["result"],
        "error":       job["error"],
    })


@app.route("/download/<job_id>/<kind>")
def download(job_id, kind):
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job or job["status"] != "done":
        abort(404)
    result = job["result"]
    paths  = {"xlsx": result["xlsx"], "map": result["map"],
               "report": result["report"]}
    if kind not in paths:
        abort(400)
    path = Path(paths[kind])
    if not path.exists():
        abort(404)
    return send_file(str(path), as_attachment=True, download_name=path.name)


@app.route("/health")
def health():
    return jsonify({"status": "ok", "timestamp": time.time()})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
