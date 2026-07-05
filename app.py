"""
app.py — Flask web interface for the US Event Aggregator

Routes
------
GET  /                        → city picker form
POST /run                     → start a scraper job, return job_id (JSON)
GET  /stream/<job_id>         → Server-Sent Events stream of progress lines
GET  /status/<job_id>         → job status + result paths (JSON)
GET  /events/<job_id>         → full event list for the browser table (JSON)
GET  /download/<job_id>/<kind>→ download xlsx | map | report
POST /plan                    → generate evening itinerary via Claude API (JSON)
GET  /health                  → uptime check for Railway / Render
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
    date_from        = request.form.get("date_from", "").strip()
    date_to          = request.form.get("date_to", "").strip()

    # weekday_after: "HH:MM" string → int minutes, or None if not set
    weekday_after = None
    wa_raw = request.form.get("weekday_after", "").strip()
    if wa_raw:
        try:
            h, m = wa_raw.split(":")
            weekday_after = int(h) * 60 + int(m)
        except Exception:
            pass

    # categories: comma-separated list, or None if not set
    cats_raw = request.form.get("categories", "").strip()
    categories = [c.strip() for c in cats_raw.split(",") if c.strip()] or None

    # price filters
    free_only = request.form.get("free_only") == "1"
    max_price = None
    mp_raw = request.form.get("max_price", "").strip()
    if mp_raw and not free_only:
        try:
            max_price = float(mp_raw)
        except Exception:
            pass

    # boroughs: comma-separated list, or None if not set
    boros_raw = request.form.get("boroughs", "").strip()
    boroughs = [b.strip().lower() for b in boros_raw.split(",") if b.strip()] or None

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

        def _clean_pick(e):
            return {k: v for k, v in e.items() if not k.startswith("_")}

        try:
            result = scraper.run(
                cities=cities,
                label=label,
                skip_attractions=skip_attractions,
                output_dir=str(OUTPUT_DIR),
                progress_cb=cb,
                date_from=date_from,
                date_to=date_to,
                weekday_after=weekday_after,
                categories=categories,
                free_only=free_only,
                max_price=max_price,
                boroughs=boroughs,
            )
            result["picks"] = [_clean_pick(p) for p in result.get("picks", [])]
            # Store events separately — keeps /status payload small
            events_data = result.pop("events_data", [])
            with _JOBS_LOCK:
                _JOBS[job_id]["status"]      = "done"
                _JOBS[job_id]["result"]      = result
                _JOBS[job_id]["events_data"] = events_data
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


@app.route("/events/<job_id>")
def get_events(job_id):
    """Return the full sorted event list for the in-browser table."""
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job or job["status"] != "done":
        abort(404)
    return jsonify(job.get("events_data", []))


@app.route("/plan", methods=["POST"])
def plan_evening():
    """Generate an evening itinerary via Claude for the selected events."""
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return jsonify({"error": "ANTHROPIC_API_KEY is not configured."}), 503

    data   = request.get_json(silent=True) or {}
    events = data.get("events", [])
    if not events:
        return jsonify({"error": "No events provided."}), 400

    # Build a concise event summary for the prompt
    lines = []
    for e in events:
        parts = [f"• {e.get('Name', '?')}"]
        if e.get("Category"): parts.append(f"({e['Category']})")
        if e.get("Date"):     parts.append(f"— {e['Date']}")
        if e.get("Time"):     parts.append(f"at {e['Time']}")
        if e.get("Venue"):    parts.append(f"@ {e['Venue']}")
        if e.get("City"):     parts.append(f"in {e['City']}")
        if e.get("Price"):    parts.append(f"[{e['Price']}]")
        if e.get("URL"):      parts.append(f"\n  {e['URL']}")
        lines.append(" ".join(parts))

    events_block = "\n".join(lines)

    prompt = f"""You are helping a New York City resident plan their evening after work. \
They finish work around 5:00–5:30 PM and want to make the most of their night. \
They are comfortable with the subway and walking.

They have selected the following event(s) as the anchor(s) for their evening:

{events_block}

Write a practical, loose evening itinerary as a timeline starting from when they leave work. Include:

- A realistic lead-up: where to grab a quick dinner or drink near the first venue's neighborhood, and how early to arrive
- Each selected event in chronological order with a sentence of context (what to expect, tips)
- Subway or walking guidance when moving between venues
- A post-event suggestion (a nearby bar, a walk, a late-night spot) if the evening ends before midnight
- One backup idea in case something falls through

Write like a knowledgeable NYC friend giving real advice — specific, opinionated, and brief. \
Use a clear timeline format (e.g. "5:30 PM —"). Keep the total response under 450 words."""

    try:
        import anthropic
        client  = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=700,
            messages=[{"role": "user", "content": prompt}],
        )
        return jsonify({"itinerary": message.content[0].text})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/health")
def health():
    return jsonify({"status": "ok", "timestamp": time.time()})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
