"""
scraper.py — library wrapper around event_aggregator_v8.py

Exposes a single public function:

    run(cities, label, skip_attractions=False, output_dir="output",
        progress_cb=None)  ->  {"xlsx": path, "map": path, "report": path,
                                "events": int, "attractions": int, "elapsed": float}

`progress_cb(msg: str)` is called with each log line so callers (Flask SSE, CLI)
can stream progress in real time.
"""

import importlib, sys, os, re, time, threading
from datetime import datetime
from pathlib import Path

# ── Import the aggregator module ──────────────────────────────────────────────
# event_aggregator_v8.py lives next to this file.  We import it as a module
# rather than exec'ing it so all of its globals are available.
_here = Path(__file__).parent
sys.path.insert(0, str(_here))

import event_aggregator_v8 as _agg


def run(cities: list[str],
        label: str,
        skip_attractions: bool = False,
        output_dir: str = "output",
        progress_cb=None,
        date_from: str = "",
        date_to: str = "",
        weekday_after: int | None = None,
        categories: list | None = None,
        free_only: bool = False,
        max_price: float | None = None,
        boroughs: list | None = None) -> dict:
    """
    Run the aggregator for the given list of city strings.

    Parameters
    ----------
    cities          : e.g. ["Philadelphia, PA", "Boston, MA"]
    label           : human-readable label used in filenames and map title
    skip_attractions: skip Wikidata / OSM (faster for multi-city runs)
    output_dir      : directory to write xlsx / html / txt files into
    progress_cb     : optional callable(str) called for each log line

    Returns
    -------
    dict with keys: xlsx, map, report, events, attractions, elapsed
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Redirect the aggregator's tprint so progress goes to the callback too
    original_tprint = _agg.tprint
    _print_lock = threading.Lock()

    def patched_tprint(*args, **kwargs):
        msg = " ".join(str(a) for a in args)
        with _print_lock:
            print(msg)
            _agg.RUNLOG.line(msg)
            if progress_cb:
                progress_cb(msg)

    _agg.tprint = patched_tprint

    # Reset the global run-log so successive web requests don't bleed together
    _agg.RUNLOG = _agg.RunLog()

    try:
        t0 = time.time()
        all_events, all_attractions = [], []
        city_lock = threading.Lock()
        city_workers = min(4, len(cities))
        src_workers  = 6

        def run_city(city_loc):
            evts, atts = _agg.fetch_city(city_loc,
                                          skip_attractions=skip_attractions,
                                          source_workers=src_workers,
                                          date_from=date_from,
                                          date_to=date_to)
            with city_lock:
                all_events.extend(evts)
                all_attractions.extend(atts)
            return city_loc, len(evts), len(atts)

        from concurrent.futures import ThreadPoolExecutor, as_completed
        with ThreadPoolExecutor(max_workers=city_workers) as pool:
            futures = {pool.submit(run_city, city): city for city in cities}
            for future in as_completed(futures):
                try:
                    city_loc, ne, na = future.result()
                    patched_tprint(f"\n  ✅  {city_loc}: {ne} events, {na} attractions")
                except Exception as e:
                    patched_tprint(f"\n  ⚠  {futures[future]} failed: {e}")

        raw_count   = len(all_events)
        all_events  = _agg.dedup(all_events, filter_articles=True, validate_events=True,
                                 date_from=date_from, date_to=date_to,
                                 weekday_after=weekday_after,
                                 categories=categories, free_only=free_only,
                                 max_price=max_price, boroughs=boroughs)
        all_atts    = _agg.dedup(all_attractions)
        dropped     = raw_count - len(all_events)
        patched_tprint(f"\n  🔍  {raw_count} raw → {len(all_events)} events kept "
                       f"({dropped} dropped)")
        all_events.sort(key=lambda e: (e.get("Date","") or "9999", e.get("City","")))

        elapsed = time.time() - t0

        if not all_events and not all_atts:
            raise RuntimeError("No data found — all sources returned 0 results.")

        # ── Write outputs ──────────────────────────────────────────────────────
        safe = re.sub(r"[^\w]", "_", label)[:40]
        ts   = datetime.today().strftime("%Y%m%d_%H%M%S")

        xlsx_path   = out_dir / f"events_{safe}_{ts}.xlsx"
        map_path    = out_dir / f"events_{safe}_{ts}_map.html"
        report_path = out_dir / f"run_report_{ts}.txt"

        picks = _agg.quick_pick(all_events, n=5)

        from openpyxl import Workbook
        wb = Workbook()
        wb.remove(wb.active)
        _agg.build_summary(wb, all_events, all_atts, label)
        if picks:      _agg.build_quick_picks_sheet(wb, picks, label)
        if all_events: _agg.build_events_sheet(wb, all_events, label)
        if all_atts:   _agg.build_attractions_sheet(wb, all_atts, label)
        wb.save(str(xlsx_path))
        patched_tprint(f"  ✅  Excel saved: {xlsx_path.name}")

        _agg.build_map(all_events, all_atts, label, str(map_path))
        patched_tprint(f"  ✅  Map saved:   {map_path.name}")

        _agg.RUNLOG.write(str(report_path), label, elapsed)
        patched_tprint(f"  ✅  Report saved: {report_path.name}")

        return {
            "xlsx":        str(xlsx_path),
            "map":         str(map_path),
            "report":      str(report_path),
            "events":      len(all_events),
            "attractions": len(all_atts),
            "elapsed":     round(elapsed, 1),
            "picks":       picks,
            "by_category": dict(
                sorted(
                    __import__("collections").Counter(
                        e.get("Category","Other") for e in all_events
                    ).most_common()
                )
            ),
        }

    finally:
        _agg.tprint = original_tprint


# ── Expose config so Flask can populate the dropdowns ─────────────────────────
CITY_OPTIONS = sorted(_agg.CITY_CONFIG.keys())   # ["albany", "albuquerque", …]
REGION_OPTIONS = list(_agg.REGIONS.keys())        # ["northeast", "southeast", …]
REGIONS = _agg.REGIONS
CITY_CONFIG = _agg.CITY_CONFIG
