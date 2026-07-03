# US Event Aggregator

A Python tool that searches the internet for **events, concerts, attractions, and things to do** across 47 US cities and outputs a searchable Excel workbook + interactive map — built for travel planning.

## Features

- **47 cities** across all US regions (Northeast, Southeast, Midwest, Southwest, West Coast)
- **Multi-city & region presets** — search "northeast" to pull all NE cities at once
- **Parallel fetching** — sources run concurrently per city, ~5-10x faster than sequential
- **8+ event sources**: TimeOut · Patch.com · Eventbrite · Songkick · AllEvents.in · DoStuff network · Yelp Events
- **Attraction sources**: Wikidata SPARQL · OpenStreetMap Overpass (real GPS coordinates)
- **Smart filtering** — past events, list articles, songs, and undated entries are removed automatically
- **Excel output** — color-coded by category, auto-filter on every column, short hyperlinks, auto row heights
- **Interactive HTML map** — Leaflet.js + MarkerCluster, category toggles, popups with dates/venues/links
- **Run report** — auto-generated `.txt` on every run listing HTTP errors and per-source counts for debugging

## Output files

Each run produces two files:
```
events_{city}_{date}.xlsx          ← searchable workbook (Events + Landmarks + Summary tabs)
events_{city}_{date}_map.html      ← interactive map (open in any browser)
```

## Quick start

```bash
pip install requests beautifulsoup4 lxml openpyxl
python event_aggregator_v8.py
```

Then follow the prompts to pick a city, region, or custom list.

## Requirements

```
requests
beautifulsoup4
lxml
openpyxl
```

## Sources

| Source | Type | Notes |
|---|---|---|
| TimeOut | Scraper | 35+ city slugs, follows list articles |
| Patch.com | Scraper | Community events via `__NEXT_DATA__` JSON |
| Eventbrite | Scraper | `__NEXT_DATA__` + JSON-LD |
| Songkick | Scraper | Metro area pages (verified slugs for 17 cities) |
| AllEvents.in | Scraper | JSON-LD + card fallback |
| DoStuff network | Scraper | 24 city sites (do312, do617, do215…) |
| Yelp Events | Scraper | `/search?cflt=yelpevents` |
| Wikidata SPARQL | API | Attractions with Wikipedia articles, real GPS coords |
| OpenStreetMap Overpass | API | Attractions with addresses/websites, real GPS coords |

## Roadmap

- [ ] Ticketmaster Discovery API integration
- [ ] SeatGeek API integration
- [ ] Flask web interface (city picker → live results → download)
- [ ] GitHub Pages landing page
- [ ] Scheduled daily refresh via GitHub Actions

## Version history

| Version | Changes |
|---|---|
| v6 | NYC-only, sequential fetching |
| v7 | 47-city config, multi-city mode, Eventbrite/Songkick/AllEvents added |
| v8 | Parallel fetching, DoStuff network, Yelp Events, run reports, interactive map, event validation |

## License

MIT
