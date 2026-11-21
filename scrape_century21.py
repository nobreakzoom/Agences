#!/usr/bin/env python3
"""Scrape les agences CENTURY 21 par région (nom + spécialités) -> xlsx.

Parcours : page index (régions) -> départements -> villes -> fiches agences.
Sortie  : century21_agences.xlsx (colonnes : Région, Nom de l'agence,
          Service de l'agence, URL).
"""
import logging
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urldefrag, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE = "https://www.century21.fr"
INDEX_URL = f"{BASE}/trouver_agence/"
OUTPUT = "century21_agences.xlsx"
WORKERS = 4          # requêtes en parallèle (reste raisonnable pour le site)
DELAY = 0.3          # pause (s) après chaque requête, par thread
TIMEOUT = 30

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("c21")

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "fr-FR,fr;q=0.9",
})
session.mount("https://", HTTPAdapter(max_retries=Retry(
    total=4, backoff_factor=1.5, status_forcelist=(429, 500, 502, 503, 504))))


# ----------------------------------------------------------------- utilitaires
def clean_url(href: str, base: str = INDEX_URL) -> str:
    """URL absolue, sans query ni fragment (supprime ?_gl=… etc.), avec « / » final."""
    url = urldefrag(urljoin(base, href))[0].split("?")[0]
    return url if url.endswith("/") else url + "/"


def fetch(url: str) -> BeautifulSoup | None:
    try:
        r = session.get(url, timeout=TIMEOUT)
        r.raise_for_status()
        time.sleep(DELAY)
        r.encoding = "utf-8"
        return BeautifulSoup(r.text, "html.parser")
    except requests.RequestException as exc:
        log.warning("Échec %s : %s", url, exc)
        return None


# ------------------------------------------------------------------- parsing
def parse_regions(soup: BeautifulSoup) -> dict[str, list[str]]:
    """{région: [liens départements / villes]} depuis la page index."""
    tabs = [t.get_text(strip=True) for t in soup.select(".sn-agence-index__table-tr")]
    regions: dict[str, list[str]] = {}
    for block in soup.select('[x-ref^="item"]'):
        m = re.fullmatch(r"item(\d+)", block.get("x-ref", ""))
        if not m or int(m.group(1)) >= len(tabs):
            continue
        links = [clean_url(a["href"]) for a in block.find_all("a", href=True)
                 if re.search(r"/trouver_agence/(d|v|cpv)-", a["href"])]
        regions[tabs[int(m.group(1))]] = list(dict.fromkeys(links))
    return regions


def parse_listing(soup: BeautifulSoup) -> tuple[list[str], list[str]]:
    """Page département/ville -> (URLs d'agences, URLs de villes).

    Les agences listées après « Nos autres agences dans les départements
    limitrophes » sont ignorées (elles appartiennent à d'autres zones).
    """
    agencies = []
    for li in soup.select("li.js-the-list-of-agencies-agency"):
        prev_h2 = li.find_previous("h2")
        if prev_h2 and "autres agences" in prev_h2.get_text():
            continue
        span = li.select_one("[data-url]")
        href = span["data-url"] if span else None
        if not href:
            a = li.find("a", href=re.compile(r"/trouver_agence/agence/"))
            href = a["href"] if a else None
        if href:
            agencies.append(clean_url(href))

    cities = []
    for a in soup.find_all("a", href=True):
        if re.search(r"/trouver_agence/(v|cpv)-", a["href"]):
            cities.append(clean_url(a["href"]))
    return list(dict.fromkeys(agencies)), list(dict.fromkeys(cities))


def parse_agency(soup: BeautifulSoup) -> tuple[str, str]:
    """Fiche agence -> (nom, services séparés par des virgules)."""
    h1 = soup.select_one("h1")
    name = h1.get_text(" ", strip=True) if h1 else ""
    name = re.sub(r"^Agence immobili\S{1,3}re\s+", "", name, flags=re.I)

    services: list[str] = []
    for h2 in soup.find_all("h2"):
        if re.search(r"sp\S{1,3}cialit", h2.get_text(), re.I):
            ul = h2.find_next_sibling("ul") or h2.find_next("ul")
            if ul:
                services = [li.get_text(" ", strip=True) for li in ul.find_all("li")]
            break
    return name, ", ".join(s for s in services if s)


# --------------------------------------------------------------------- crawl
def next_pages(soup: BeautifulSoup, url: str) -> list[str]:
    nxt = soup.select_one('a[rel="next"]')
    return [urljoin(url, nxt["href"])] if nxt and nxt.get("href") else []


def crawl_listing(url: str, want_cities: bool) -> tuple[list[str], list[str]]:
    """Récupère agences (+ villes) d'une page, en suivant une éventuelle pagination."""
    agencies, cities, seen, queue = [], [], set(), [url]
    while queue:
        u = queue.pop(0)
        if u in seen:
            continue
        seen.add(u)
        soup = fetch(u)
        if soup is None:
            continue
        a, c = parse_listing(soup)
        agencies += a
        if want_cities:
            cities += c
        queue += next_pages(soup, u)
    return agencies, cities


def main() -> int:
    index = fetch(INDEX_URL)
    if index is None:
        log.error("Impossible de charger la page index.")
        return 1
    regions = parse_regions(index)
    log.info("%d régions trouvées", len(regions))

    agency_region: dict[str, str] = {}   # url agence -> région (1re rencontrée)
    city_region: dict[str, str] = {}     # url ville  -> région

    # 1) départements -> villes (+ agences éventuelles)
    dept_jobs = []
    for region, links in regions.items():
        for link in links:
            if "/trouver_agence/d-" in link:
                dept_jobs.append((region, link))
            else:
                city_region.setdefault(link, region)
    with ThreadPoolExecutor(WORKERS) as ex:
        results = list(ex.map(lambda j: crawl_listing(j[1], True), dept_jobs))
    for (region, _), (agencies, cities) in zip(dept_jobs, results):
        for c in cities:
            city_region.setdefault(c, region)
        for a in agencies:
            agency_region.setdefault(a, region)
    log.info("%d villes à parcourir", len(city_region))

    # 2) villes -> agences
    city_jobs = list(city_region.items())
    with ThreadPoolExecutor(WORKERS) as ex:
        results = list(ex.map(lambda j: crawl_listing(j[0], False), city_jobs))
    for (_, region), (agencies, _) in zip(city_jobs, results):
        for a in agencies:
            agency_region.setdefault(a, region)
    log.info("%d agences uniques", len(agency_region))

    # 3) fiches agences -> nom + spécialités
    def get_agency(url: str):
        soup = fetch(url)
        return (url, *parse_agency(soup)) if soup else (url, "", "")

    with ThreadPoolExecutor(WORKERS) as ex:
        fiches = list(ex.map(get_agency, agency_region))

    rows = sorted(
        ((agency_region[u], n, s, u) for u, n, s in fiches if n),
        key=lambda r: (r[0], r[1]),
    )
    if not rows:
        log.error("Aucune agence récupérée (site bloqué ou structure modifiée ?).")
        return 1

    # 4) xlsx
    wb = Workbook()
    ws = wb.active
    ws.title = "Agences"
    ws.append(["Région", "Nom de l'agence", "Service de l'agence", "URL"])
    for row in rows:
        ws.append(row)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, width in zip("ABCD", (28, 45, 45, 60)):
        ws.column_dimensions[col].width = width
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(OUTPUT)
    log.info("%s écrit : %d agences", OUTPUT, len(rows))
    return 0


if __name__ == "__main__":
    sys.exit(main())
