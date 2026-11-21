#!/usr/bin/env python3
"""Scrape les agences ORPI (nom, adresse, services) -> xlsx.

Parcours : page index -> pages régions / départements / villes
(liens /agences-immobilieres-<slug>/). Chaque page contient des « cartes »
agences (data-oncrawl="agency-card") ; on dédoublonne sur l'URL de l'agence.
Sortie : orpi_agences.xlsx (Région, Département, Nom, Adresse, Services,
Téléphone, URL). Région/département déduits du code postal.
"""
import logging
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urldefrag, urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE = "https://www.orpi.com"
INDEX_URL = f"{BASE}/agences-immobilieres/"
OUTPUT = "orpi_agences.xlsx"
WORKERS = 4
DELAY = 0.3
TIMEOUT = 30
MAX_PAGES = int(os.environ.get("MAX_PAGES", "5000"))   # garde-fou

LISTING_RE = re.compile(r"^/agences-immobilieres-[a-z0-9-]+/?$")

# Région -> {code département: nom}. Sert à déduire Région/Département
# du code postal de l'adresse (indépendant de la structure des pages).
GEO = {
    "Auvergne-Rhône-Alpes": {"01": "Ain", "03": "Allier", "07": "Ardèche", "15": "Cantal",
        "26": "Drôme", "38": "Isère", "42": "Loire", "43": "Haute-Loire",
        "63": "Puy-de-Dôme", "69": "Rhône", "73": "Savoie", "74": "Haute-Savoie"},
    "Bourgogne-Franche-Comté": {"21": "Côte-d'Or", "25": "Doubs", "39": "Jura", "58": "Nièvre",
        "70": "Haute-Saône", "71": "Saône-et-Loire", "89": "Yonne",
        "90": "Territoire de Belfort"},
    "Bretagne": {"22": "Côtes-d'Armor", "29": "Finistère", "35": "Ille-et-Vilaine",
        "56": "Morbihan"},
    "Centre-Val de Loire": {"18": "Cher", "28": "Eure-et-Loir", "36": "Indre",
        "37": "Indre-et-Loire", "41": "Loir-et-Cher", "45": "Loiret"},
    "Corse": {"2A": "Corse-du-Sud", "2B": "Haute-Corse"},
    "Grand Est": {"08": "Ardennes", "10": "Aube", "51": "Marne", "52": "Haute-Marne",
        "54": "Meurthe-et-Moselle", "55": "Meuse", "57": "Moselle", "67": "Bas-Rhin",
        "68": "Haut-Rhin", "88": "Vosges"},
    "Hauts-de-France": {"02": "Aisne", "59": "Nord", "60": "Oise", "62": "Pas-de-Calais",
        "80": "Somme"},
    "Île-de-France": {"75": "Paris", "77": "Seine-et-Marne", "78": "Yvelines", "91": "Essonne",
        "92": "Hauts-de-Seine", "93": "Seine-Saint-Denis", "94": "Val-de-Marne",
        "95": "Val-d'Oise"},
    "Normandie": {"14": "Calvados", "27": "Eure", "50": "Manche", "61": "Orne",
        "76": "Seine-Maritime"},
    "Nouvelle-Aquitaine": {"16": "Charente", "17": "Charente-Maritime", "19": "Corrèze",
        "23": "Creuse", "24": "Dordogne", "33": "Gironde", "40": "Landes",
        "47": "Lot-et-Garonne", "64": "Pyrénées-Atlantiques", "79": "Deux-Sèvres",
        "86": "Vienne", "87": "Haute-Vienne"},
    "Occitanie": {"09": "Ariège", "11": "Aude", "12": "Aveyron", "30": "Gard",
        "31": "Haute-Garonne", "32": "Gers", "34": "Hérault", "46": "Lot", "48": "Lozère",
        "65": "Hautes-Pyrénées", "66": "Pyrénées-Orientales", "81": "Tarn",
        "82": "Tarn-et-Garonne"},
    "Pays de la Loire": {"44": "Loire-Atlantique", "49": "Maine-et-Loire", "53": "Mayenne",
        "72": "Sarthe", "85": "Vendée"},
    "Provence-Alpes-Côte d'Azur": {"04": "Alpes-de-Haute-Provence", "05": "Hautes-Alpes",
        "06": "Alpes-Maritimes", "13": "Bouches-du-Rhône", "83": "Var", "84": "Vaucluse"},
    "France d'outre-mer": {"971": "Guadeloupe", "972": "Martinique", "973": "Guyane",
        "974": "La Réunion", "976": "Mayotte"},
}
DEPT = {code: (reg, name) for reg, d in GEO.items() for code, name in d.items()}


def geo_from_address(address: str) -> tuple[str, str]:
    """(région, département) déduits du dernier code postal à 5 chiffres."""
    cps = re.findall(r"\b(\d{5})\b", address)
    if not cps:
        return "", ""
    cp = cps[-1]
    if cp.startswith("20"):
        code = "2A" if int(cp[:3]) <= 201 else "2B"
    elif cp.startswith("97"):
        code = cp[:3]
    else:
        code = cp[:2]
    reg, name = DEPT.get(code, ("Autre", code))
    return reg, f"{code} {name}" if code in DEPT else name


logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("orpi")

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "fr-FR,fr;q=0.9",
})
session.mount("https://", HTTPAdapter(max_retries=Retry(
    total=4, backoff_factor=1.5, status_forcelist=(429, 500, 502, 503, 504))))


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


def norm(href: str, base: str) -> str:
    url = urldefrag(urljoin(base, href))[0].split("?")[0]
    return url if url.endswith("/") else url + "/"


def parse_cards(soup: BeautifulSoup) -> list[dict]:
    """Cartes agences d'une page -> [{nom, adresse, services, tel, url}]."""
    out = []
    for card in soup.select('[data-oncrawl="agency-card"]'):
        name_el = card.select_one('[data-oncrawl="agency-name"]')
        if not name_el:
            continue
        addr_el = card.select_one('[data-oncrawl="agency-address"]') or card.find("address")
        address = ", ".join(addr_el.stripped_strings) if addr_el else ""
        services = [li.get_text(" ", strip=True)
                    for li in card.select("ul.list-unstyled li")]
        phone_el = card.select_one('[data-oncrawl="agency-phone"]')
        link_el = card.select_one('[data-oncrawl="agency-show-link"]')
        out.append({
            "nom": name_el.get_text(" ", strip=True),
            "adresse": address,
            "services": ", ".join(s for s in services if s),
            "tel": phone_el.get_text(strip=True) if phone_el else "",
            "url": norm(link_el["href"], BASE) if link_el and link_el.get("href") else "",
        })
    return out


def parse_links(soup: BeautifulSoup, page_url: str) -> list[str]:
    """Pages de liste (région/département/ville) + pagination éventuelle."""
    links = []
    for a in soup.find_all("a", href=True):
        path = re.sub(r"^https://www\.orpi\.com", "", urldefrag(a["href"])[0].split("?")[0])
        if LISTING_RE.match(path):
            links.append(norm(a["href"], page_url))
    nxt = soup.select_one('a[rel="next"]')
    if nxt and nxt.get("href"):
        links.append(urljoin(page_url, nxt["href"]))
    return list(dict.fromkeys(links))


def main() -> int:
    seen = {INDEX_URL}
    frontier = [INDEX_URL]
    agencies: dict[str, dict] = {}

    while frontier and len(seen) <= MAX_PAGES:
        with ThreadPoolExecutor(WORKERS) as ex:
            soups = list(ex.map(fetch, frontier))
        nxt_frontier = []
        for url, soup in zip(frontier, soups):
            if soup is None:
                continue
            for ag in parse_cards(soup):
                key = ag["url"] or f'{ag["nom"]}|{ag["adresse"]}'
                agencies.setdefault(key, ag)
            for link in parse_links(soup, url):
                if link not in seen:
                    seen.add(link)
                    nxt_frontier.append(link)
        frontier = nxt_frontier
        log.info("%d pages vues, %d agences uniques", len(seen), len(agencies))

    if not agencies:
        log.error("Aucune agence récupérée (site bloqué ou structure modifiée ?).")
        return 1

    for a in agencies.values():
        a["region"], a["dept"] = geo_from_address(a["adresse"])
    rows = sorted(agencies.values(), key=lambda a: (a["region"], a["dept"], a["nom"]))
    counts: dict[str, int] = {}
    for a in rows:
        counts[a["dept"]] = counts.get(a["dept"], 0) + 1
    log.info("Agences par département : %s", dict(sorted(counts.items())))
    wb = Workbook()
    ws = wb.active
    ws.title = "Agences"
    ws.append(["Région", "Département", "Nom de l'agence", "Adresse", "Services",
               "Téléphone", "URL"])
    for a in rows:
        ws.append([a["region"], a["dept"], a["nom"], a["adresse"], a["services"],
                   a["tel"], a["url"]])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, width in zip("ABCDEFG", (26, 24, 40, 45, 55, 18, 45)):
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
