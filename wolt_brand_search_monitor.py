#!/usr/bin/env python3
"""Search a brand across Wolt pharmacies and maintain a growing item_id catalog."""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_DIR = Path(__file__).resolve().parent
WOLT_PROJECT_DIR = BASE_DIR / "wolt_project"
DEFAULT_RESULTS_DIR = WOLT_PROJECT_DIR / "RESULTS"
DEFAULT_ITEMS_CATALOG = WOLT_PROJECT_DIR / "state" / "wolt_item_ids_catalog.csv"
DEFAULT_VITRUM_REFERENCE = WOLT_PROJECT_DIR / "state" / "wolt_vitrum_item_reference.csv"
DEFAULT_VITRUM_CANONICAL = WOLT_PROJECT_DIR / "state" / "wolt_vitrum_canonical_catalog.csv"
DEFAULT_VITRUM_UNMAPPED = WOLT_PROJECT_DIR / "state" / "wolt_vitrum_unmapped.csv"
DEFAULT_ACTIVE_INGREDIENT_REFERENCE = WOLT_PROJECT_DIR / "state" / "wolt_active_ingredient_reference.csv"
DEFAULT_FONT_PATH = BASE_DIR / "fonts" / "DejaVuSans.ttf"
_ENV_LOADED = False
SOLD_OUT_MARKERS = (
    "sold out",
    "out of stock",
    "unavailable",
    "нет в наличии",
    "законч",
    "временно недоступ",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Run brand search through Wolt pharmacies and update item_id catalog "
            "with new IDs (brand-aware)."
        )
    )
    parser.add_argument(
        "--pharmacies-csv",
        required=True,
        help="CSV with pharmacy list (must contain slug,name,venue_url)",
    )
    parser.add_argument("--brand", required=True, help="Brand label for catalog (e.g. Vitrum)")
    parser.add_argument(
        "--query",
        default="",
        help="Search query. If empty, uses --brand value.",
    )
    parser.add_argument("--language", default="ru", help="Search language parameter")
    parser.add_argument("--city-slug", default="almaty", help="City slug for product links")
    parser.add_argument("--timeout", type=float, default=25.0, help="HTTP timeout in seconds")
    parser.add_argument("--sleep-ms", type=int, default=120, help="Pause between requests, ms")
    parser.add_argument(
        "--results-dir",
        default=str(DEFAULT_RESULTS_DIR),
        help="Folder for output reports",
    )
    parser.add_argument(
        "--items-catalog",
        default=str(DEFAULT_ITEMS_CATALOG),
        help="Global item_id catalog CSV path",
    )
    parser.add_argument(
        "--vitrum-reference-csv",
        default=str(DEFAULT_VITRUM_REFERENCE),
        help="Detailed Vitrum item_id -> canonical SKU mapping CSV",
    )
    parser.add_argument(
        "--vitrum-canonical-csv",
        default=str(DEFAULT_VITRUM_CANONICAL),
        help="Canonical Vitrum SKU catalog (aggregated) CSV",
    )
    parser.add_argument(
        "--vitrum-unmapped-csv",
        default=str(DEFAULT_VITRUM_UNMAPPED),
        help="Vitrum rows that failed automatic canonical mapping",
    )
    parser.add_argument(
        "--send-telegram",
        action="store_true",
        help="Send summary + files to Telegram (uses TELEGRAM_BOT_TOKEN/TELEGRAM_CHAT_ID by default)",
    )
    parser.add_argument("--telegram-bot-token", default="", help="Telegram bot token override")
    parser.add_argument("--telegram-chat-id", default="", help="Telegram chat_id override")
    parser.add_argument(
        "--telegram-top-n",
        type=int,
        default=12,
        help="How many best offers to include in Telegram message (default: %(default)s)",
    )
    return parser.parse_args()


def load_env_from_file() -> None:
    global _ENV_LOADED
    if _ENV_LOADED:
        return
    env_path = BASE_DIR / ".env"
    if env_path.exists():
        for raw in env_path.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip("'\"")
            if key and key not in os.environ:
                os.environ[key] = value
    _ENV_LOADED = True


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip().lower())


def build_session() -> requests.Session:
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session = requests.Session()
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
            ),
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
    )
    return session


def read_pharmacies(path: Path, city_slug: str) -> list[dict[str, str]]:
    if not path.exists():
        raise FileNotFoundError(f"Pharmacies CSV not found: {path}")
    rows: list[dict[str, str]] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        if not reader.fieldnames:
            raise ValueError(f"CSV has no header: {path}")
        for row in reader:
            row = {k: (v or "").strip() for k, v in row.items()}
            slug = row.get("slug") or row.get("venue_slug")
            if not slug:
                continue
            rows.append(
                {
                    "slug": slug,
                    "name": row.get("name") or row.get("pharmacy") or slug,
                    "venue_url": row.get("venue_url")
                    or f"https://wolt.com/en/kaz/{city_slug}/venue/{slug}",
                }
            )
    if not rows:
        raise ValueError(f"No valid pharmacy rows found in {path}")
    return rows


def fetch_items(
    session: requests.Session,
    venue_slug: str,
    query: str,
    language: str,
    timeout: float,
) -> list[dict[str, Any]]:
    url = (
        "https://consumer-api.wolt.com/consumer-api/consumer-assortment/"
        f"v1/venues/slug/{venue_slug}/assortment/items/search"
    )
    response = session.post(
        url,
        params={"language": language},
        json={"q": query},
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    items = payload.get("items") if isinstance(payload, dict) else None
    return items if isinstance(items, list) else []


def parse_disable_info(item: dict[str, Any]) -> tuple[bool, str]:
    info = item.get("disabled_info")
    if not isinstance(info, dict):
        return False, ""
    disable_text = str(info.get("disable_text") or "")
    low = normalize_text(disable_text)
    return any(marker in low for marker in SOLD_OUT_MARKERS), disable_text


def item_status(item: dict[str, Any]) -> tuple[str, str]:
    sold_out, disable_text = parse_disable_info(item)
    if sold_out:
        return "out_of_stock", disable_text
    if disable_text:
        return "unavailable", disable_text
    return "in_stock", ""


def minor_to_major(value: Any) -> str:
    if isinstance(value, (int, float)):
        return f"{float(value)/100.0:.2f}"
    return ""


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fieldnames})


def _to_int(value: Any) -> int | None:
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        v = value.strip()
        if not v:
            return None
        try:
            return int(float(v))
        except Exception:
            return None
    return None


def _vitrum_make_result(
    *,
    raw_name: str,
    normalized_name: str,
    canonical_sku: str,
    canonical_name: str,
    product_line: str,
    dosage_or_volume: str,
    pack_size: str,
    form_factor: str,
    flavor: str,
    active_ingredient: str,
    confidence: str,
    rule: str,
    is_vitrum: bool,
) -> dict[str, str]:
    return {
        "raw_name": raw_name,
        "normalized_name": normalized_name,
        "canonical_sku": canonical_sku,
        "canonical_name": canonical_name,
        "product_line": product_line,
        "dosage_or_volume": dosage_or_volume,
        "pack_size": pack_size,
        "form_factor": form_factor,
        "flavor": flavor,
        "active_ingredient": active_ingredient,
        "confidence": confidence,
        "rule": rule,
        "is_vitrum": "1" if is_vitrum else "0",
    }


def normalize_vitrum_name(value: str) -> str:
    text = (value or "").upper().replace("Ё", "Е")
    text = text.replace("®", " ").replace("™", " ")
    text = text.replace("№", " N").replace("#", " N")
    text = text.replace("VITRUM", "ВИТРУМ")
    text = text.replace("IMMUNAKTIV", "ИММУНАКТИВ")
    text = text.replace("IMMUNOAKTIV", "ИММУНАКТИВ")
    text = text.replace("ИММУНОАКТИВ", "ИММУНАКТИВ")
    text = text.replace("ИМУНАКТИВ", "ИММУНАКТИВ")
    text = re.sub(r"\bNO(?=\s*\d)", "N", text)
    text = re.sub(r"\bN\s*(\d{1,3})\b", r"N\1", text)
    text = re.sub(r"[^0-9A-ZА-Я+]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _extract_pack_size(text: str) -> str:
    m = re.search(r"\bN(\d{1,3})\b", text)
    if m:
        return m.group(1)
    return ""


def _extract_mg_or_ml(text: str) -> str:
    mg = re.search(r"\b(\d{2,5})\s*МГ\b", text)
    if mg:
        return f"{mg.group(1)} мг"
    ml = re.search(r"\b(\d{1,3})\s*МЛ\b", text)
    if ml:
        return f"{ml.group(1)} мл"
    return ""


def _infer_vitrum_active_ingredient(norm: str, *, is_vitrum: bool) -> str:
    if not is_vitrum:
        if "ВИТРОКАП" in norm:
            return "комплексный БАД (состав зависит от SKU)"
        return ""
    if "ВИТАМИН С" in norm or "ВИТРУМ С" in norm:
        return "аскорбиновая кислота"
    if "Д3" in norm:
        return "холекальциферол"
    if "МАГНИ" in norm and "В6" in norm:
        return "магний + пиридоксин"
    return "витаминно-минеральный комплекс"


def canonicalize_vitrum_name(raw_name: str) -> dict[str, str]:
    name = (raw_name or "").strip()
    norm = normalize_vitrum_name(name)
    pack = _extract_pack_size(norm)
    mg_or_ml = _extract_mg_or_ml(norm)

    def make(
        sku: str,
        canonical_name: str,
        line: str,
        dosage_or_volume: str,
        form_factor: str,
        flavor: str,
        rule: str,
        active_ingredient: str = "",
        confidence: str = "high",
        is_vitrum: bool = True,
    ) -> dict[str, str]:
        return _vitrum_make_result(
            raw_name=name,
            normalized_name=norm,
            canonical_sku=sku,
            canonical_name=canonical_name,
            product_line=line,
            dosage_or_volume=dosage_or_volume,
            pack_size=pack,
            form_factor=form_factor,
            flavor=flavor,
            active_ingredient=active_ingredient or _infer_vitrum_active_ingredient(norm, is_vitrum=is_vitrum),
            confidence=confidence,
            rule=rule,
            is_vitrum=is_vitrum,
        )

    if "ВИТРОКАП" in norm:
        return make(
            sku="vitrocap_n30_caps",
            canonical_name="ВИТРОКАП N30 капс",
            line="Витрокап",
            dosage_or_volume=mg_or_ml,
            form_factor="капсулы",
            flavor="",
            rule="vitrocap",
            is_vitrum=False,
        )

    if "ВИТРУМ" not in norm:
        return make(
            sku="unknown_non_vitrum",
            canonical_name=name or "Unknown non-vitrum item",
            line="Unknown",
            dosage_or_volume=mg_or_ml,
            form_factor="",
            flavor="",
            rule="missing_vitrum_token",
            confidence="low",
            is_vitrum=False,
        )

    if "КИДС" in norm and "МАРМЕЛ" in norm:
        return make(
            sku="vitrum_kids_gummies_n60_apple",
            canonical_name="ВИТРУМ КИДС МАРМЕЛАДКИ N60 ЯБЛОКО",
            line="Витрум Кидс",
            dosage_or_volume=mg_or_ml,
            form_factor="мармеладки",
            flavor="яблоко",
            rule="kids_gummies",
        )
    if "КИДС" in norm:
        return make(
            sku="vitrum_kids_effervescent_n18",
            canonical_name="ВИТРУМ КИДС N18 ТАБ ШИП",
            line="Витрум Кидс",
            dosage_or_volume=mg_or_ml,
            form_factor="шипучие таблетки",
            flavor="",
            rule="kids_effervescent",
        )

    if "ЮНИОР" in norm and "МАРМЕЛ" in norm:
        return make(
            sku="vitrum_junior_gummies_n60_blackcurrant",
            canonical_name="ВИТРУМ ЮНИОР МАРМЕЛАДКИ N60 ЧЕРНАЯ СМОРОДИНА",
            line="Витрум Юниор",
            dosage_or_volume=mg_or_ml,
            form_factor="мармеладки",
            flavor="черная смородина",
            rule="junior_gummies",
        )
    if "ЮНИОР" in norm:
        return make(
            sku="vitrum_junior_effervescent_n18",
            canonical_name="ВИТРУМ ЮНИОР N18 ТАБ ШИП",
            line="Витрум Юниор",
            dosage_or_volume=mg_or_ml,
            form_factor="шипучие таблетки",
            flavor="",
            rule="junior_effervescent",
        )

    if "ПРЕНАТАЛ" in norm:
        return make(
            sku="vitrum_prenatal_standard_n30_caps",
            canonical_name="ВИТРУМ ПРЕНАТАЛ СТАНДАРТ N30 КАПС",
            line="Витрум Пренатал",
            dosage_or_volume=mg_or_ml,
            form_factor="капсулы",
            flavor="",
            rule="prenatal_standard",
        )

    if "МАГНИ" in norm and "В6" in norm:
        if pack == "90":
            return make(
                sku="vitrum_magnesium_b6_n90_tabs",
                canonical_name="ВИТРУМ МАГНИЙ В6 N90 ТАБ",
                line="Витрум Магний B6",
                dosage_or_volume=mg_or_ml,
                form_factor="таблетки",
                flavor="",
                rule="magnesium_b6_n90",
            )
        return make(
            sku="vitrum_magnesium_b6_n60_tabs",
            canonical_name="ВИТРУМ МАГНИЙ В6 N60 ТАБ П.О.",
            line="Витрум Магний B6",
            dosage_or_volume=mg_or_ml,
            form_factor="таблетки",
            flavor="",
            rule="magnesium_b6_n60",
            confidence="high" if pack in {"", "60"} else "medium",
        )

    if "Д3" in norm:
        if "МАКС" in norm:
            return make(
                sku="vitrum_vitamin_d3_max_n30_tabs",
                canonical_name="ВИТРУМ ВИТАМИН Д3 МАКС N30 ТАБ",
                line="Витрум Витамин D3",
                dosage_or_volume=mg_or_ml,
                form_factor="таблетки",
                flavor="",
                rule="vitamin_d3_max",
            )
        if "СПРЕЙ" in norm or "АКТИВ" in norm or "10МЛ" in norm:
            return make(
                sku="vitrum_vitamin_d3_aktiv_10ml_spray",
                canonical_name="ВИТРУМ ВИТАМИН Д3 АКТИВ 10МЛ СПРЕЙ",
                line="Витрум Витамин D3",
                dosage_or_volume=mg_or_ml or "10 мл",
                form_factor="спрей",
                flavor="",
                rule="vitamin_d3_aktiv",
            )

    if "ИММУНАКТИВ" in norm:
        if "ШИП" in norm:
            return make(
                sku="vitrum_immunaktiv_effervescent_n20",
                canonical_name="ВИТРУМ ИММУНАКТИВ ШИПУЧИЙ N20 ТАБ",
                line="Витрум Иммунактив",
                dosage_or_volume=mg_or_ml,
                form_factor="шипучие таблетки",
                flavor="",
                rule="immunaktiv_effervescent_n20",
            )
        if pack == "60":
            return make(
                sku="vitrum_immunaktiv_n60_tabs",
                canonical_name="ВИТРУМ ИММУНАКТИВ N60 ТАБ П.О.",
                line="Витрум Иммунактив",
                dosage_or_volume=mg_or_ml,
                form_factor="таблетки",
                flavor="",
                rule="immunaktiv_n60",
            )
        return make(
            sku="vitrum_immunaktiv_n30_tabs",
            canonical_name="ВИТРУМ ИММУНАКТИВ N30 ТАБ П.О.",
            line="Витрум Иммунактив",
            dosage_or_volume=mg_or_ml,
            form_factor="таблетки",
            flavor="",
            rule="immunaktiv_n30",
            confidence="high" if pack in {"", "30"} else "medium",
        )

    if "ЭНЕРДЖ" in norm:
        if "ШИП" in norm:
            return make(
                sku="vitrum_energy_effervescent_n20",
                canonical_name="ВИТРУМ ЭНЕРДЖИ ШИПУЧИЙ N20 ТАБ",
                line="Витрум Энерджи",
                dosage_or_volume=mg_or_ml,
                form_factor="шипучие таблетки",
                flavor="",
                rule="energy_effervescent_n20",
            )
        return make(
            sku="vitrum_energy_n30_tabs",
            canonical_name="ВИТРУМ ЭНЕРДЖИ N30 ТАБ П.О.",
            line="Витрум Энерджи",
            dosage_or_volume=mg_or_ml,
            form_factor="таблетки",
            flavor="",
            rule="energy_n30",
            confidence="high" if pack in {"", "30"} else "medium",
        )

    if "ВИТАЛИТИ" in norm or "50+" in norm:
        return make(
            sku="vitrum_vitality_50plus_n30_tabs",
            canonical_name="ВИТРУМ ВИТАЛИТИ 50+ N30 ТАБ П.О.",
            line="Витрум Виталити 50+",
            dosage_or_volume=mg_or_ml,
            form_factor="таблетки",
            flavor="",
            rule="vitality_50plus_n30",
            confidence="high" if pack in {"", "30"} else "medium",
        )

    has_vitamin_c = "ВИТАМИН С" in norm or "ВИТРУМ С" in norm
    if has_vitamin_c:
        if "ШИП" in norm or "900МГ" in norm:
            return make(
                sku="vitrum_vitamin_c_effervescent_900mg_n20",
                canonical_name="ВИТРУМ С ШИПУЧИЙ 900МГ N20 ТАБ ШИП",
                line="Витрум Витамин C",
                dosage_or_volume=mg_or_ml or "900 мг",
                form_factor="шипучие таблетки",
                flavor="",
                rule="vitamin_c_effervescent",
            )
        if "600МГ" in norm or "ЖЕВ" in norm or "АПЕЛЬСИН" in norm:
            return make(
                sku="vitrum_vitamin_c_orange_600mg_n30_chewable",
                canonical_name="ВИТРУМ ВИТАМИН С 600МГ N30 ТАБ ЖЕВ АПЕЛЬСИН",
                line="Витрум Витамин C",
                dosage_or_volume=mg_or_ml or "600 мг",
                form_factor="жевательные таблетки",
                flavor="апельсин",
                rule="vitamin_c_orange_chewable",
            )

    return make(
        sku="unknown_vitrum",
        canonical_name=name or "Unknown vitrum item",
        line="Unknown Vitrum",
        dosage_or_volume=mg_or_ml,
        form_factor="",
        flavor="",
        rule="fallback",
        confidence="low",
        is_vitrum=True,
    )


CYR_TO_LAT = str.maketrans(
    {
        "а": "a",
        "б": "b",
        "в": "v",
        "г": "g",
        "д": "d",
        "е": "e",
        "ё": "e",
        "ж": "zh",
        "з": "z",
        "и": "i",
        "й": "y",
        "к": "k",
        "л": "l",
        "м": "m",
        "н": "n",
        "о": "o",
        "п": "p",
        "р": "r",
        "с": "s",
        "т": "t",
        "у": "u",
        "ф": "f",
        "х": "h",
        "ц": "c",
        "ч": "ch",
        "ш": "sh",
        "щ": "sch",
        "ъ": "",
        "ы": "y",
        "ь": "",
        "э": "e",
        "ю": "yu",
        "я": "ya",
    }
)

GENERIC_BRAND_PROFILES: dict[str, dict[str, Any]] = {
    "enterogermina": {
        "display": "Энтерожермина",
        "aliases": ["энтерожермина", "энтерогермина", "enterogermina", "entero germina"],
    },
    "snoop": {
        "display": "Снуп",
        "aliases": ["снуп", "snoop"],
    },
    "aqualor": {
        "display": "Аквалор",
        "aliases": ["аквалор", "aqualor"],
    },
    "lazolvan": {
        "display": "Лазолван",
        "aliases": ["лазолван", "lazolvan", "lasolvan", "лазолван рино", "лазолван макс"],
    },
    "festal": {
        "display": "Фестал",
        "aliases": ["фестал", "festal"],
    },
    "no_shpa": {
        "display": "Но-Шпа",
        "aliases": ["но шпа", "но-шпа", "ношпа", "no-shpa", "noshpa", "no shpa"],
    },
    "coldrex": {
        "display": "Колдрекс",
        "aliases": ["колдрекс", "coldrex", "колдрекс хотрем", "coldrex hotrem", "hotrem", "колдрем"],
    },
    "magne_b6": {
        "display": "Магне B6",
        "aliases": [
            "магне в6",
            "магне б6",
            "магне-в6",
            "магне-б6",
            "магнеb6",
            "магнеv6",
            "магне б 6",
            "магне в 6",
            "magne b6",
            "magne-b6",
            "magneb6",
        ],
        "required_compact_groups": [["магне", "magne"], ["в6", "б6", "b6", "v6"]],
    },
    "ksefokam": {
        "display": "Ксефокам",
        "aliases": ["ксефокам", "ксефокам рапид", "ksefokam", "xefokam", "xefocam", "xefocam rapid"],
    },
    "new_nordic": {
        "display": "New Nordic",
        "aliases": [
            "new nordic",
            "newnordic",
            "нью нордик",
            "new nordic blue berry",
            "new nordic fish oil",
            "new nordic hair volume",
            "new nordic red oil",
            "new nordic multivitamin",
            "new nordic folkepilen",
            "new nordic apple cider",
            "new nordic active liver",
            "new nordic active memory",
            "нордик blue berry",
            "нордик fish oil",
            "нордик hair volume",
            "нордик red oil",
            "нордик multivitamin",
            "нордик folkepilen",
            "нордик apple cider",
        ],
        "weak_aliases": ["nordic", "нордик"],
        "weak_require_tokens": [
            "blue berry",
            "blueberry",
            "eyebright",
            "fish oil",
            "omega",
            "омега",
            "hair volume",
            "red oil",
            "multivitamin",
            "folkepilen",
            "folkepillen",
            "apple cider",
            "active liver",
            "active memory",
            "рыбий жир",
            "черник",
            "яблоч",
        ],
        "exclude_compact_tokens": ["naturals"],
    },
}

ACTIVE_INGREDIENT_REFERENCE_ROWS: list[dict[str, str]] = [
    {
        "brand_key": "vitrum",
        "brand_display": "Vitrum",
        "active_ingredient": "витаминно-минеральный комплекс (в зависимости от SKU)",
        "note": "Для отдельных SKU: D3, витамин C, магний+пиридоксин.",
        "source_url": "https://www.medicinform.net/instrukcii-lekarstv/vitrum-immunaktiv.html",
    },
    {
        "brand_key": "enterogermina",
        "brand_display": "Энтерожермина",
        "active_ingredient": "споры Bacillus clausii",
        "note": "Для форм Forte концентрация выше, но действующее начало то же.",
        "source_url": "https://drugs.africa/drugs/enterogermina-2-billion-5-ml-oral-suspension",
    },
    {
        "brand_key": "snoop",
        "brand_display": "Снуп",
        "active_ingredient": "ксилометазолин",
        "note": "Для Snoop Extra учитывается комбинация с декспантенолом.",
        "source_url": "https://www.vidal.ru/drugs/snup__39926",
    },
    {
        "brand_key": "aqualor",
        "brand_display": "Аквалор",
        "active_ingredient": "стерильная натуральная морская вода",
        "note": "Некоторые позиции содержат Aloe vera + римскую ромашку.",
        "source_url": "https://aqualor.ru/produkty/aqualor-soft",
    },
    {
        "brand_key": "lazolvan",
        "brand_display": "Лазолван",
        "active_ingredient": "амброксол",
        "note": "Для линейки Рино: трамазолин.",
        "source_url": "https://www.vidal.ru/drugs/lazolvan__10005",
    },
    {
        "brand_key": "festal",
        "brand_display": "Фестал",
        "active_ingredient": "панкреатин + гемицеллюлаза + компоненты желчи",
        "note": "Комбинированный ферментный препарат.",
        "source_url": "https://www.vidal.ru/drugs/festal__10963",
    },
    {
        "brand_key": "no_shpa",
        "brand_display": "Но-Шпа",
        "active_ingredient": "дротаверин",
        "note": "Для таблеток и инъекционной формы действующее вещество одно.",
        "source_url": "https://www.vidal.ru/drugs/no-spa__1131",
    },
    {
        "brand_key": "coldrex",
        "brand_display": "Колдрекс",
        "active_ingredient": "парацетамол + фенилэфрин + аскорбиновая кислота (линейка HotRem)",
        "note": "Состав зависит от конкретной подлинейки.",
        "source_url": "https://www.vidal.ru/drugs/coldrex-hotrem-lemon__20036",
    },
    {
        "brand_key": "magne_b6",
        "brand_display": "Магне B6",
        "active_ingredient": "магния лактат + пиридоксин",
        "note": "Для Premium дозировки отличаются, но пара активных веществ сохраняется.",
        "source_url": "https://www.vidal.kz/drugs/magne_b6__6537",
    },
    {
        "brand_key": "ksefokam",
        "brand_display": "Ксефокам",
        "active_ingredient": "лорноксикам",
        "note": "Ксефокам и Ксефокам Рапид содержат лорноксикам.",
        "source_url": "https://www.vidal.ru/drugs/ksefokam-rapid__43181",
    },
    {
        "brand_key": "new_nordic",
        "brand_display": "New Nordic",
        "active_ingredient": "мультикомпонентные БАД (зависят от SKU)",
        "note": "Линии Apple Cider / Blue Berry / Fish Oil имеют разные составы.",
        "source_url": (
            "https://newnordicusa.com/product/apple-cider/ | "
            "https://newnordicusa.com/product/blue-berry/ | "
            "https://oral.europharma.kz/bad-nordic-fish-oil-rybiy-zhir-700-mg-no60-kaps"
        ),
    },
]


def compact_token(value: str) -> str:
    return re.sub(r"[^a-z0-9а-яё]+", "", normalize_text(value))


def slugify_token(value: str) -> str:
    text = normalize_text(value).translate(CYR_TO_LAT)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def normalize_generic_name(value: str) -> str:
    text = (value or "").upper().replace("Ё", "Е")
    text = text.replace("®", " ").replace("™", " ")
    text = text.replace("№", " N").replace("#", " N")
    text = re.sub(r"\bNO(?=\s*\d)", "N", text)
    text = re.sub(r"\bN\s*(\d{1,4})\b", r"N\1", text)
    text = re.sub(r"[^0-9A-ZА-Я+/%.,]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def resolve_brand_profile(brand: str) -> tuple[str, dict[str, Any]]:
    brand_norm = normalize_text(brand)
    brand_compact = compact_token(brand)
    if brand_norm in {"vitrum", "витрум"}:
        return "vitrum", {"display": "Витрум", "aliases": ["витрум", "vitrum"]}

    for key, profile in GENERIC_BRAND_PROFILES.items():
        candidates = {compact_token(alias) for alias in profile.get("aliases", [])}
        candidates.add(compact_token(key))
        candidates.add(compact_token(profile.get("display", "")))
        if brand_compact in candidates:
            return key, profile

    dynamic_profile = {
        "display": brand.strip() or "Brand",
        "aliases": [brand.strip() or "brand"],
    }
    return slugify_token(brand) or "brand", dynamic_profile


def _alias_matches_item(alias: str, *, norm_item: str, compact_item: str) -> bool:
    alias_norm = normalize_generic_name(alias)
    alias_compact = compact_token(alias)
    if alias_norm and alias_norm in norm_item:
        return True
    if alias_compact and len(alias_compact) >= 4 and alias_compact in compact_item:
        return True
    return False


def _contains_compact_token(compact_item: str, token: str) -> bool:
    compact = compact_token(token)
    return bool(compact and compact in compact_item)


def _weak_brand_match(profile: dict[str, Any], *, norm_item: str, compact_item: str) -> bool:
    weak_aliases = [str(alias).strip() for alias in profile.get("weak_aliases", []) if str(alias).strip()]
    if not weak_aliases:
        return False
    if not any(_alias_matches_item(alias, norm_item=norm_item, compact_item=compact_item) for alias in weak_aliases):
        return False

    weak_require_tokens = [str(token).strip() for token in profile.get("weak_require_tokens", []) if str(token).strip()]
    if not weak_require_tokens:
        return True
    return any(
        normalize_generic_name(token) in norm_item or _contains_compact_token(compact_item, token)
        for token in weak_require_tokens
    )


def _required_compact_groups_match(profile: dict[str, Any], compact_item: str) -> bool:
    groups = profile.get("required_compact_groups", [])
    if not groups:
        return True
    for group in groups:
        tokens = [str(token).strip() for token in group if str(token).strip()]
        if not tokens:
            continue
        if not any(_contains_compact_token(compact_item, token) for token in tokens):
            return False
    return True


def _excluded_compact_tokens_match(profile: dict[str, Any], compact_item: str) -> bool:
    excluded = [str(token).strip() for token in profile.get("exclude_compact_tokens", []) if str(token).strip()]
    return any(_contains_compact_token(compact_item, token) for token in excluded)


def _format_number(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.3f}".rstrip("0").rstrip(".")


def _looks_like_route_token(token: str) -> bool:
    return bool(
        token
        and (
            "/" in token
            or token in {"Р", "РР", "Д", "И", "В", "М", "ВВ", "ВМ", "П", "О", "ПП", "ПЛЕН"}
            or token.startswith(("Д/", "П/", "Р/"))
        )
    )


def infer_generic_product_line(
    *,
    brand_key: str,
    norm: str,
    stripped: str,
    dosage_or_volume: str,
    form_factor: str,
) -> str:
    if brand_key == "ksefokam":
        if "РАПИД" in norm:
            return "РАПИД"
        if any(token in norm for token in ("ЛИОФ", "ИНЪ", "В/В", "В/М", "АМП")):
            return "ИНЪЕКЦИИ"
        return "ТАБЛЕТКИ"

    if brand_key == "no_shpa":
        if any(token in norm for token in ("Р-Р", "РАСТВОР", "Д/И", "В/В", "В/М", "АМП")):
            return "РАСТВОР ДЛЯ ИНЪЕКЦИЙ"
        if "ФОРТЕ" in norm or dosage_or_volume.startswith("80 "):
            return "ФОРТЕ"
        return "ТАБЛЕТКИ"

    if brand_key == "lazolvan":
        if "РИНО" in norm:
            return "РИНО"
        if "МАКС" in norm:
            return "МАКС"
        if "ЮНИОР" in norm:
            return "ЮНИОР"
        if any(token in norm for token in ("ИНГ", "ВНУТР", "РАСТВОР")):
            return "РАСТВОР"
        if "СИРОП" in norm:
            return "СИРОП"
        if "ТАБ" in norm:
            return "ТАБЛЕТКИ"
        return "БАЗА"

    if brand_key == "magne_b6":
        if "ПРЕМИУМ" in norm:
            return "ПРЕМИУМ"
        if any(token in norm for token in ("РАСТВОР", "АМП", "ПИТЬЕВОЙ")):
            return "РАСТВОР"
        return "КЛАССИК"

    if brand_key == "coldrex":
        if "КОЛДРЕМ" in norm or "HOTREM" in norm:
            return "КОЛДРЕМ"
        if "МАКСГРИПП" in norm:
            return "МАКСГРИПП"
        if "БРОНХО" in norm:
            return "БРОНХО"
        return "БАЗА"

    if brand_key == "snoop":
        if "ЭКСТРА" in norm:
            return "ЭКСТРА"
        return "КЛАССИК"

    if brand_key == "enterogermina":
        if "ФОРТЕ" in norm:
            return "ФОРТЕ"
        if "КАПС" in norm:
            return "КАПСУЛЫ"
        return "СУСПЕНЗИЯ"

    if brand_key == "festal":
        return "Н" if re.search(r"\bН\b", norm) else "БАЗА"

    if brand_key == "aqualor":
        if "АКТИВ ФОРТЕ" in norm:
            return "АКТИВ ФОРТЕ"
        if "АКТИВ СОФТ" in norm:
            return "АКТИВ СОФТ"
        if "ЭКСТРА ФОРТЕ" in norm:
            return "ЭКСТРА ФОРТЕ"
        if "ГОРЛО" in norm:
            return "ГОРЛО"
        if "БЕБИ" in norm:
            return "БЕБИ"
        if "СОФТ" in norm:
            return "СОФТ"
        if "НОРМ" in norm:
            return "НОРМ"
        if "ПРОТЕКТ" in norm:
            return "ПРОТЕКТ"
        return "БАЗА"

    if brand_key == "new_nordic":
        if any(token in norm for token in ("АППЛСИДЕР", "APPLE", "ЯБЛОЧ", "APPL")):
            return "APPLE CIDER"
        if any(token in norm for token in ("БЛЮБЕР", "BLUE BERRY", "BLUEBERRY", "ЧЕРНИК")):
            return "BLUE BERRY"
        if any(token in norm for token in ("FISH OIL", "РЫБИЙ", "ОМЕГА")):
            return "FISH OIL"
        return "БАЗА"

    stop_tokens = {
        "N",
        "ТАБ",
        "ТАБЛ",
        "ТАБЛЕТКИ",
        "КАПС",
        "КАПСУЛЫ",
        "САШЕ",
        "ШТ",
        "СПРЕЙ",
        "СИРОП",
        "РАСТВОР",
        "ПОРОШОК",
        "ДЛЯ",
        "И",
        "СО",
        "ВКУСОМ",
        "ПАКЕТ",
        "ПАКЕТИК",
        "ДУШ",
        "НАЗ",
    }
    tokens: list[str] = []
    for token in stripped.split():
        if not token or token in stop_tokens:
            continue
        if re.fullmatch(r"N?\d{1,4}", token):
            continue
        if re.fullmatch(r"\d+(?:[.,]\d+)?(?:МГ|МЛ|Г|МЕ|IU|%)?", token):
            continue
        if _looks_like_route_token(token):
            continue
        tokens.append(token)
    return " ".join(tokens[:6]).strip() or (form_factor.upper() if form_factor else "БАЗА")


def infer_generic_form_factor(
    *,
    brand_key: str,
    norm: str,
    product_line: str,
    extracted_form: str,
) -> str:
    if brand_key == "ksefokam":
        if product_line == "ИНЪЕКЦИИ":
            return "инъекционная форма"
        return "таблетки"
    if brand_key == "no_shpa":
        if product_line == "РАСТВОР ДЛЯ ИНЪЕКЦИЙ":
            return "раствор"
        return "таблетки"
    if brand_key == "snoop":
        return "спрей"
    if brand_key == "festal":
        return "таблетки"
    if brand_key == "coldrex":
        if "КОЛДРЕМ" in product_line or "МАКСГРИПП" in product_line:
            return "порошок/саше"
    if brand_key == "enterogermina":
        if "КАПС" in product_line:
            return "капсулы"
        return "суспензия"
    if brand_key == "aqualor":
        if "КАПЛ" in norm:
            return "капли"
        if "ГОРЛО" in product_line:
            return "спрей для горла"
        return "спрей"
    if brand_key == "lazolvan":
        if "ТАБ" in norm:
            return "таблетки"
        if "КАПС" in norm:
            return "капсулы"
        if "СИРОП" in norm:
            return "сироп"
        if any(token in norm for token in ("ИНГ", "ВНУТР", "РАСТВОР")):
            return "раствор"
    return extracted_form


def normalize_brand_dosage(
    *,
    brand_key: str,
    norm: str,
    product_line: str,
    dosage_or_volume: str,
) -> str:
    dosage = (dosage_or_volume or "").strip()
    if brand_key == "enterogermina":
        if product_line == "ФОРТЕ" and dosage in {"", "5 мл"}:
            return "4 млрд/5 мл"
        if product_line == "СУСПЕНЗИЯ" and dosage in {"", "5 мл"}:
            return "2 млрд/5 мл"
    if brand_key == "no_shpa" and product_line == "РАСТВОР ДЛЯ ИНЪЕКЦИЙ":
        if any(token in norm for token in ("40МГ/2МЛ", "2%", "2МЛ")):
            return "40 мг/2 мл"
        if dosage in {"40 мг", "2 мл", "", "2 %"}:
            return "40 мг/2 мл"
    return dosage


def resolve_brand_active_ingredient(brand_key: str, norm: str, product_line: str) -> tuple[str, str]:
    if brand_key == "ksefokam":
        return "лорноксикам", "high"
    if brand_key == "no_shpa":
        return "дротаверин", "high"
    if brand_key == "lazolvan":
        if "РИНО" in norm:
            return "трамазолин", "high"
        return "амброксол", "high"
    if brand_key == "snoop":
        if "ЭКСТРА" in norm:
            return "ксилометазолин + декспантенол", "medium"
        return "ксилометазолин", "high"
    if brand_key == "festal":
        return "панкреатин + гемицеллюлаза + компоненты желчи", "high"
    if brand_key == "magne_b6":
        return "магния лактат + пиридоксин", "high"
    if brand_key == "enterogermina":
        return "споры Bacillus clausii", "high"
    if brand_key == "aqualor":
        if "АЛОЭ" in norm or "РОМАШ" in norm:
            return "стерильная морская вода + Aloe vera + римская ромашка", "high"
        return "стерильная морская вода", "high"
    if brand_key == "coldrex":
        if "БРОНХО" in norm:
            return "гуайфенезин", "high"
        if "КОЛДРЕМ" in norm or "HOTREM" in norm or "МАКСГРИПП" in norm:
            return "парацетамол + фенилэфрин + аскорбиновая кислота", "high"
        return "комбинированный состав (линейка Колдрекс)", "medium"
    if brand_key == "new_nordic":
        if "APPLE CIDER" in product_line:
            return "яблочный уксус + экстракт артишока + экстракт одуванчика", "high"
        if "BLUE BERRY" in product_line:
            return "экстракт черники + лютеин + очанка + экстракт виноградных косточек", "high"
        if "FISH OIL" in product_line:
            return "рыбий жир (омега-3)", "high"
        return "комплексный БАД (состав зависит от SKU)", "medium"
    return "", "low"


def export_active_ingredient_reference(path: Path) -> None:
    checked_at = datetime.now().date().isoformat()
    rows = [
        {
            "checked_at": checked_at,
            **row,
        }
        for row in ACTIVE_INGREDIENT_REFERENCE_ROWS
    ]
    write_csv(
        path,
        rows,
        [
            "checked_at",
            "brand_key",
            "brand_display",
            "active_ingredient",
            "note",
            "source_url",
        ],
    )


def _extract_pack_size_generic(text: str) -> str:
    m = re.search(r"\bN(\d{1,4})\b", text)
    if m:
        return m.group(1)
    m = re.search(r"\b(\d{1,4})\s*(ТАБ|ТАБЛ|ТАБЛЕТ|TABS?|КАПС|CAPS?|ШТ|PCS|САШЕ|SACHETS?|АМП)\b", text)
    if m:
        return m.group(1)
    return ""


def _extract_dosage_generic(text: str) -> str:
    billion_ratio = re.search(r"\b(\d{1,3}(?:[.,]\d+)?)\s*МЛРД\s*/\s*(\d{1,3}(?:[.,]\d+)?)\s*(МЛ|ML)\b", text)
    if billion_ratio:
        left = _format_number(float(billion_ratio.group(1).replace(",", ".")))
        right = _format_number(float(billion_ratio.group(2).replace(",", ".")))
        return f"{left} млрд/{right} мл"

    ratio = re.search(r"\b(\d{1,4}(?:[.,]\d+)?)\s*(МГ|MG)\s*/\s*(\d{1,4}(?:[.,]\d+)?)\s*(МЛ|ML)\b", text)
    if ratio:
        left = _format_number(float(ratio.group(1).replace(",", ".")))
        right = _format_number(float(ratio.group(3).replace(",", ".")))
        return f"{left} мг/{right} мл"

    compact_ratio = re.search(r"\b(0[.,]\d+)\s*/\s*(\d{1,4}(?:[.,]\d+)?)\s*(МЛ|ML)\b", text)
    if compact_ratio:
        left_g = float(compact_ratio.group(1).replace(",", "."))
        right = _format_number(float(compact_ratio.group(2).replace(",", ".")))
        left_mg = _format_number(left_g * 1000.0)
        return f"{left_mg} мг/{right} мл"

    dose = re.search(r"\b(\d{1,4}(?:[.,]\d+)?)\s*(МКГ|MCG|МГ|MG|Г|G|МЛ|ML|МЕ|IU|%)\b", text)
    if not dose:
        return ""
    value = _format_number(float(dose.group(1).replace(",", ".")))
    unit_raw = dose.group(2).upper()
    unit_map = {
        "МКГ": "мкг",
        "MCG": "мкг",
        "МГ": "мг",
        "MG": "мг",
        "Г": "г",
        "G": "г",
        "МЛ": "мл",
        "ML": "мл",
        "МЕ": "МЕ",
        "IU": "IU",
        "%": "%",
    }
    unit = unit_map.get(unit_raw, unit_raw.lower())
    return f"{value} {unit}".strip()


def _extract_form_generic(text: str) -> str:
    rules = [
        (r"\bШИП", "шипучие таблетки"),
        (r"\bСПРЕЙ\b", "спрей"),
        (r"\bСИРОП\b", "сироп"),
        (r"\bЛИОФ", "лиофилизат"),
        (r"\bСУСП", "суспензия"),
        (r"\bР-Р\b|\bРАСТВОР\b", "раствор"),
        (r"\bИНГАЛ", "раствор для ингаляций"),
        (r"\bАМП", "ампулы"),
        (r"\bТАБ|ТАБЛ", "таблетки"),
        (r"\bКАПС", "капсулы"),
        (r"\bПАСТИЛ|ЛЕДЕНЦ", "пастилки"),
        (r"\bСАШЕ|SACHET|ПОРОШ", "порошок/саше"),
        (r"\bКАПЛ", "капли"),
        (r"\bГЕЛЬ", "гель"),
        (r"\bКРЕМ", "крем"),
        (r"\bМАЗ", "мазь"),
    ]
    for pattern, label in rules:
        if re.search(pattern, text):
            return label
    return ""


def _extract_flavor_generic(text: str) -> str:
    flavors = [
        ("АПЕЛЬСИН", "апельсин"),
        ("ЛИМОН", "лимон"),
        ("МЯТА", "мята"),
        ("МЕНТОЛ", "ментол"),
        ("МАЛИН", "малина"),
        ("ЯБЛОК", "яблоко"),
        ("КЛУБНИК", "клубника"),
        ("СМОРОДИН", "смородина"),
        ("ВИШН", "вишня"),
    ]
    for token, label in flavors:
        if token in text:
            return label
    return ""


def canonicalize_generic_brand_name(brand: str, raw_name: str) -> dict[str, str]:
    brand_key, profile = resolve_brand_profile(brand)
    display_name = str(profile.get("display") or brand or "Brand").strip()
    name = (raw_name or "").strip()
    norm = normalize_generic_name(name)
    aliases = [normalize_generic_name(alias) for alias in profile.get("aliases", []) if alias]
    aliases = [alias for alias in aliases if alias]
    if not aliases:
        aliases = [normalize_generic_name(display_name)]

    stripped = norm
    brand_match = False
    for alias in aliases:
        if not alias:
            continue
        if re.search(rf"\b{re.escape(alias)}\b", stripped):
            brand_match = True
        stripped = re.sub(rf"\b{re.escape(alias)}\b", " ", stripped)
    stripped = re.sub(r"\s+", " ", stripped).strip()

    pack_size = _extract_pack_size_generic(norm)
    dosage_or_volume = _extract_dosage_generic(norm)
    form_factor = _extract_form_generic(norm)
    flavor = _extract_flavor_generic(norm)
    product_line = infer_generic_product_line(
        brand_key=brand_key,
        norm=norm,
        stripped=stripped,
        dosage_or_volume=dosage_or_volume,
        form_factor=form_factor,
    )
    dosage_or_volume = normalize_brand_dosage(
        brand_key=brand_key,
        norm=norm,
        product_line=product_line,
        dosage_or_volume=dosage_or_volume,
    )
    form_factor = infer_generic_form_factor(
        brand_key=brand_key,
        norm=norm,
        product_line=product_line,
        extracted_form=form_factor,
    )
    active_ingredient, ingredient_confidence = resolve_brand_active_ingredient(brand_key, norm, product_line)

    line_slug = slugify_token(product_line) or "base"
    ingredient_slug = slugify_token(active_ingredient)
    dosage_slug = slugify_token(dosage_or_volume)
    form_slug = slugify_token(form_factor)
    pack_slug = f"n{pack_size}" if pack_size else ""
    canonical_sku_parts = [slugify_token(display_name) or brand_key, line_slug]
    if ingredient_slug:
        canonical_sku_parts.append(ingredient_slug)
    if dosage_slug:
        canonical_sku_parts.append(dosage_slug)
    if pack_slug:
        canonical_sku_parts.append(pack_slug)
    if form_slug:
        canonical_sku_parts.append(form_slug)
    canonical_sku = "_".join(part for part in canonical_sku_parts if part) or f"{brand_key}_unknown"

    canonical_name_parts = [display_name.upper(), product_line]
    if dosage_or_volume:
        canonical_name_parts.append(dosage_or_volume.upper())
    if pack_size:
        canonical_name_parts.append(f"N{pack_size}")
    if form_factor:
        canonical_name_parts.append(form_factor.upper())
    canonical_name = " ".join(part for part in canonical_name_parts if part).strip()

    if not brand_match:
        confidence = "low"
        rule = "generic_missing_brand_token"
        canonical_sku = f"unknown_{slugify_token(display_name) or brand_key}"
    elif product_line and (pack_size or dosage_or_volume):
        confidence = "high"
        rule = "generic_brand_with_pack_or_dosage"
    else:
        confidence = "medium"
        rule = "generic_brand_parser"
    if ingredient_confidence == "medium" and confidence == "high":
        confidence = "medium"
    if ingredient_confidence == "low":
        confidence = "low" if not brand_match else "medium"
        rule = "generic_unknown_active_ingredient" if brand_match else rule

    return _vitrum_make_result(
        raw_name=name,
        normalized_name=norm,
        canonical_sku=canonical_sku,
        canonical_name=canonical_name or (name or display_name),
        product_line=product_line,
        dosage_or_volume=dosage_or_volume,
        pack_size=pack_size,
        form_factor=form_factor,
        flavor=flavor,
        active_ingredient=active_ingredient,
        confidence=confidence,
        rule=rule,
        is_vitrum=False,
    )


def canonicalize_brand_name(brand: str, raw_name: str) -> dict[str, str]:
    brand_norm = normalize_text(brand)
    if brand_norm == "vitrum" or "витрум" in brand_norm:
        return canonicalize_vitrum_name(raw_name)
    return canonicalize_generic_brand_name(brand, raw_name)


def item_matches_brand_name(brand: str, item_name: str) -> bool:
    _, profile = resolve_brand_profile(brand)
    aliases = list(profile.get("aliases", []))
    display = str(profile.get("display") or "").strip()
    if display:
        aliases.append(display)
    aliases.append(brand)

    norm_item = normalize_generic_name(item_name)
    compact_item = compact_token(item_name)
    if _excluded_compact_tokens_match(profile, compact_item):
        return False
    if not _required_compact_groups_match(profile, compact_item):
        return False

    for alias in aliases:
        if _alias_matches_item(alias, norm_item=norm_item, compact_item=compact_item):
            return True
    return _weak_brand_match(profile, norm_item=norm_item, compact_item=compact_item)


def build_vitrum_reference_rows(
    catalog: dict[tuple[str, str], dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    detailed_rows: list[dict[str, Any]] = []
    for rec in catalog.values():
        item_name = str(rec.get("item_name") or "").strip()
        brand = str(rec.get("brand") or "").strip()
        if "vitrum" not in normalize_text(brand) and "витрум" not in normalize_text(item_name):
            continue

        canonical = canonicalize_vitrum_name(item_name)
        row = {
            "brand": brand,
            "item_id": str(rec.get("item_id") or "").strip(),
            "item_name": item_name,
            "first_seen": str(rec.get("first_seen") or "").strip(),
            "last_seen": str(rec.get("last_seen") or "").strip(),
            "seen_runs": _to_int(rec.get("seen_runs")) or 0,
            "last_seen_venue_slug": str(rec.get("last_seen_venue_slug") or "").strip(),
            "last_seen_pharmacy": str(rec.get("last_seen_pharmacy") or "").strip(),
            "canonical_sku": canonical["canonical_sku"],
            "canonical_name": canonical["canonical_name"],
            "product_line": canonical["product_line"],
            "dosage_or_volume": canonical["dosage_or_volume"],
            "pack_size": canonical["pack_size"],
            "form_factor": canonical["form_factor"],
            "flavor": canonical["flavor"],
            "active_ingredient": canonical["active_ingredient"],
            "confidence": canonical["confidence"],
            "rule": canonical["rule"],
            "is_vitrum": canonical["is_vitrum"],
            "normalized_name": canonical["normalized_name"],
        }
        detailed_rows.append(row)

    detailed_rows.sort(key=lambda x: (x.get("canonical_sku", ""), normalize_text(str(x.get("item_name", ""))), str(x.get("item_id", ""))))

    canonical_map: dict[str, dict[str, Any]] = {}
    for row in detailed_rows:
        key = str(row.get("canonical_sku") or "")
        bucket = canonical_map.setdefault(
            key,
            {
                "canonical_sku": key,
                "canonical_name": row.get("canonical_name") or "",
                "product_line": row.get("product_line") or "",
                "dosage_or_volume": row.get("dosage_or_volume") or "",
                "pack_size": row.get("pack_size") or "",
                "form_factor": row.get("form_factor") or "",
                "flavor": row.get("flavor") or "",
                "active_ingredient": row.get("active_ingredient") or "",
                "item_ids_count": 0,
                "aliases_count": 0,
                "seen_runs_total": 0,
                "confidence_min": row.get("confidence") or "",
                "_item_ids": set(),
                "_aliases": set(),
                "_rules": set(),
            },
        )
        item_id = str(row.get("item_id") or "").strip()
        item_name = str(row.get("item_name") or "").strip()
        if item_id:
            bucket["_item_ids"].add(item_id)
        if item_name:
            bucket["_aliases"].add(item_name)
        rule = str(row.get("rule") or "").strip()
        if rule:
            bucket["_rules"].add(rule)
        bucket["seen_runs_total"] += _to_int(row.get("seen_runs")) or 0

        if row.get("confidence") == "low":
            bucket["confidence_min"] = "low"
        elif row.get("confidence") == "medium" and bucket["confidence_min"] == "high":
            bucket["confidence_min"] = "medium"

    canonical_rows: list[dict[str, Any]] = []
    for entry in canonical_map.values():
        item_ids = sorted(entry.pop("_item_ids"))
        aliases = sorted(entry.pop("_aliases"))
        rules = sorted(entry.pop("_rules"))
        entry["item_ids_count"] = len(item_ids)
        entry["aliases_count"] = len(aliases)
        entry["item_ids"] = " | ".join(item_ids)
        entry["aliases_examples"] = " || ".join(aliases[:8])
        entry["rules_used"] = ",".join(rules)
        canonical_rows.append(entry)
    canonical_rows.sort(key=lambda x: normalize_text(str(x.get("canonical_name", ""))))

    unmapped_rows = [row for row in detailed_rows if str(row.get("canonical_sku")) in {"unknown_vitrum", "unknown_non_vitrum"}]
    return detailed_rows, canonical_rows, unmapped_rows


def export_vitrum_reference(
    catalog: dict[tuple[str, str], dict[str, Any]],
    reference_path: Path,
    canonical_path: Path,
    unmapped_path: Path,
) -> dict[str, Any]:
    detailed_rows, canonical_rows, unmapped_rows = build_vitrum_reference_rows(catalog)

    write_csv(
        reference_path,
        detailed_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "first_seen",
            "last_seen",
            "seen_runs",
            "last_seen_venue_slug",
            "last_seen_pharmacy",
            "canonical_sku",
            "canonical_name",
            "product_line",
            "dosage_or_volume",
            "pack_size",
            "form_factor",
            "flavor",
            "active_ingredient",
            "confidence",
            "rule",
            "is_vitrum",
            "normalized_name",
        ],
    )
    write_csv(
        canonical_path,
        canonical_rows,
        [
            "canonical_sku",
            "canonical_name",
            "product_line",
            "dosage_or_volume",
            "pack_size",
            "form_factor",
            "flavor",
            "active_ingredient",
            "item_ids_count",
            "aliases_count",
            "seen_runs_total",
            "confidence_min",
            "rules_used",
            "item_ids",
            "aliases_examples",
        ],
    )
    write_csv(
        unmapped_path,
        unmapped_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "canonical_sku",
            "canonical_name",
            "active_ingredient",
            "confidence",
            "rule",
            "normalized_name",
            "first_seen",
            "last_seen",
            "last_seen_pharmacy",
        ],
    )
    return {
        "reference_path": reference_path,
        "canonical_path": canonical_path,
        "unmapped_path": unmapped_path,
        "detailed_count": len(detailed_rows),
        "canonical_count": len(canonical_rows),
        "unmapped_count": len(unmapped_rows),
    }


def build_brand_reference_rows(
    catalog: dict[tuple[str, str], dict[str, Any]],
    brand: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    brand_norm = normalize_text(brand)
    detailed_rows: list[dict[str, Any]] = []
    for rec in catalog.values():
        rec_brand = str(rec.get("brand") or "").strip()
        if normalize_text(rec_brand) != brand_norm:
            continue

        item_name = str(rec.get("item_name") or "").strip()
        canonical = canonicalize_brand_name(rec_brand, item_name)
        row = {
            "brand": rec_brand,
            "item_id": str(rec.get("item_id") or "").strip(),
            "item_name": item_name,
            "first_seen": str(rec.get("first_seen") or "").strip(),
            "last_seen": str(rec.get("last_seen") or "").strip(),
            "seen_runs": _to_int(rec.get("seen_runs")) or 0,
            "last_seen_venue_slug": str(rec.get("last_seen_venue_slug") or "").strip(),
            "last_seen_pharmacy": str(rec.get("last_seen_pharmacy") or "").strip(),
            "canonical_sku": canonical["canonical_sku"],
            "canonical_name": canonical["canonical_name"],
            "product_line": canonical["product_line"],
            "dosage_or_volume": canonical["dosage_or_volume"],
            "pack_size": canonical["pack_size"],
            "form_factor": canonical["form_factor"],
            "flavor": canonical["flavor"],
            "active_ingredient": canonical["active_ingredient"],
            "confidence": canonical["confidence"],
            "rule": canonical["rule"],
            "normalized_name": canonical["normalized_name"],
        }
        detailed_rows.append(row)

    detailed_rows.sort(
        key=lambda x: (
            x.get("canonical_sku", ""),
            normalize_text(str(x.get("item_name", ""))),
            str(x.get("item_id", "")),
        )
    )

    canonical_map: dict[str, dict[str, Any]] = {}
    for row in detailed_rows:
        key = str(row.get("canonical_sku") or "")
        if not key:
            key = "unknown"
        bucket = canonical_map.setdefault(
            key,
            {
                "canonical_sku": key,
                "canonical_name": row.get("canonical_name") or "",
                "product_line": row.get("product_line") or "",
                "dosage_or_volume": row.get("dosage_or_volume") or "",
                "pack_size": row.get("pack_size") or "",
                "form_factor": row.get("form_factor") or "",
                "flavor": row.get("flavor") or "",
                "active_ingredient": row.get("active_ingredient") or "",
                "item_ids_count": 0,
                "aliases_count": 0,
                "seen_runs_total": 0,
                "confidence_min": row.get("confidence") or "",
                "_item_ids": set(),
                "_aliases": set(),
                "_rules": set(),
            },
        )
        item_id = str(row.get("item_id") or "").strip()
        item_name = str(row.get("item_name") or "").strip()
        if item_id:
            bucket["_item_ids"].add(item_id)
        if item_name:
            bucket["_aliases"].add(item_name)
        rule = str(row.get("rule") or "").strip()
        if rule:
            bucket["_rules"].add(rule)
        bucket["seen_runs_total"] += _to_int(row.get("seen_runs")) or 0

        if row.get("confidence") == "low":
            bucket["confidence_min"] = "low"
        elif row.get("confidence") == "medium" and bucket["confidence_min"] == "high":
            bucket["confidence_min"] = "medium"

    canonical_rows: list[dict[str, Any]] = []
    for entry in canonical_map.values():
        item_ids = sorted(entry.pop("_item_ids"))
        aliases = sorted(entry.pop("_aliases"))
        rules = sorted(entry.pop("_rules"))
        entry["item_ids_count"] = len(item_ids)
        entry["aliases_count"] = len(aliases)
        entry["item_ids"] = " | ".join(item_ids)
        entry["aliases_examples"] = " || ".join(aliases[:8])
        entry["rules_used"] = ",".join(rules)
        canonical_rows.append(entry)
    canonical_rows.sort(key=lambda x: normalize_text(str(x.get("canonical_name", ""))))

    unmapped_rows = [
        row
        for row in detailed_rows
        if str(row.get("canonical_sku")).startswith("unknown_") or str(row.get("confidence")) == "low"
    ]
    return detailed_rows, canonical_rows, unmapped_rows


def export_brand_reference(
    catalog: dict[tuple[str, str], dict[str, Any]],
    brand: str,
    reference_path: Path,
    canonical_path: Path,
    unmapped_path: Path,
) -> dict[str, Any]:
    detailed_rows, canonical_rows, unmapped_rows = build_brand_reference_rows(catalog=catalog, brand=brand)

    write_csv(
        reference_path,
        detailed_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "first_seen",
            "last_seen",
            "seen_runs",
            "last_seen_venue_slug",
            "last_seen_pharmacy",
            "canonical_sku",
            "canonical_name",
            "product_line",
            "dosage_or_volume",
            "pack_size",
            "form_factor",
            "flavor",
            "active_ingredient",
            "confidence",
            "rule",
            "normalized_name",
        ],
    )
    write_csv(
        canonical_path,
        canonical_rows,
        [
            "canonical_sku",
            "canonical_name",
            "product_line",
            "dosage_or_volume",
            "pack_size",
            "form_factor",
            "flavor",
            "active_ingredient",
            "item_ids_count",
            "aliases_count",
            "seen_runs_total",
            "confidence_min",
            "rules_used",
            "item_ids",
            "aliases_examples",
        ],
    )
    write_csv(
        unmapped_path,
        unmapped_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "canonical_sku",
            "canonical_name",
            "active_ingredient",
            "confidence",
            "rule",
            "normalized_name",
            "first_seen",
            "last_seen",
            "last_seen_pharmacy",
        ],
    )
    return {
        "reference_path": reference_path,
        "canonical_path": canonical_path,
        "unmapped_path": unmapped_path,
        "detailed_count": len(detailed_rows),
        "canonical_count": len(canonical_rows),
        "unmapped_count": len(unmapped_rows),
    }


def build_best_offers(item_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    best: dict[str, dict[str, Any]] = {}
    for row in item_rows:
        if str(row.get("status")) != "in_stock":
            continue
        price_minor = _to_int(row.get("price_minor"))
        if price_minor is None:
            continue
        key = normalize_text(str(row.get("item_name") or ""))
        if not key:
            continue
        prev = best.get(key)
        if prev is None or price_minor < _to_int(prev.get("price_minor") or 10**12):
            best[key] = row
    offers = list(best.values())
    offers.sort(key=lambda row: (_to_int(row.get("price_minor")) or 10**12, normalize_text(str(row.get("item_name", "")))))
    return offers


def _split_text(text: str, chunk_size: int = 3500) -> list[str]:
    text = text or ""
    if len(text) <= chunk_size:
        return [text]
    chunks: list[str] = []
    start = 0
    while start < len(text):
        end = min(len(text), start + chunk_size)
        cut = text.rfind("\n", start, end)
        if cut <= start:
            cut = end
        chunks.append(text[start:cut].strip())
        start = cut
    return [c for c in chunks if c]


def send_telegram_message(token: str, chat_id: str, text: str) -> None:
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    for chunk in _split_text(text):
        resp = requests.post(url, data={"chat_id": chat_id, "text": chunk}, timeout=30)
        resp.raise_for_status()


def send_telegram_document(token: str, chat_id: str, file_path: Path, caption: str | None = None) -> None:
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    with open(file_path, "rb") as fh:
        files = {"document": (file_path.name, fh)}
        data = {"chat_id": chat_id}
        if caption:
            data["caption"] = caption
        resp = requests.post(url, data=data, files=files, timeout=60)
    resp.raise_for_status()


def send_telegram_photo(token: str, chat_id: str, photo_path: Path, caption: str | None = None) -> None:
    url = f"https://api.telegram.org/bot{token}/sendPhoto"
    with open(photo_path, "rb") as fh:
        files = {"photo": (photo_path.name, fh)}
        data = {"chat_id": chat_id}
        if caption:
            data["caption"] = caption
        resp = requests.post(url, data=data, files=files, timeout=60)
    resp.raise_for_status()


def create_visual_summary(
    out_path: Path,
    brand: str,
    query: str,
    checked_at: str,
    pharmacies_count: int,
    item_rows_count: int,
    in_stock_count: int,
    out_of_stock_count: int,
    not_found_pharmacies: int,
    new_ids_count: int,
    top_offers: list[dict[str, Any]],
) -> bool:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except Exception:
        return False

    width = 1500
    row_h = 42
    top_n = min(len(top_offers), 12)
    height = 270 + top_n * row_h + 50
    img = Image.new("RGB", (width, height), "#F5F7FB")
    draw = ImageDraw.Draw(img)

    font_regular = ImageFont.load_default()
    font_bold = ImageFont.load_default()
    if DEFAULT_FONT_PATH.exists():
        try:
            font_regular = ImageFont.truetype(str(DEFAULT_FONT_PATH), 22)
            font_bold = ImageFont.truetype(str(DEFAULT_FONT_PATH), 30)
        except Exception:
            pass

    draw.rectangle((0, 0, width, 95), fill="#0E7490")
    draw.text((28, 26), f"Wolt Price Monitor | {brand}", fill="white", font=font_bold)
    draw.text((28, 62), f"query: {query} | {checked_at}", fill="white", font=font_regular)

    cards = [
        f"Pharmacies: {pharmacies_count}",
        f"Items found: {item_rows_count}",
        f"In stock: {in_stock_count}",
        f"Out of stock: {out_of_stock_count}",
        f"No matches: {not_found_pharmacies}",
        f"New item_id: {new_ids_count}",
    ]
    x = 26
    y = 118
    for card in cards:
        draw.rounded_rectangle((x, y, x + 230, y + 58), radius=12, fill="white", outline="#D8DEE9")
        draw.text((x + 14, y + 17), card, fill="#1F2937", font=font_regular)
        x += 244

    draw.text((28, 196), "Top best prices (in-stock)", fill="#111827", font=font_bold)
    y0 = 230
    draw.rectangle((24, y0, width - 24, y0 + row_h), fill="#E2E8F0")
    draw.text((36, y0 + 10), "#", fill="#111827", font=font_regular)
    draw.text((74, y0 + 10), "Product", fill="#111827", font=font_regular)
    draw.text((970, y0 + 10), "Price", fill="#111827", font=font_regular)
    draw.text((1120, y0 + 10), "Pharmacy", fill="#111827", font=font_regular)

    for idx, row in enumerate(top_offers[:top_n], start=1):
        y = y0 + idx * row_h
        fill = "#FFFFFF" if idx % 2 else "#F8FAFC"
        draw.rectangle((24, y, width - 24, y + row_h), fill=fill)

        name = str(row.get("item_name") or "")
        if len(name) > 62:
            name = name[:59] + "..."
        pharmacy = str(row.get("pharmacy") or "")
        if len(pharmacy) > 40:
            pharmacy = pharmacy[:37] + "..."
        price_text = str(row.get("price") or "")
        if price_text:
            price_text = f"{price_text} KZT"

        draw.text((36, y + 10), str(idx), fill="#111827", font=font_regular)
        draw.text((74, y + 10), name, fill="#111827", font=font_regular)
        draw.text((970, y + 10), price_text, fill="#111827", font=font_regular)
        draw.text((1120, y + 10), pharmacy, fill="#111827", font=font_regular)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path, format="PNG")
    return True


def build_telegram_text(
    brand: str,
    query: str,
    checked_at: str,
    pharmacies_count: int,
    item_rows: list[dict[str, Any]],
    summary_rows: list[dict[str, Any]],
    new_ids_count: int,
    top_offers: list[dict[str, Any]],
    top_n: int,
) -> str:
    status_counts = Counter(str(row.get("status", "")) for row in item_rows)
    not_found_pharmacies = sum(1 for row in summary_rows if int(row.get("matched_items") or 0) == 0 and not row.get("error"))

    lines = [
        f"Wolt prices report",
        f"Brand: {brand}",
        f"Query: {query}",
        f"Run: {checked_at}",
        "",
        f"Pharmacies: {pharmacies_count}",
        f"Items found: {len(item_rows)}",
        f"In stock: {status_counts.get('in_stock', 0)}",
        f"Out of stock: {status_counts.get('out_of_stock', 0)}",
        f"Unavailable: {status_counts.get('unavailable', 0)}",
        f"No matches pharmacies: {not_found_pharmacies}",
        f"New item_id this run: {new_ids_count}",
        "",
        f"Top {min(top_n, len(top_offers))} best offers:",
    ]

    for idx, row in enumerate(top_offers[:top_n], start=1):
        lines.append(
            f"{idx}. {row.get('item_name')} | {row.get('price')} KZT | {row.get('pharmacy')}"
        )
        lines.append(f"   Product: {row.get('product_link')}")
        lines.append(f"   Venue: {row.get('venue_url')}")
    return "\n".join(lines)


def load_catalog(path: Path) -> dict[tuple[str, str], dict[str, Any]]:
    if not path.exists():
        return {}
    catalog: dict[tuple[str, str], dict[str, Any]] = {}
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            brand = (row.get("brand") or "").strip()
            item_id = (row.get("item_id") or "").strip()
            if not brand or not item_id:
                continue
            key = (normalize_text(brand), item_id)
            catalog[key] = {
                "brand": brand,
                "item_id": item_id,
                "item_name": (row.get("item_name") or "").strip(),
                "first_seen": (row.get("first_seen") or "").strip(),
                "last_seen": (row.get("last_seen") or "").strip(),
                "seen_runs": int((row.get("seen_runs") or "0").strip() or "0"),
                "last_seen_venue_slug": (row.get("last_seen_venue_slug") or "").strip(),
                "last_seen_pharmacy": (row.get("last_seen_pharmacy") or "").strip(),
                "canonical_sku": (row.get("canonical_sku") or "").strip(),
                "canonical_name": (row.get("canonical_name") or "").strip(),
                "canonical_line": (row.get("canonical_line") or "").strip(),
                "canonical_pack_size": (row.get("canonical_pack_size") or "").strip(),
                "canonical_form_factor": (row.get("canonical_form_factor") or "").strip(),
                "canonical_dosage_or_volume": (row.get("canonical_dosage_or_volume") or "").strip(),
                "canonical_flavor": (row.get("canonical_flavor") or "").strip(),
                "canonical_active_ingredient": (row.get("canonical_active_ingredient") or "").strip(),
                "canonical_confidence": (row.get("canonical_confidence") or "").strip(),
                "canonical_rule": (row.get("canonical_rule") or "").strip(),
            }
    return catalog


def save_catalog(path: Path, catalog: dict[tuple[str, str], dict[str, Any]]) -> None:
    rows = sorted(
        catalog.values(),
        key=lambda x: (normalize_text(str(x.get("brand", ""))), normalize_text(str(x.get("item_name", "")))),
    )
    write_csv(
        path,
        rows,
        [
            "brand",
            "item_id",
            "item_name",
            "canonical_sku",
            "canonical_name",
            "canonical_line",
            "canonical_pack_size",
            "canonical_form_factor",
            "canonical_dosage_or_volume",
            "canonical_flavor",
            "canonical_active_ingredient",
            "canonical_confidence",
            "canonical_rule",
            "first_seen",
            "last_seen",
            "seen_runs",
            "last_seen_venue_slug",
            "last_seen_pharmacy",
        ],
    )


def run() -> int:
    args = parse_args()
    load_env_from_file()
    query = args.query.strip() or args.brand.strip()
    brand = args.brand.strip()
    brand_norm = normalize_text(brand)
    is_vitrum_brand = brand_norm == "vitrum" or "витрум" in brand_norm
    strict_brand_filter = not is_vitrum_brand
    if not brand:
        raise ValueError("Brand must not be empty")

    pharmacies = read_pharmacies(Path(args.pharmacies_csv).expanduser().resolve(), city_slug=args.city_slug)
    session = build_session()
    checked_at = datetime.now().isoformat(timespec="seconds")

    item_rows: list[dict[str, Any]] = []
    summary_rows: list[dict[str, Any]] = []
    errors = 0

    for idx, ph in enumerate(pharmacies, start=1):
        slug = ph["slug"]
        name = ph["name"]
        url = ph["venue_url"]
        try:
            items = fetch_items(
                session=session,
                venue_slug=slug,
                query=query,
                language=args.language,
                timeout=args.timeout,
            )
            counts = {"in_stock": 0, "out_of_stock": 0, "unavailable": 0}
            raw_items_count = 0
            filtered_out_count = 0
            matched_items_count = 0

            for item in items:
                if not isinstance(item, dict):
                    continue
                raw_items_count += 1
                item_id = str(item.get("id") or "").strip()
                item_name = str(item.get("name") or "").strip()
                if not item_id or not item_name:
                    continue
                if strict_brand_filter and not item_matches_brand_name(brand, item_name):
                    filtered_out_count += 1
                    continue

                status, disable_text = item_status(item)
                if status in counts:
                    counts[status] += 1
                matched_items_count += 1

                price_minor = item.get("price")
                canonical = canonicalize_brand_name(brand, item_name)
                item_rows.append(
                    {
                        "checked_at": checked_at,
                        "brand": brand,
                        "query": query,
                        "pharmacy": name,
                        "venue_slug": slug,
                        "venue_url": url,
                        "item_id": item_id,
                        "item_name": item_name,
                        "status": status,
                        "disable_text": disable_text,
                        "price_minor": price_minor if price_minor is not None else "",
                        "price": minor_to_major(price_minor),
                        "purchasable_balance": item.get("purchasable_balance", ""),
                        "product_link": f"https://wolt.com/en/kaz/{args.city_slug}/venue/{slug}?menuItem={item_id}",
                        "search_link": f"https://wolt.com/en/kaz/{args.city_slug}/search?q={quote_plus(item_name)}",
                        "canonical_sku": canonical.get("canonical_sku", ""),
                        "canonical_name": canonical.get("canonical_name", ""),
                        "canonical_line": canonical.get("product_line", ""),
                        "canonical_pack_size": canonical.get("pack_size", ""),
                        "canonical_form_factor": canonical.get("form_factor", ""),
                        "canonical_dosage_or_volume": canonical.get("dosage_or_volume", ""),
                        "canonical_flavor": canonical.get("flavor", ""),
                        "canonical_active_ingredient": canonical.get("active_ingredient", ""),
                        "canonical_confidence": canonical.get("confidence", ""),
                        "canonical_rule": canonical.get("rule", ""),
                    }
                )

            summary_rows.append(
                {
                    "checked_at": checked_at,
                    "brand": brand,
                    "query": query,
                    "pharmacy": name,
                    "venue_slug": slug,
                    "venue_url": url,
                    "raw_items": raw_items_count,
                    "filtered_out_items": filtered_out_count,
                    "matched_items": matched_items_count,
                    "in_stock_items": counts["in_stock"],
                    "out_of_stock_items": counts["out_of_stock"],
                    "unavailable_items": counts["unavailable"],
                    "error": "",
                }
            )
        except Exception as exc:
            errors += 1
            summary_rows.append(
                {
                    "checked_at": checked_at,
                    "brand": brand,
                    "query": query,
                    "pharmacy": name,
                    "venue_slug": slug,
                    "venue_url": url,
                    "raw_items": 0,
                    "filtered_out_items": 0,
                    "matched_items": 0,
                    "in_stock_items": 0,
                    "out_of_stock_items": 0,
                    "unavailable_items": 0,
                    "error": str(exc),
                }
            )

        if args.sleep_ms > 0:
            time.sleep(args.sleep_ms / 1000.0)
        if idx % 15 == 0 or idx == len(pharmacies):
            print(f"[{idx}/{len(pharmacies)}] processed, item rows={len(item_rows)}, errors={errors}")

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    results_dir = Path(args.results_dir).expanduser().resolve()
    safe_brand = slugify_token(brand) or "brand"
    base = f"wolt_brand_{safe_brand}_{timestamp}"

    items_report_path = results_dir / f"{base}_items.csv"
    summary_report_path = results_dir / f"{base}_summary.csv"

    write_csv(
        items_report_path,
        item_rows,
        [
            "checked_at",
            "brand",
            "query",
            "pharmacy",
            "venue_slug",
            "venue_url",
            "item_id",
            "item_name",
            "status",
            "disable_text",
            "price_minor",
            "price",
            "purchasable_balance",
            "product_link",
            "search_link",
            "canonical_sku",
            "canonical_name",
            "canonical_line",
            "canonical_pack_size",
            "canonical_form_factor",
            "canonical_dosage_or_volume",
            "canonical_flavor",
            "canonical_active_ingredient",
            "canonical_confidence",
            "canonical_rule",
        ],
    )

    write_csv(
        summary_report_path,
        summary_rows,
        [
            "checked_at",
            "brand",
            "query",
            "pharmacy",
            "venue_slug",
            "venue_url",
            "raw_items",
            "filtered_out_items",
            "matched_items",
            "in_stock_items",
            "out_of_stock_items",
            "unavailable_items",
            "error",
        ],
    )

    # Update growing item_id catalog by brand
    catalog_path = Path(args.items_catalog).expanduser().resolve()
    catalog = load_catalog(catalog_path)
    observed_now: dict[tuple[str, str], dict[str, str]] = {}
    new_rows: list[dict[str, Any]] = []

    for row in item_rows:
        key = (normalize_text(str(row["brand"])), str(row["item_id"]))
        observed_now[key] = {
            "brand": str(row["brand"]),
            "item_id": str(row["item_id"]),
            "item_name": str(row["item_name"]),
            "venue_slug": str(row["venue_slug"]),
            "pharmacy": str(row["pharmacy"]),
            "canonical_sku": str(row.get("canonical_sku") or ""),
            "canonical_name": str(row.get("canonical_name") or ""),
            "canonical_line": str(row.get("canonical_line") or ""),
            "canonical_pack_size": str(row.get("canonical_pack_size") or ""),
            "canonical_form_factor": str(row.get("canonical_form_factor") or ""),
            "canonical_dosage_or_volume": str(row.get("canonical_dosage_or_volume") or ""),
            "canonical_flavor": str(row.get("canonical_flavor") or ""),
            "canonical_active_ingredient": str(row.get("canonical_active_ingredient") or ""),
            "canonical_confidence": str(row.get("canonical_confidence") or ""),
            "canonical_rule": str(row.get("canonical_rule") or ""),
        }

    for key, obs in observed_now.items():
        if key not in catalog:
            catalog[key] = {
                "brand": obs["brand"],
                "item_id": obs["item_id"],
                "item_name": obs["item_name"],
                "first_seen": checked_at,
                "last_seen": checked_at,
                "seen_runs": 1,
                "last_seen_venue_slug": obs["venue_slug"],
                "last_seen_pharmacy": obs["pharmacy"],
                "canonical_sku": obs["canonical_sku"],
                "canonical_name": obs["canonical_name"],
                "canonical_line": obs["canonical_line"],
                "canonical_pack_size": obs["canonical_pack_size"],
                "canonical_form_factor": obs["canonical_form_factor"],
                "canonical_dosage_or_volume": obs["canonical_dosage_or_volume"],
                "canonical_flavor": obs["canonical_flavor"],
                "canonical_active_ingredient": obs["canonical_active_ingredient"],
                "canonical_confidence": obs["canonical_confidence"],
                "canonical_rule": obs["canonical_rule"],
            }
            new_rows.append(
                {
                    "brand": obs["brand"],
                    "item_id": obs["item_id"],
                    "item_name": obs["item_name"],
                    "canonical_sku": obs["canonical_sku"],
                    "canonical_name": obs["canonical_name"],
                    "canonical_active_ingredient": obs["canonical_active_ingredient"],
                    "canonical_confidence": obs["canonical_confidence"],
                    "first_seen": checked_at,
                }
            )
        else:
            rec = catalog[key]
            rec["item_name"] = obs["item_name"] or rec.get("item_name", "")
            rec["last_seen"] = checked_at
            rec["seen_runs"] = int(rec.get("seen_runs", 0)) + 1
            rec["last_seen_venue_slug"] = obs["venue_slug"]
            rec["last_seen_pharmacy"] = obs["pharmacy"]
            rec["canonical_sku"] = obs["canonical_sku"] or rec.get("canonical_sku", "")
            rec["canonical_name"] = obs["canonical_name"] or rec.get("canonical_name", "")
            rec["canonical_line"] = obs["canonical_line"] or rec.get("canonical_line", "")
            rec["canonical_pack_size"] = obs["canonical_pack_size"] or rec.get("canonical_pack_size", "")
            rec["canonical_form_factor"] = obs["canonical_form_factor"] or rec.get("canonical_form_factor", "")
            rec["canonical_dosage_or_volume"] = obs["canonical_dosage_or_volume"] or rec.get("canonical_dosage_or_volume", "")
            rec["canonical_flavor"] = obs["canonical_flavor"] or rec.get("canonical_flavor", "")
            rec["canonical_active_ingredient"] = obs["canonical_active_ingredient"] or rec.get("canonical_active_ingredient", "")
            rec["canonical_confidence"] = obs["canonical_confidence"] or rec.get("canonical_confidence", "")
            rec["canonical_rule"] = obs["canonical_rule"] or rec.get("canonical_rule", "")

    for rec in catalog.values():
        rec_brand_raw = str(rec.get("brand") or "").strip()
        if normalize_text(rec_brand_raw) != brand_norm:
            continue
        rec_name = str(rec.get("item_name") or "")
        canonical = canonicalize_brand_name(rec_brand_raw or brand, rec_name)
        rec["canonical_sku"] = canonical.get("canonical_sku", "")
        rec["canonical_name"] = canonical.get("canonical_name", "")
        rec["canonical_line"] = canonical.get("product_line", "")
        rec["canonical_pack_size"] = canonical.get("pack_size", "")
        rec["canonical_form_factor"] = canonical.get("form_factor", "")
        rec["canonical_dosage_or_volume"] = canonical.get("dosage_or_volume", "")
        rec["canonical_flavor"] = canonical.get("flavor", "")
        rec["canonical_active_ingredient"] = canonical.get("active_ingredient", "")
        rec["canonical_confidence"] = canonical.get("confidence", "")
        rec["canonical_rule"] = canonical.get("rule", "")

    removed_noise = 0
    if strict_brand_filter:
        to_remove: list[tuple[str, str]] = []
        for key, rec in catalog.items():
            rec_brand_raw = str(rec.get("brand") or "").strip()
            if normalize_text(rec_brand_raw) != brand_norm:
                continue
            rec_name = str(rec.get("item_name") or "")
            if not item_matches_brand_name(rec_brand_raw or brand, rec_name):
                to_remove.append(key)
        for key in to_remove:
            catalog.pop(key, None)
        removed_noise = len(to_remove)

    save_catalog(catalog_path, catalog)
    active_reference_path = catalog_path.parent / DEFAULT_ACTIVE_INGREDIENT_REFERENCE.name
    export_active_ingredient_reference(active_reference_path)

    vitrum_export: dict[str, Any] | None = None
    brand_export: dict[str, Any] | None = None
    if is_vitrum_brand:
        vitrum_export = export_vitrum_reference(
            catalog=catalog,
            reference_path=Path(args.vitrum_reference_csv).expanduser().resolve(),
            canonical_path=Path(args.vitrum_canonical_csv).expanduser().resolve(),
            unmapped_path=Path(args.vitrum_unmapped_csv).expanduser().resolve(),
        )
    else:
        state_dir = catalog_path.parent
        brand_export = export_brand_reference(
            catalog=catalog,
            brand=brand,
            reference_path=state_dir / f"wolt_{safe_brand}_item_reference.csv",
            canonical_path=state_dir / f"wolt_{safe_brand}_canonical_catalog.csv",
            unmapped_path=state_dir / f"wolt_{safe_brand}_unmapped.csv",
        )

    new_items_path = results_dir / f"{base}_new_item_ids.csv"
    write_csv(
        new_items_path,
        new_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "canonical_sku",
            "canonical_name",
            "canonical_active_ingredient",
            "canonical_confidence",
            "first_seen",
        ],
    )

    top_offers = build_best_offers(item_rows)
    best_prices_path = results_dir / f"{base}_best_prices.csv"
    write_csv(
        best_prices_path,
        top_offers,
        [
            "checked_at",
            "brand",
            "query",
            "pharmacy",
            "venue_slug",
            "venue_url",
            "item_id",
            "item_name",
            "status",
            "price_minor",
            "price",
            "product_link",
            "search_link",
            "canonical_sku",
            "canonical_name",
            "canonical_line",
            "canonical_pack_size",
            "canonical_form_factor",
            "canonical_dosage_or_volume",
            "canonical_flavor",
            "canonical_active_ingredient",
            "canonical_confidence",
            "canonical_rule",
        ],
    )

    status_counts = Counter(str(row.get("status", "")) for row in item_rows)
    not_found_pharmacies = sum(1 for row in summary_rows if int(row.get("matched_items") or 0) == 0 and not row.get("error"))
    dashboard_path = results_dir / f"{base}_dashboard.png"
    dashboard_ok = create_visual_summary(
        out_path=dashboard_path,
        brand=brand,
        query=query,
        checked_at=checked_at,
        pharmacies_count=len(pharmacies),
        item_rows_count=len(item_rows),
        in_stock_count=status_counts.get("in_stock", 0),
        out_of_stock_count=status_counts.get("out_of_stock", 0),
        not_found_pharmacies=not_found_pharmacies,
        new_ids_count=len(new_rows),
        top_offers=top_offers,
    )

    telegram_text = build_telegram_text(
        brand=brand,
        query=query,
        checked_at=checked_at,
        pharmacies_count=len(pharmacies),
        item_rows=item_rows,
        summary_rows=summary_rows,
        new_ids_count=len(new_rows),
        top_offers=top_offers,
        top_n=max(1, args.telegram_top_n),
    )

    if args.send_telegram:
        token = args.telegram_bot_token.strip() or os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
        chat_id = args.telegram_chat_id.strip() or os.environ.get("TELEGRAM_CHAT_ID", "").strip()
        if not token or not chat_id:
            raise RuntimeError("Telegram send enabled, but TELEGRAM_BOT_TOKEN/TELEGRAM_CHAT_ID is missing")
        if dashboard_ok and dashboard_path.exists():
            send_telegram_photo(token, chat_id, dashboard_path, caption=f"Wolt {brand}: dashboard")
        send_telegram_message(token, chat_id, telegram_text)
        send_telegram_document(token, chat_id, best_prices_path, caption=f"Wolt {brand}: best prices + links")
        send_telegram_document(token, chat_id, items_report_path, caption=f"Wolt {brand}: full items report")
        send_telegram_document(token, chat_id, new_items_path, caption=f"Wolt {brand}: new item_ids ({len(new_rows)})")
        if vitrum_export:
            send_telegram_document(
                token,
                chat_id,
                Path(vitrum_export["canonical_path"]),
                caption=f"Wolt {brand}: canonical SKU catalog",
            )
        elif brand_export:
            send_telegram_document(
                token,
                chat_id,
                Path(brand_export["canonical_path"]),
                caption=f"Wolt {brand}: canonical SKU catalog",
            )

    print(f"Checked pharmacies: {len(pharmacies)}")
    print(f"Item rows: {len(item_rows)}")
    print(f"Pharmacy errors: {errors}")
    print(f"Items report: {items_report_path}")
    print(f"Summary report: {summary_report_path}")
    print(f"Best prices report: {best_prices_path}")
    print(f"Updated item catalog: {catalog_path}")
    print(f"Active ingredients reference: {active_reference_path}")
    if strict_brand_filter:
        print(f"Removed noisy non-brand IDs from catalog: {removed_noise}")
    print(f"New item IDs this run: {len(new_rows)}")
    print(f"New IDs file: {new_items_path}")
    if vitrum_export:
        print(
            "Vitrum reference: "
            f"{vitrum_export['reference_path']} "
            f"(rows={vitrum_export['detailed_count']})"
        )
        print(
            "Vitrum canonical catalog: "
            f"{vitrum_export['canonical_path']} "
            f"(sku={vitrum_export['canonical_count']}, unmapped={vitrum_export['unmapped_count']})"
        )
        print(f"Vitrum unmapped: {vitrum_export['unmapped_path']}")
    elif brand_export:
        print(
            "Brand reference: "
            f"{brand_export['reference_path']} "
            f"(rows={brand_export['detailed_count']})"
        )
        print(
            "Brand canonical catalog: "
            f"{brand_export['canonical_path']} "
            f"(sku={brand_export['canonical_count']}, unmapped={brand_export['unmapped_count']})"
        )
        print(f"Brand unmapped: {brand_export['unmapped_path']}")
    if dashboard_ok:
        print(f"Dashboard image: {dashboard_path}")
    if args.send_telegram:
        print("Telegram: sent")
    return 0 if errors == 0 else 1


def main() -> None:
    try:
        code = run()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)
    sys.exit(code)


if __name__ == "__main__":
    main()
