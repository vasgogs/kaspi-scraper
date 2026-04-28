#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import html
import difflib
import random
import re
import shutil
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote_plus, urljoin

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


CITY_CODE = "750000000"
WARMUP_URL = f"https://kaspi.kz/shop/p/l-tset-tabletki-5-mg-30-sht-101614081/?c={CITY_CODE}"
BASE_URL = "https://kaspi.kz"
SEARCH_URL = f"{BASE_URL}/shop/search/?text={{query}}"
DEFAULT_INPUT = Path("/home/vas/kaspi-scraper/list for Natasha.csv")
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
]
FORM_TOKENS = {
    "таб",
    "капс",
    "капли",
    "спрей",
    "гель",
    "крем",
    "мазь",
    "сусп",
    "раств",
    "супп",
    "сироп",
    "шамп",
    "паст",
    "жев",
    "гран",
    "ампулы",
    "шипучий",
}
LIQUID_FORM_TOKENS = {"капли", "спрей", "сусп", "раств", "сироп"}
UNIT_TOKENS = {"мг", "мл", "шт", "г", "кг", "л", "ед", "ме"}
DROP_TOKENS = {"stada"}
TOKEN_ALIASES = {
    "таблетки": "таб",
    "таблетка": "таб",
    "табл": "таб",
    "капсулы": "капс",
    "капсула": "капс",
    "капли": "капли",
    "капля": "капли",
    "спрей": "спрей",
    "аэрозоль": "спрей",
    "суспензия": "сусп",
    "пастилки": "паст",
    "пастилка": "паст",
    "мармеладки": "жев",
    "мармелад": "жев",
    "жевательные": "жев",
    "гаммисы": "жев",
    "гранулы": "гран",
    "шампунь": "шамп",
    "суппозитории": "супп",
    "свечи": "супп",
    "раствор": "раств",
    "шипучие": "шипучий",
    "беби": "бэби",
}
SPECIAL_FAMILY_ALIASES = {
    ("лазолван", "юниор"): ("лазолван", "детский"),
}
UNITS_RE = r"мг|мл|шт|г|кг|л|ед|ме"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Fill list for Natasha.csv with Kaspi names and links from Kaspi search."
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT),
        help=f"Input CSV path (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Only process the first N data rows.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Rewrite rows even if Kaspi_link is already filled.",
    )
    return parser.parse_args()


def tokenize(text: str) -> list[str]:
    value = html.unescape(str(text or "")).lower().replace("ё", "е")
    value = value.replace("®", " ").replace("™", " ").replace("№", " n ")
    value = value.replace("х", "x")
    value = re.sub(rf"(\d+(?:\.\d+)?)(?=({UNITS_RE})\b)", r"\1 ", value)
    value = re.sub(r"(\d)\s*x\s*(\d)", r"\1 x \2", value)
    value = value.replace("+", " + ")
    value = re.sub(r"(?<=\d),(?=\d)", ".", value)
    value = re.sub(r"[^0-9a-zа-я.+]+", " ", value)
    raw_tokens = [TOKEN_ALIASES.get(token, token) for token in value.split() if token]
    tokens: list[str] = []
    index = 0
    while index < len(raw_tokens):
        token = raw_tokens[index]
        next_token = raw_tokens[index + 1] if index + 1 < len(raw_tokens) else None
        prev_token = tokens[-1] if tokens else None
        if token in DROP_TOKENS:
            index += 1
            continue
        if next_token and len(token) == 1 and token.isalpha() and re.fullmatch(r"\d+(?:\.\d+)?", next_token):
            if token in {"в", "b", "д", "d"} or (token in {"с", "c"} and prev_token not in {"витамин"}):
                tokens.append(token + next_token)
                index += 2
                continue
        tokens.append(token)
        index += 1
    return tokens


def family(tokens: list[str]) -> tuple[str, ...]:
    result: list[str] = []
    for token in tokens:
        if token in DROP_TOKENS or token == "n":
            continue
        if not result and token in FORM_TOKENS:
            continue
        if token in FORM_TOKENS or token in UNIT_TOKENS or re.fullmatch(r"\d+(?:\.\d+)?", token):
            break
        if token in {"x", "+", "со", "вкусом", "для"} and result:
            break
        result.append(token)
        if len(result) >= 6:
            break
    family_key = tuple(result)
    return SPECIAL_FAMILY_ALIASES.get(family_key, family_key)


def primary_form(tokens: list[str]) -> str:
    for token in tokens:
        if token in FORM_TOKENS:
            return token
    return ""


def numeric_tokens(tokens: list[str]) -> set[str]:
    return {token for token in tokens if re.search(r"\d", token)}


def pack_counts(tokens: list[str]) -> set[str]:
    counts: set[str] = set()
    for index, token in enumerate(tokens):
        next_token = tokens[index + 1] if index + 1 < len(tokens) else None
        if token == "n" and next_token and re.fullmatch(r"\d+(?:\.\d+)?", next_token):
            counts.add(next_token)
        if re.fullmatch(r"\d+(?:\.\d+)?", token) and next_token == "шт":
            counts.add(token)
    return counts


def has_liquid_signal(tokens: list[str]) -> bool:
    return any(token in LIQUID_FORM_TOKENS for token in tokens) or "мл" in tokens


def family_variants(key: tuple[str, ...]) -> set[tuple[str, ...]]:
    variants = {key}
    if key and key[-1] == "st":
        variants.add(key[:-1])
    else:
        variants.add(key + ("st",))
    return variants


def is_ordered_subsequence(shorter: tuple[str, ...], longer: tuple[str, ...]) -> bool:
    if not shorter:
        return False
    position = 0
    for token in longer:
        if token == shorter[position]:
            position += 1
            if position == len(shorter):
                return True
    return False


def family_score(left: tuple[str, ...], right: tuple[str, ...]) -> float:
    if left == right:
        return 1.0
    if left in family_variants(right) or right in family_variants(left):
        return 0.95
    min_len = min(len(left), len(right))
    if min_len >= 2 and left[:min_len] == right[:min_len]:
        return 0.88
    shorter, longer = (left, right) if len(left) <= len(right) else (right, left)
    if len(shorter) >= 2 and is_ordered_subsequence(shorter, longer):
        return 0.9
    if len(shorter) == 1 and shorter[0] in set(longer):
        return 0.78
    if min_len >= 1 and left[0] == right[0]:
        return 0.2
    overlap = set(left) & set(right)
    if overlap:
        return 2 * len(overlap) / (len(set(left)) + len(set(right)))
    return 0.0


def form_score(left: str, right: str) -> float:
    if not left or not right:
        return 0.7
    return 1.0 if left == right else 0.0


def match_score(source_name: str, candidate_name: str) -> float:
    source_tokens = tokenize(source_name)
    candidate_tokens = tokenize(candidate_name)
    left_family = family(source_tokens)
    right_family = family(candidate_tokens)
    left_family_score = family_score(left_family, right_family)
    if left_family_score == 0:
        return -1.0
    left_numeric = numeric_tokens(source_tokens)
    right_numeric = numeric_tokens(candidate_tokens)
    if left_numeric and not (left_numeric & right_numeric):
        return -1.0
    source_pack_counts = pack_counts(source_tokens)
    candidate_pack_counts = pack_counts(candidate_tokens)
    if source_pack_counts:
        if candidate_pack_counts and not (source_pack_counts & candidate_pack_counts):
            return -1.0
        if (
            not candidate_pack_counts
            and not (source_pack_counts & right_numeric)
            and "мл" not in source_tokens
            and has_liquid_signal(candidate_tokens)
        ):
            return -1.0
    numeric_score = (
        2 * len(left_numeric & right_numeric) / (len(left_numeric) + len(right_numeric))
        if (left_numeric or right_numeric)
        else 1.0
    )
    similarity = difflib.SequenceMatcher(
        None, " ".join(source_tokens), " ".join(candidate_tokens)
    ).ratio()
    return (
        0.48 * left_family_score
        + 0.20 * numeric_score
        + 0.17 * form_score(primary_form(source_tokens), primary_form(candidate_tokens))
        + 0.15 * similarity
    )


def ensure_columns(fieldnames: list[str]) -> list[str]:
    for required in ("Kaspi_Name", "Kaspi_link"):
        if required not in fieldnames:
            fieldnames.append(required)
    return fieldnames


def backup_path(input_path: Path) -> Path:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    return input_path.with_name(f"{input_path.stem}.backup_{stamp}{input_path.suffix}")


def report_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_kaspi_report.csv")


def save_csv(path: Path, fieldnames: list[str], rows: list[dict]) -> None:
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    tmp_path.replace(path)


def save_report(path: Path, rows: list[dict]) -> None:
    fieldnames = [
        "row",
        "sku",
        "status",
        "candidate_rank",
        "search_name",
        "card_name",
        "search_score",
        "card_score",
        "link",
        "notes",
    ]
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    tmp_path.replace(path)


def extract_card_name(page) -> str:
    try:
        value = page.evaluate(
            "() => window.digitalData?.product?.name || "
            "window.BACKEND?.components?.item?.card?.title || "
            "window.BACKEND?.components?.item?.name || null"
        )
    except Exception:
        value = None
    if value:
        return str(value).strip()
    try:
        heading = page.locator("h1").first
        if heading.count():
            return heading.inner_text().strip()
    except Exception:
        pass
    return ""


def warm_context(page) -> None:
    page.goto(WARMUP_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(3500)


def search_candidates(page, sku: str, limit: int = 5) -> list[dict]:
    query = quote_plus(sku.strip())
    page.goto(SEARCH_URL.format(query=query), wait_until="domcontentloaded", timeout=60000)
    try:
        page.locator("[data-product-id]").first.wait_for(timeout=12000)
        page.wait_for_timeout(800)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(2500)
    cards = page.locator("[data-product-id]")
    count = cards.count()
    results: list[dict] = []
    for index in range(min(count, limit)):
        card = cards.nth(index)
        try:
            search_name = card.locator(".item-card__name").inner_text().strip()
        except Exception:
            continue
        link_locator = card.locator("a[href*='/shop/p/']")
        if not link_locator.count():
            continue
        href = link_locator.first.get_attribute("href")
        if not href:
            continue
        full_url = urljoin(BASE_URL, href)
        results.append(
            {
                "rank": index + 1,
                "search_name": search_name,
                "url": full_url,
                "search_score": match_score(sku, search_name),
            }
        )
    return results


def verify_product(page, sku: str, candidate: dict) -> dict:
    page.goto(candidate["url"], wait_until="domcontentloaded", timeout=60000)
    try:
        page.locator("h1").first.wait_for(timeout=12000)
        page.wait_for_timeout(800)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(2500)
    card_name = extract_card_name(page)
    return {
        **candidate,
        "resolved_url": page.url,
        "card_name": card_name,
        "card_score": match_score(sku, card_name or candidate["search_name"]),
    }


def verify_best_candidates(page, sku: str, candidates: list[dict], limit: int = 3) -> list[dict]:
    verified: list[dict] = []
    for candidate in candidates[:limit]:
        verified_candidate = verify_product(page, sku, candidate)
        verified.append(verified_candidate)
        if verified_candidate["card_score"] >= 0.94:
            break
    return verified


def process_csv(input_path: Path, limit: int = 0, overwrite: bool = False) -> None:
    with input_path.open(encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        fieldnames = ensure_columns(list(reader.fieldnames or []))
        rows = list(reader)

    backup = backup_path(input_path)
    shutil.copy2(input_path, backup)
    report_rows: list[dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(
            viewport={"width": 1600, "height": 900},
            locale="ru-RU",
            user_agent=random.choice(USER_AGENTS),
        )
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        )
        search_page = context.new_page()
        product_page = context.new_page()
        warm_context(product_page)

        total = len(rows) if limit <= 0 else min(limit, len(rows))
        for index, row in enumerate(rows[:total], start=2):
            sku = (row.get("СКЮ") or "").strip()
            if not sku:
                report_rows.append(
                    {
                        "row": index,
                        "sku": "",
                        "status": "empty_sku",
                        "candidate_rank": "",
                        "search_name": "",
                        "card_name": "",
                        "search_score": "",
                        "card_score": "",
                        "link": "",
                        "notes": "",
                    }
                )
                continue

            if not overwrite and (row.get("Kaspi_link") or "").strip():
                report_rows.append(
                    {
                        "row": index,
                        "sku": sku,
                        "status": "skipped_existing",
                        "candidate_rank": "",
                        "search_name": row.get("Kaspi_Name", ""),
                        "card_name": row.get("Kaspi_Name", ""),
                        "search_score": "",
                        "card_score": "",
                        "link": row.get("Kaspi_link", ""),
                        "notes": "",
                    }
                )
                continue

            row["Kaspi_Name"] = ""
            row["Kaspi_link"] = ""

            try:
                candidates = search_candidates(search_page, sku)
            except PlaywrightTimeoutError:
                warm_context(product_page)
                candidates = search_candidates(search_page, sku)

            if not candidates:
                report_rows.append(
                    {
                        "row": index,
                        "sku": sku,
                        "status": "no_search_results",
                        "candidate_rank": "",
                        "search_name": "",
                        "card_name": "",
                        "search_score": "",
                        "card_score": "",
                        "link": "",
                        "notes": "",
                    }
                )
                continue

            try:
                verified_candidates = verify_best_candidates(product_page, sku, candidates)
            except PlaywrightTimeoutError:
                warm_context(product_page)
                verified_candidates = verify_best_candidates(product_page, sku, candidates)

            if not verified_candidates:
                best = max(candidates, key=lambda item: item["search_score"])
                report_rows.append(
                    {
                        "row": index,
                        "sku": sku,
                        "status": "search_not_verified",
                        "candidate_rank": best["rank"],
                        "search_name": best["search_name"],
                        "card_name": "",
                        "search_score": f"{best['search_score']:.3f}",
                        "card_score": "",
                        "link": best["url"],
                        "notes": "Could not verify Kaspi product card",
                    }
                )
                continue

            verified = max(verified_candidates, key=lambda item: item["card_score"])

            if verified["card_score"] >= 0.80:
                row["Kaspi_Name"] = verified["card_name"] or verified["search_name"]
                row["Kaspi_link"] = verified["resolved_url"]
                status = "matched"
                notes = ""
            else:
                status = "low_card_confidence"
                notes = "No verified Kaspi card matched strongly enough"

            report_rows.append(
                {
                    "row": index,
                    "sku": sku,
                    "status": status,
                    "candidate_rank": verified["rank"],
                    "search_name": verified["search_name"],
                    "card_name": verified["card_name"],
                    "search_score": f"{verified['search_score']:.3f}",
                    "card_score": f"{verified['card_score']:.3f}",
                    "link": verified["resolved_url"],
                    "notes": notes,
                }
            )

            save_csv(input_path, fieldnames, rows)
            save_report(report_path(input_path), report_rows)
            print(f"[{index - 1}/{total}] {sku} -> {status}", flush=True)
            time.sleep(1.0)

        browser.close()

    save_csv(input_path, fieldnames, rows)
    save_report(report_path(input_path), report_rows)
    print(f"Backup: {backup}")
    print(f"Report: {report_path(input_path)}")


def main() -> int:
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 1
    process_csv(input_path, limit=args.limit, overwrite=args.overwrite)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
