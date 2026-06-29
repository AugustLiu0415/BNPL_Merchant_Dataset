#!/usr/bin/env python3
"""
Classify BNPL merchants by public listing status.

This script is intended as the second-stage pipeline after the BNPL merchant
industry classifier. It works at the deduplicated merchant/brand level and
adds listing/financial-data fields that can be used to build a public-company
research sample.

Default input:
    ~/Desktop/BNPL_Merchant_Dataset/Data_Clean/bnpl_merchant_overlap_summary_final_v7.xlsx
    sheet: Clean_Overlap

Default output:
    ~/Desktop/BNPL_Merchant_Dataset/Data_Clean/bnpl_merchant_public_listing_status.xlsx

Recommended first pass:
    python classify_merchants_public_listing_status.py --limit 100 --sleep 0.2

Fast offline smoke test:
    python classify_merchants_public_listing_status.py --offline --limit 25 --output /tmp/public_status_test.xlsx

Notes
-----
Automated public/private classification is inherently noisy because many
merchants are brands owned by larger public companies. The script therefore
separates:
    1. listing_status: the merchant/brand listing interpretation.
    2. financial_data_level: whether public financials are merchant-level,
       parent-level, or unavailable.
    3. needs_listing_review: rows that should be manually checked before use.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import sys
import time
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote_plus, urlencode, urlparse
from urllib.request import Request, urlopen

import pandas as pd


# ============================================================
# Paths
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent
DESKTOP_DATA_CLEAN = Path.home() / "Desktop" / "BNPL_Merchant_Dataset" / "Data_Clean"

DEFAULT_INPUT = DESKTOP_DATA_CLEAN / "bnpl_merchant_overlap_summary_final_v7.xlsx"
DEFAULT_SHEET = "Clean_Overlap"
DEFAULT_OUTPUT = DESKTOP_DATA_CLEAN / "bnpl_merchant_public_listing_status.xlsx"
DEFAULT_CACHE = DESKTOP_DATA_CLEAN / "public_listing_status_cache.jsonl"
DEFAULT_LOG = DESKTOP_DATA_CLEAN / "public_listing_status_run_log.csv"
DEFAULT_MANUAL_OVERRIDES = DESKTOP_DATA_CLEAN / "public_listing_manual_overrides.csv"
DEFAULT_REFERENCE_CACHE = DESKTOP_DATA_CLEAN / "public_reference_cache"


# ============================================================
# Network settings and public data sources
# ============================================================

HEADERS = {
    "User-Agent": os.environ.get(
        "BNPL_RESEARCH_USER_AGENT",
        (
            "August Liu academic BNPL merchant research "
            "(set BNPL_RESEARCH_USER_AGENT with contact email if needed)"
        ),
    ),
    "Accept": "application/json,text/plain,*/*",
}

REQUEST_TIMEOUT = 18

SEC_COMPANY_TICKERS_EXCHANGE_URL = "https://www.sec.gov/files/company_tickers_exchange.json"
YAHOO_SEARCH_URL = "https://query2.finance.yahoo.com/v1/finance/search"
WIKIDATA_SEARCH_URL = "https://www.wikidata.org/w/api.php"
WIKIDATA_ENTITY_URL = "https://www.wikidata.org/wiki/Special:EntityData/{qid}.json"


def http_get_json(url: str, params: dict[str, Any] | None = None) -> Any:
    full_url = url
    if params:
        full_url = f"{url}?{urlencode(params)}"
    request = Request(full_url, headers=HEADERS)
    try:
        with urlopen(request, timeout=REQUEST_TIMEOUT) as response:
            raw = response.read().decode("utf-8")
    except HTTPError as exc:
        raise RuntimeError(f"HTTP {exc.code} for {full_url}") from exc
    except URLError as exc:
        raise RuntimeError(f"Network error for {full_url}: {exc}") from exc
    return json.loads(raw)


# ============================================================
# Domain model
# ============================================================

LISTING_COLUMNS = [
    "listing_status",
    "financial_data_available",
    "financial_data_level",
    "public_company_name",
    "public_parent_name",
    "ticker",
    "exchange",
    "country_market",
    "cik",
    "matched_entity_name",
    "matched_entity_type",
    "listing_match_score",
    "listing_match_confidence",
    "listing_match_method",
    "listing_source",
    "listing_source_url",
    "listing_notes",
    "needs_listing_review",
]

REVIEW_STATUSES = {
    "ambiguous",
    "direct_public_candidate_needs_review",
    "subsidiary_public_candidate_needs_review",
    "acquired_or_delisted",
    "error",
}


@dataclass
class PublicRecord:
    name: str
    ticker: str = ""
    exchange: str = ""
    country_market: str = ""
    cik: str = ""
    source: str = ""
    source_url: str = ""
    aliases: list[str] = field(default_factory=list)
    entity_type: str = "company"


@dataclass
class CandidateMatch:
    record: PublicRecord
    score: float
    method: str
    notes: str = ""


@dataclass
class ListingResult:
    merchant_key: str
    merchant_name: str
    listing_status: str
    financial_data_available: str
    financial_data_level: str
    public_company_name: str = ""
    public_parent_name: str = ""
    ticker: str = ""
    exchange: str = ""
    country_market: str = ""
    cik: str = ""
    matched_entity_name: str = ""
    matched_entity_type: str = ""
    listing_match_score: float = 0.0
    listing_match_confidence: str = "low"
    listing_match_method: str = ""
    listing_source: str = ""
    listing_source_url: str = ""
    listing_notes: str = ""
    needs_listing_review: bool = True
    candidate_json: str = ""
    error_message: str = ""


# ============================================================
# Name normalization and matching
# ============================================================

LEGAL_SUFFIXES = {
    "inc",
    "incorporated",
    "corp",
    "corporation",
    "co",
    "company",
    "companies",
    "ltd",
    "limited",
    "llc",
    "com",
    "plc",
    "holdings",
    "holding",
    "sa",
    "s a",
    "ag",
    "nv",
    "n v",
    "se",
    "spa",
    "s p a",
    "ab",
    "oyj",
    "kk",
    "k k",
    "pte",
    "gmbh",
    "sarl",
    "bv",
    "b v",
    "asa",
    "as",
    "a s",
}

GENERIC_TOKENS = {
    "shop",
    "store",
    "official",
    "online",
    "usa",
    "us",
    "uk",
    "the",
    "and",
    "of",
    "for",
}

YAHOO_NON_OPERATING_TERMS = {
    " etf",
    " etn",
    " fund",
    " trust",
    " income shares",
    " enhanced high income",
    " cdr",
    " cad hedged",
    " hedged",
    " tracker",
    " trackx",
    " warrant",
    " warrants",
    " certificate",
    " certificates",
    " note",
    " notes",
    " bond",
    " bonds",
    " units",
    " unit",
    " option",
    " options",
    " bull ",
    " bear ",
}

PRIMARY_EXCHANGE_RANK = {
    "NASDAQ": 0,
    "NasdaqGS": 0,
    "NasdaqGM": 0,
    "NYSE": 0,
    "NYSEArca": 1,
    "XETRA": 1,
    "London": 1,
    "London Stock Exchange": 1,
    "Tokyo Stock Exchange": 1,
    "Toronto": 2,
    "Swiss": 2,
    "Frankfurt": 3,
    "Vienna": 4,
    "Munich": 4,
    "Stuttgart": 4,
    "Mexico": 4,
    "Buenos Aires": 4,
    "Sao Paulo": 4,
    "São Paulo": 4,
    "SET": 4,
}


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def normalize_name(value: str) -> str:
    text = clean_text(value).lower()
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return ""
    tokens = [t for t in text.split() if t not in LEGAL_SUFFIXES]
    while tokens and tokens[-1] in LEGAL_SUFFIXES:
        tokens.pop()
    return " ".join(tokens).strip()


def key_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", normalize_name(value))


def useful_tokens(value: str) -> set[str]:
    tokens = set(normalize_name(value).split())
    return {t for t in tokens if t not in LEGAL_SUFFIXES and t not in GENERIC_TOKENS and len(t) > 1}


def ratio(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return 100.0 * SequenceMatcher(None, a, b).ratio()


def token_sort_ratio(a: str, b: str) -> float:
    aa = " ".join(sorted(normalize_name(a).split()))
    bb = " ".join(sorted(normalize_name(b).split()))
    return ratio(aa, bb)


def token_set_ratio(a: str, b: str) -> float:
    ta = useful_tokens(a)
    tb = useful_tokens(b)
    if not ta or not tb:
        return ratio(normalize_name(a), normalize_name(b))
    common = ta & tb
    if not common:
        return 0.0
    left = " ".join(sorted(ta))
    right = " ".join(sorted(tb))
    overlap = len(common) / max(len(ta), len(tb))
    containment = len(common) / min(len(ta), len(tb))
    subset_score = 100.0 if ta == tb else 75.0 + 20.0 * overlap
    if containment < 1.0:
        subset_score = 100.0 * overlap
    return max(ratio(left, right), subset_score)


def name_match_score(merchant_name: str, company_name: str) -> float:
    merchant_norm = normalize_name(merchant_name)
    company_norm = normalize_name(company_name)
    if not merchant_norm or not company_norm:
        return 0.0

    scores = [
        ratio(merchant_norm, company_norm),
        token_sort_ratio(merchant_norm, company_norm),
        token_set_ratio(merchant_norm, company_norm),
    ]

    merchant_key = key_name(merchant_norm)
    company_key = key_name(company_norm)
    if merchant_key and merchant_key == company_key:
        scores.append(100.0)
    elif merchant_key and company_key:
        merchant_tokens = useful_tokens(merchant_norm)
        company_tokens = useful_tokens(company_norm)
        token_extra = max(0, len(company_tokens) - len(merchant_tokens))
        starts_or_ends = company_norm.startswith(merchant_norm) or company_norm.endswith(merchant_norm)
        if len(merchant_key) >= 5 and merchant_key in company_key:
            scores.append(94.0 if starts_or_ends or token_extra <= 2 else 84.0)
        if len(company_key) >= 5 and company_key in merchant_key:
            scores.append(92.0)

    return max(scores)


def search_friendly_name(value: str) -> str:
    text = clean_text(value)
    text = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", text)
    text = re.sub(r"(?<=[A-Z])(?=[A-Z][a-z])", " ", text)
    text = re.sub(r"([A-Za-z])([0-9])", r"\1 \2", text)
    text = re.sub(r"([0-9])([A-Za-z])", r"\1 \2", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or value


def get_domain(url: str) -> str:
    raw = clean_text(url)
    if not raw:
        return ""
    try:
        parsed = urlparse(raw if re.match(r"^https?://", raw) else f"https://{raw}")
    except ValueError:
        return ""
    host = parsed.netloc.lower()
    if host.startswith("www."):
        host = host[4:]
    return host


def domain_name_signal(merchant_name: str, official_url: str) -> float:
    domain = get_domain(official_url)
    if not domain:
        return 0.0
    domain_core = domain.split(".")[0]
    merchant_key = key_name(merchant_name)
    domain_key = key_name(domain_core)
    if not merchant_key or not domain_key:
        return 0.0
    if merchant_key == domain_key:
        return 8.0
    if len(merchant_key) >= 5 and merchant_key in domain_key:
        return 5.0
    if len(domain_key) >= 5 and domain_key in merchant_key:
        return 4.0
    return 0.0


def score_against_record(merchant_name: str, record: PublicRecord, official_url: str = "") -> float:
    names = [record.name, *record.aliases]
    score = max((name_match_score(merchant_name, n) for n in names if n), default=0.0)
    score += domain_name_signal(merchant_name, official_url)
    return min(score, 100.0)


def is_non_operating_security_name(name: str) -> bool:
    lowered = f" {clean_text(name).lower()} "
    return any(term in lowered for term in YAHOO_NON_OPERATING_TERMS)


def market_rank(record: PublicRecord) -> int:
    exchange = clean_text(record.exchange)
    if exchange in PRIMARY_EXCHANGE_RANK:
        return PRIMARY_EXCHANGE_RANK[exchange]
    if "." not in clean_text(record.ticker):
        return 2
    return 5


def sort_candidate_matches(candidates: list[CandidateMatch]) -> list[CandidateMatch]:
    return sorted(
        candidates,
        key=lambda c: (
            -c.score,
            market_rank(c.record),
            len(clean_text(c.record.ticker)),
            len(clean_text(c.record.name)),
        ),
    )


def high_confidence_direct_match(merchant_name: str, candidates: list[CandidateMatch]) -> bool:
    if not candidates:
        return False
    best = candidates[0]
    if best.score < 94:
        return False
    if best.score >= 99:
        return True
    if len(candidates) == 1:
        return True
    high_candidates = [c for c in candidates if c.score >= best.score - 1.0]
    distinct_names = {key_name(c.record.name) for c in high_candidates}
    if len(distinct_names) <= 1:
        return True
    best_key = key_name(best.record.name)
    merchant_key = key_name(merchant_name)
    return bool(merchant_key and (merchant_key == best_key or best_key.startswith(merchant_key)))


# ============================================================
# Input table handling
# ============================================================

def read_input_table(input_path: Path, sheet_name: str) -> pd.DataFrame:
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    xls = pd.ExcelFile(input_path)
    selected_sheet = sheet_name
    if selected_sheet not in xls.sheet_names:
        selected_sheet = xls.sheet_names[0]
        print(f"Sheet '{sheet_name}' not found; using first sheet '{selected_sheet}'.")
    df = pd.read_excel(input_path, sheet_name=selected_sheet)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def get_merchant_name(row: pd.Series) -> str:
    for col in [
        "canonical_company_name_clean",
        "canonical_company_name",
        "raw_company_name",
        "normalized_company_name",
        "clean_company_key",
    ]:
        value = clean_text(row.get(col, ""))
        if value:
            return value
    return ""


def get_merchant_key(row: pd.Series) -> str:
    for col in ["clean_company_key", "normalized_company_name", "canonical_company_name_clean"]:
        value = clean_text(row.get(col, ""))
        if value:
            return key_name(value)
    return key_name(get_merchant_name(row))


def get_official_url(row: pd.Series) -> str:
    for col in ["official_url", "final_url", "merchant_urls_clean", "merchant_url"]:
        value = clean_text(row.get(col, ""))
        if value:
            first_url = re.split(r"\s*[|;]\s*", value)[0]
            return first_url
    return ""


# ============================================================
# Manual overrides
# ============================================================

def load_manual_overrides(path: Path) -> dict[str, dict[str, str]]:
    overrides: dict[str, dict[str, str]] = {}
    if not path.exists():
        return overrides
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for raw in reader:
            row = {str(k).strip(): clean_text(v) for k, v in raw.items() if k is not None}
            status = row.get("listing_status", "")
            if not status:
                continue
            raw_key = row.get("merchant_key") or row.get("clean_company_key") or row.get("merchant_name")
            merchant_key = key_name(raw_key)
            if merchant_key:
                overrides[merchant_key] = row
    return overrides


def result_from_manual_override(
    row: pd.Series,
    merchant_key: str,
    merchant_name: str,
    override: dict[str, str],
) -> ListingResult:
    status = override.get("listing_status", "unknown")
    confidence = override.get("listing_match_confidence", "manual")
    needs_review = override.get("needs_listing_review", "").lower() in {"1", "true", "yes", "y"}
    if not override.get("needs_listing_review"):
        needs_review = status in REVIEW_STATUSES
    financial_data_available = override.get("financial_data_available", "")
    financial_data_level = override.get("financial_data_level", "")
    if not financial_data_available:
        if status == "direct_public":
            financial_data_available = "yes"
        elif status == "subsidiary_of_public":
            financial_data_available = "partial"
        elif status in {"private_no_public_match", "no_public_match"}:
            financial_data_available = "no"
        else:
            financial_data_available = "unknown"
    if not financial_data_level:
        if status == "direct_public":
            financial_data_level = "merchant_level"
        elif status == "subsidiary_of_public":
            financial_data_level = "parent_level"
        elif status in {"private_no_public_match", "no_public_match"}:
            financial_data_level = "none_public"
        else:
            financial_data_level = "unknown"

    source_url = (
        override.get("listing_source_url", "")
        or override.get("source_url", "")
        or override.get("url", "")
    )
    matched_entity_name = (
        override.get("matched_entity_name", "")
        or override.get("public_company_name", "")
        or override.get("public_parent_name", "")
    )

    return ListingResult(
        merchant_key=merchant_key,
        merchant_name=merchant_name,
        listing_status=status,
        financial_data_available=financial_data_available,
        financial_data_level=financial_data_level,
        public_company_name=override.get("public_company_name", ""),
        public_parent_name=override.get("public_parent_name", ""),
        ticker=override.get("ticker", ""),
        exchange=override.get("exchange", ""),
        country_market=override.get("country_market", ""),
        cik=override.get("cik", ""),
        matched_entity_name=matched_entity_name,
        matched_entity_type=override.get("matched_entity_type", "manual"),
        listing_match_score=float_or_zero(override.get("listing_match_score", "")),
        listing_match_confidence=confidence,
        listing_match_method="manual_override",
        listing_source=override.get("listing_source", "manual_override"),
        listing_source_url=source_url,
        listing_notes=override.get("listing_notes", override.get("notes", "")),
        needs_listing_review=needs_review,
    )


def apply_manual_overrides_to_results(
    df: pd.DataFrame,
    results: dict[str, ListingResult],
    manual_overrides: dict[str, dict[str, str]],
) -> int:
    """Apply manual overrides on top of cached or newly generated results."""
    applied = 0
    if not manual_overrides:
        return applied
    for _, row in df.iterrows():
        merchant_key = get_merchant_key(row)
        override = manual_overrides.get(merchant_key)
        if not override:
            continue
        merchant_name = get_merchant_name(row)
        results[merchant_key] = result_from_manual_override(row, merchant_key, merchant_name, override)
        applied += 1
    return applied


def build_manual_override_template(df: pd.DataFrame, max_rows: int = 500) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for _, row in df.head(max_rows).iterrows():
        rows.append(
            {
                "merchant_key": get_merchant_key(row),
                "merchant_name": get_merchant_name(row),
                "listing_status": "",
                "financial_data_available": "",
                "financial_data_level": "",
                "public_company_name": "",
                "public_parent_name": "",
                "ticker": "",
                "exchange": "",
                "country_market": "",
                "cik": "",
                "listing_match_confidence": "",
                "needs_listing_review": "",
                "source_url": "",
                "notes": "",
            }
        )
    return pd.DataFrame(rows)


def float_or_zero(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


# ============================================================
# Reference data: SEC
# ============================================================

class PublicReferenceIndex:
    def __init__(self, records: list[PublicRecord]) -> None:
        self.records = records
        self.exact: dict[str, list[PublicRecord]] = defaultdict(list)
        self.token_index: dict[str, list[PublicRecord]] = defaultdict(list)
        self.ticker_index: dict[str, list[PublicRecord]] = defaultdict(list)
        self._build()

    def _build(self) -> None:
        for record in self.records:
            names = [record.name, *record.aliases]
            seen_names: set[str] = set()
            for name in names:
                norm_key = key_name(name)
                if norm_key and norm_key not in seen_names:
                    self.exact[norm_key].append(record)
                    seen_names.add(norm_key)
                for token in useful_tokens(name):
                    self.token_index[token].append(record)
            if record.ticker:
                self.ticker_index[record.ticker.upper()].append(record)

    def find_candidates(
        self,
        merchant_name: str,
        official_url: str = "",
        max_candidates: int = 6,
    ) -> list[CandidateMatch]:
        merchant_key = key_name(merchant_name)
        candidates: dict[tuple[str, str], PublicRecord] = {}

        for record in self.exact.get(merchant_key, []):
            candidates[(record.source, record.ticker or record.name)] = record

        upper_name = clean_text(merchant_name).upper()
        if 1 <= len(upper_name) <= 8:
            for record in self.ticker_index.get(upper_name, []):
                candidates[(record.source, record.ticker or record.name)] = record

        token_hits: dict[tuple[str, str], PublicRecord] = {}
        for token in useful_tokens(merchant_name):
            for record in self.token_index.get(token, [])[:500]:
                token_hits[(record.source, record.ticker or record.name)] = record
        candidates.update(token_hits)

        matches: list[CandidateMatch] = []
        for record in candidates.values():
            score = score_against_record(merchant_name, record, official_url)
            if score >= 70:
                method = "sec_exact_or_token_fuzzy"
                if key_name(record.name) == merchant_key:
                    method = "sec_exact_name"
                elif record.ticker and record.ticker.upper() == upper_name:
                    method = "sec_ticker_exact"
                matches.append(CandidateMatch(record=record, score=score, method=method))

        return sort_candidate_matches(matches)[:max_candidates]


def get_json_with_cache(
    url: str,
    cache_path: Path,
    offline: bool,
    sleep: float = 0.0,
    max_age_days: int = 30,
) -> Any:
    if cache_path.exists():
        age_seconds = time.time() - cache_path.stat().st_mtime
        if offline or age_seconds < max_age_days * 86400:
            with cache_path.open("r", encoding="utf-8") as f:
                return json.load(f)

    if offline:
        raise FileNotFoundError(f"Offline mode and cache not found: {cache_path}")

    if sleep:
        time.sleep(sleep)
    data = http_get_json(url)
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with cache_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return data


def load_sec_records(reference_cache: Path, offline: bool, sleep: float) -> list[PublicRecord]:
    try:
        data = get_json_with_cache(
            SEC_COMPANY_TICKERS_EXCHANGE_URL,
            reference_cache / "sec_company_tickers_exchange.json",
            offline=offline,
            sleep=sleep,
            max_age_days=30,
        )
    except Exception as exc:
        print(f"Warning: SEC reference data unavailable: {exc}", file=sys.stderr)
        return []

    fields = data.get("fields", [])
    rows = data.get("data", [])
    records: list[PublicRecord] = []
    for item in rows:
        row = dict(zip(fields, item))
        cik = str(row.get("cik", "")).zfill(10) if row.get("cik", "") != "" else ""
        name = clean_text(row.get("name", ""))
        ticker = clean_text(row.get("ticker", ""))
        exchange = clean_text(row.get("exchange", ""))
        if not name or not ticker:
            continue
        records.append(
            PublicRecord(
                name=name,
                ticker=ticker,
                exchange=exchange,
                country_market="US",
                cik=cik,
                source="SEC company_tickers_exchange",
                source_url=SEC_COMPANY_TICKERS_EXCHANGE_URL,
                aliases=[re.sub(r"\s+-\s+.*$", "", name).strip()],
            )
        )
    return records


# ============================================================
# Yahoo Finance per-merchant search
# ============================================================

def yahoo_search_records(
    merchant_name: str,
    reference_cache: Path,
    offline: bool,
    sleep: float,
    max_results: int = 8,
) -> list[PublicRecord]:
    query_options = [
        merchant_name,
        search_friendly_name(merchant_name),
        normalize_name(merchant_name),
    ]
    seen_queries: set[str] = set()
    records: list[PublicRecord] = []

    for query in query_options:
        query = clean_text(query)
        query_key = key_name(query)
        if not query or query_key in seen_queries:
            continue
        seen_queries.add(query_key)

        cache_key = f"{key_name(merchant_name) or 'unknown'}__{query_key or 'query'}"
        cache_path = reference_cache / "yahoo_search" / f"{cache_key}.json"
        query_url = f"{YAHOO_SEARCH_URL}?q={quote_plus(query)}&quotesCount={max_results}&newsCount=0"

        try:
            data = get_json_with_cache(query_url, cache_path, offline=offline, sleep=sleep, max_age_days=14)
        except Exception as exc:
            print(f"Warning: Yahoo search unavailable for {merchant_name} / {query}: {exc}", file=sys.stderr)
            continue

        for quote in data.get("quotes", []):
            quote_type = clean_text(quote.get("quoteType", quote.get("typeDisp", ""))).upper()
            if quote_type and quote_type != "EQUITY":
                continue
            symbol = clean_text(quote.get("symbol", ""))
            if not symbol:
                continue
            name = clean_text(quote.get("longname", "")) or clean_text(quote.get("shortname", ""))
            if not name:
                continue
            if is_non_operating_security_name(name):
                continue
            exchange = clean_text(quote.get("exchDisp", "")) or clean_text(quote.get("exchange", ""))
            records.append(
                PublicRecord(
                    name=name,
                    ticker=symbol,
                    exchange=exchange,
                    country_market=clean_text(quote.get("market", "")),
                    source="Yahoo Finance search",
                    source_url=f"https://finance.yahoo.com/quote/{symbol}",
                    aliases=[clean_text(quote.get("shortname", "")), clean_text(quote.get("longname", ""))],
                )
            )

    deduped: dict[tuple[str, str], PublicRecord] = {}
    for record in records:
        deduped[(record.ticker, key_name(record.name))] = record
    return list(deduped.values())


# ============================================================
# Wikidata public parent lookup
# ============================================================

def wikidata_api_get(url: str, params: dict[str, Any], offline: bool, sleep: float) -> Any:
    if offline:
        raise RuntimeError("Wikidata lookup skipped in offline mode.")
    if sleep:
        time.sleep(sleep)
    return http_get_json(url, params=params)


def wikidata_search_entities(
    merchant_name: str,
    reference_cache: Path,
    offline: bool,
    sleep: float,
    limit: int = 3,
) -> list[dict[str, Any]]:
    cache_path = reference_cache / "wikidata_search" / f"{key_name(merchant_name) or 'unknown'}.json"
    if cache_path.exists():
        with cache_path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        try:
            data = wikidata_api_get(
                WIKIDATA_SEARCH_URL,
                {
                    "action": "wbsearchentities",
                    "search": merchant_name,
                    "language": "en",
                    "format": "json",
                    "limit": limit,
                },
                offline=offline,
                sleep=sleep,
            )
        except Exception as exc:
            print(f"Warning: Wikidata search unavailable for {merchant_name}: {exc}", file=sys.stderr)
            return []
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        with cache_path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    return list(data.get("search", []))[:limit]


def wikidata_get_entity(qid: str, reference_cache: Path, offline: bool, sleep: float) -> dict[str, Any]:
    cache_path = reference_cache / "wikidata_entities" / f"{qid}.json"
    if cache_path.exists():
        with cache_path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        url = WIKIDATA_ENTITY_URL.format(qid=qid)
        try:
            data = get_json_with_cache(url, cache_path, offline=offline, sleep=sleep, max_age_days=60)
        except Exception as exc:
            print(f"Warning: Wikidata entity unavailable for {qid}: {exc}", file=sys.stderr)
            return {}
    return data.get("entities", {}).get(qid, {})


def wikidata_label(entity: dict[str, Any]) -> str:
    labels = entity.get("labels", {})
    for lang in ["en", "mul"]:
        if lang in labels:
            return clean_text(labels[lang].get("value", ""))
    if labels:
        first = next(iter(labels.values()))
        return clean_text(first.get("value", ""))
    return ""


def wikidata_aliases(entity: dict[str, Any]) -> list[str]:
    aliases = []
    for lang in ["en", "mul"]:
        for alias in entity.get("aliases", {}).get(lang, []):
            value = clean_text(alias.get("value", ""))
            if value:
                aliases.append(value)
    return aliases


def wikidata_claim_values(entity: dict[str, Any], prop: str) -> list[Any]:
    values: list[Any] = []
    for claim in entity.get("claims", {}).get(prop, []):
        mainsnak = claim.get("mainsnak", {})
        datavalue = mainsnak.get("datavalue", {})
        value = datavalue.get("value")
        if value is not None:
            values.append(value)
    return values


def wikidata_qids(entity: dict[str, Any], prop: str) -> list[str]:
    qids: list[str] = []
    for value in wikidata_claim_values(entity, prop):
        if isinstance(value, dict) and "id" in value:
            qids.append(clean_text(value["id"]))
    return qids


def wikidata_strings(entity: dict[str, Any], prop: str) -> list[str]:
    values: list[str] = []
    for value in wikidata_claim_values(entity, prop):
        if isinstance(value, str):
            values.append(clean_text(value))
    return values


def wikidata_urls(entity: dict[str, Any], prop: str) -> list[str]:
    return wikidata_strings(entity, prop)


def wikidata_entity_to_public_record(
    qid: str,
    entity: dict[str, Any],
    reference_cache: Path,
    offline: bool,
    sleep: float,
    entity_type: str = "company",
) -> PublicRecord | None:
    label = wikidata_label(entity)
    tickers = wikidata_strings(entity, "P249")
    if not tickers:
        return None

    exchange_qids = wikidata_qids(entity, "P414")
    exchange_labels: list[str] = []
    for exchange_qid in exchange_qids[:3]:
        exchange_entity = wikidata_get_entity(exchange_qid, reference_cache, offline=offline, sleep=sleep)
        exchange_label = wikidata_label(exchange_entity)
        if exchange_label:
            exchange_labels.append(exchange_label)

    ticker = tickers[0]
    return PublicRecord(
        name=label,
        ticker=ticker,
        exchange="; ".join(exchange_labels),
        country_market="",
        source="Wikidata",
        source_url=f"https://www.wikidata.org/wiki/{qid}",
        aliases=wikidata_aliases(entity),
        entity_type=entity_type,
    )


def wikidata_direct_or_parent_candidates(
    merchant_name: str,
    official_url: str,
    reference_cache: Path,
    offline: bool,
    sleep: float,
    max_entities: int = 3,
) -> list[CandidateMatch]:
    matches: list[CandidateMatch] = []
    search_hits = wikidata_search_entities(
        merchant_name,
        reference_cache=reference_cache,
        offline=offline,
        sleep=sleep,
        limit=max_entities,
    )
    merchant_domain = get_domain(official_url)

    for hit in search_hits:
        qid = clean_text(hit.get("id", ""))
        if not qid:
            continue
        entity = wikidata_get_entity(qid, reference_cache, offline=offline, sleep=sleep)
        if not entity:
            continue
        label = wikidata_label(entity)
        aliases = wikidata_aliases(entity)
        entity_score = max([name_match_score(merchant_name, label), *[name_match_score(merchant_name, a) for a in aliases]])
        official_domains = {get_domain(u) for u in wikidata_urls(entity, "P856")}
        official_domains = {d for d in official_domains if d}
        if merchant_domain and merchant_domain in official_domains:
            entity_score = min(100.0, entity_score + 10.0)

        direct_record = wikidata_entity_to_public_record(
            qid,
            entity,
            reference_cache=reference_cache,
            offline=offline,
            sleep=sleep,
            entity_type="wikidata_entity",
        )
        if direct_record and entity_score >= 75:
            matches.append(
                CandidateMatch(
                    record=direct_record,
                    score=entity_score,
                    method="wikidata_direct_ticker",
                    notes=f"entity={label}; qid={qid}",
                )
            )

        parent_qids = []
        for prop in ["P749", "P127"]:
            parent_qids.extend(wikidata_qids(entity, prop))
        seen_parent_qids: set[str] = set()
        for parent_qid in parent_qids[:4]:
            if parent_qid in seen_parent_qids:
                continue
            seen_parent_qids.add(parent_qid)
            parent_entity = wikidata_get_entity(parent_qid, reference_cache, offline=offline, sleep=sleep)
            if not parent_entity:
                continue
            parent_record = wikidata_entity_to_public_record(
                parent_qid,
                parent_entity,
                reference_cache=reference_cache,
                offline=offline,
                sleep=sleep,
                entity_type="public_parent",
            )
            if parent_record and entity_score >= 70:
                matches.append(
                    CandidateMatch(
                        record=parent_record,
                        score=entity_score,
                        method="wikidata_public_parent",
                        notes=f"merchant_entity={label}; merchant_qid={qid}; parent_qid={parent_qid}",
                    )
                )

    matches.sort(key=lambda m: m.score, reverse=True)
    return matches


# ============================================================
# Classification logic
# ============================================================

def confidence_from_score(score: float, method: str) -> str:
    if method == "manual_override":
        return "manual"
    if score >= 96:
        return "high"
    if score >= 88:
        return "medium"
    return "low"


def candidate_to_dict(candidate: CandidateMatch) -> dict[str, Any]:
    record = candidate.record
    return {
        "name": record.name,
        "ticker": record.ticker,
        "exchange": record.exchange,
        "country_market": record.country_market,
        "cik": record.cik,
        "source": record.source,
        "source_url": record.source_url,
        "entity_type": record.entity_type,
        "score": round(candidate.score, 2),
        "method": candidate.method,
        "notes": candidate.notes,
    }


def make_result_from_candidate(
    merchant_key: str,
    merchant_name: str,
    candidate: CandidateMatch,
    listing_status: str,
    needs_review: bool,
    notes: str = "",
) -> ListingResult:
    record = candidate.record
    is_parent = listing_status.startswith("subsidiary")
    confidence = confidence_from_score(candidate.score, candidate.method)
    public_company_name = "" if is_parent else record.name
    public_parent_name = record.name if is_parent else ""

    return ListingResult(
        merchant_key=merchant_key,
        merchant_name=merchant_name,
        listing_status=listing_status,
        financial_data_available="partial" if is_parent else "yes",
        financial_data_level="parent_level" if is_parent else "merchant_level",
        public_company_name=public_company_name,
        public_parent_name=public_parent_name,
        ticker=record.ticker,
        exchange=record.exchange,
        country_market=record.country_market,
        cik=record.cik,
        matched_entity_name=record.name,
        matched_entity_type=record.entity_type,
        listing_match_score=round(candidate.score, 2),
        listing_match_confidence=confidence,
        listing_match_method=candidate.method,
        listing_source=record.source,
        listing_source_url=record.source_url,
        listing_notes=notes or candidate.notes,
        needs_listing_review=needs_review,
    )


def classify_one_merchant(
    row: pd.Series,
    sec_index: PublicReferenceIndex,
    manual_overrides: dict[str, dict[str, str]],
    reference_cache: Path,
    offline: bool,
    sleep: float,
    use_yahoo: bool,
    use_wikidata: bool,
    classify_no_match_as_private: bool,
) -> ListingResult:
    merchant_key = get_merchant_key(row)
    merchant_name = get_merchant_name(row)
    official_url = get_official_url(row)

    if not merchant_name:
        return ListingResult(
            merchant_key=merchant_key,
            merchant_name="",
            listing_status="error",
            financial_data_available="unknown",
            financial_data_level="unknown",
            listing_notes="Missing merchant name.",
            needs_listing_review=True,
            error_message="Missing merchant name.",
        )

    override = manual_overrides.get(merchant_key)
    if override:
        return result_from_manual_override(row, merchant_key, merchant_name, override)

    all_candidates: list[CandidateMatch] = []

    sec_candidates = sec_index.find_candidates(merchant_name, official_url=official_url)
    all_candidates.extend(sec_candidates)

    if use_yahoo and not offline:
        yahoo_records = yahoo_search_records(
            merchant_name,
            reference_cache=reference_cache,
            offline=offline,
            sleep=sleep,
        )
        yahoo_candidates = [
            CandidateMatch(
                record=r,
                score=score_against_record(merchant_name, r, official_url),
                method="yahoo_finance_search",
            )
            for r in yahoo_records
        ]
        all_candidates.extend([c for c in yahoo_candidates if c.score >= 70])

    direct_candidates = [c for c in all_candidates if c.record.entity_type != "public_parent"]
    direct_candidates = sort_candidate_matches(direct_candidates)

    if direct_candidates:
        best = direct_candidates[0]
        if high_confidence_direct_match(merchant_name, direct_candidates):
            result = make_result_from_candidate(
                merchant_key,
                merchant_name,
                best,
                listing_status="direct_public",
                needs_review=False if best.score >= 96 else True,
                notes="Direct public company match. Use merchant-level public financials.",
            )
            result.candidate_json = json.dumps([candidate_to_dict(c) for c in direct_candidates[:5]], ensure_ascii=False)
            return result
        if best.score >= 86:
            result = make_result_from_candidate(
                merchant_key,
                merchant_name,
                best,
                listing_status="direct_public_candidate_needs_review",
                needs_review=True,
                notes="Possible direct public company match; manually verify before using in regressions.",
            )
            result.candidate_json = json.dumps([candidate_to_dict(c) for c in direct_candidates[:5]], ensure_ascii=False)
            return result

    wikidata_candidates: list[CandidateMatch] = []
    if use_wikidata and not offline:
        wikidata_candidates = wikidata_direct_or_parent_candidates(
            merchant_name,
            official_url=official_url,
            reference_cache=reference_cache,
            offline=offline,
            sleep=sleep,
        )
        all_candidates.extend(wikidata_candidates)

        direct_wd = [c for c in wikidata_candidates if c.method == "wikidata_direct_ticker"]
        direct_wd = sort_candidate_matches(direct_wd)
        if direct_wd and direct_wd[0].score >= 88:
            best = direct_wd[0]
            result = make_result_from_candidate(
                merchant_key,
                merchant_name,
                best,
                listing_status="direct_public" if best.score >= 94 else "direct_public_candidate_needs_review",
                needs_review=best.score < 94,
                notes="Wikidata reports a ticker for the merchant entity.",
            )
            result.candidate_json = json.dumps([candidate_to_dict(c) for c in all_candidates[:8]], ensure_ascii=False)
            return result

        parent_candidates = [c for c in wikidata_candidates if c.method == "wikidata_public_parent"]
        parent_candidates = sort_candidate_matches(parent_candidates)
        if parent_candidates and parent_candidates[0].score >= 82:
            best = parent_candidates[0]
            status = "subsidiary_of_public"
            review = best.score < 90
            if review:
                status = "subsidiary_public_candidate_needs_review"
            result = make_result_from_candidate(
                merchant_key,
                merchant_name,
                best,
                listing_status=status,
                needs_review=review,
                notes="Merchant/brand appears linked to a public parent. Use parent-level public financials only.",
            )
            result.candidate_json = json.dumps([candidate_to_dict(c) for c in all_candidates[:8]], ensure_ascii=False)
            return result

    all_candidates = sort_candidate_matches(all_candidates)
    if all_candidates and all_candidates[0].score >= 75:
        best = all_candidates[0]
        result = make_result_from_candidate(
            merchant_key,
            merchant_name,
            best,
            listing_status="ambiguous",
            needs_review=True,
            notes="Weak or conflicting public-company evidence. Manual review required.",
        )
        result.financial_data_available = "unknown"
        result.financial_data_level = "unknown"
        result.candidate_json = json.dumps([candidate_to_dict(c) for c in all_candidates[:8]], ensure_ascii=False)
        return result

    status = "private_no_public_match" if classify_no_match_as_private else "no_public_match"
    notes = (
        "No public listing match found from enabled sources. Treat as private only after manual confirmation."
        if not classify_no_match_as_private
        else "No public listing match found from enabled sources; classified as private by user option."
    )
    return ListingResult(
        merchant_key=merchant_key,
        merchant_name=merchant_name,
        listing_status=status,
        financial_data_available="no",
        financial_data_level="none_public",
        listing_match_score=0.0,
        listing_match_confidence="low",
        listing_match_method="no_match",
        listing_source="enabled_public_sources",
        listing_notes=notes,
        needs_listing_review=False if classify_no_match_as_private else True,
        candidate_json=json.dumps([candidate_to_dict(c) for c in all_candidates[:8]], ensure_ascii=False),
    )


# ============================================================
# Cache/log/output
# ============================================================

def load_result_cache(path: Path) -> dict[str, ListingResult]:
    cache: dict[str, ListingResult] = {}
    if not path.exists():
        return cache
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            raw = line.strip()
            if not raw:
                continue
            try:
                data = json.loads(raw)
                result = ListingResult(**data)
                cache[result.merchant_key] = result
            except Exception:
                continue
    return cache


def append_result_cache(path: Path, result: ListingResult) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(asdict(result), ensure_ascii=False) + "\n")


def write_log_row(path: Path, result: ListingResult) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    row = asdict(result)
    fieldnames = [
        "merchant_key",
        "merchant_name",
        "listing_status",
        "ticker",
        "exchange",
        "public_company_name",
        "public_parent_name",
        "listing_match_score",
        "listing_match_confidence",
        "listing_match_method",
        "needs_listing_review",
        "error_message",
    ]
    file_exists = path.exists()
    with path.open("a", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerow({field: row.get(field, "") for field in fieldnames})


def apply_results(df: pd.DataFrame, results: dict[str, ListingResult]) -> pd.DataFrame:
    out = df.copy()
    for col in LISTING_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    if "candidate_json" not in out.columns:
        out["candidate_json"] = ""

    for idx, row in out.iterrows():
        key = get_merchant_key(row)
        result = results.get(key)
        if not result:
            continue
        data = asdict(result)
        for col in LISTING_COLUMNS:
            value = data.get(col, "")
            if col == "needs_listing_review":
                value = bool(value)
            out.at[idx, col] = value
        out.at[idx, "candidate_json"] = data.get("candidate_json", "")
    return out


def build_summary(out: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    def add(metric: str, value: Any) -> None:
        rows.append({"metric": metric, "value": value})

    add("total_rows", len(out))
    for col in [
        "listing_status",
        "financial_data_available",
        "financial_data_level",
        "listing_match_confidence",
        "needs_listing_review",
        "verified_broad_industry",
        "provider_count_clean",
    ]:
        if col in out.columns:
            counts = out[col].fillna("").astype(str).value_counts(dropna=False)
            for key, value in counts.items():
                add(f"{col}::{key}", int(value))
    return pd.DataFrame(rows)


def save_output_workbook(df: pd.DataFrame, results: dict[str, ListingResult], output_path: Path) -> None:
    out = apply_results(df, results)
    summary = build_summary(out)
    needs_review = out[out["needs_listing_review"].astype(str).str.lower().isin({"true", "1"})].copy()
    public_sample = out[
        out["listing_status"].isin(["direct_public", "subsidiary_of_public"])
        & ~out["needs_listing_review"].astype(str).str.lower().isin({"true", "1"})
    ].copy()
    direct_public = out[out["listing_status"].eq("direct_public")].copy()
    parent_public = out[out["listing_status"].eq("subsidiary_of_public")].copy()
    manual_overrides_applied = out[out["listing_match_method"].eq("manual_override")].copy()
    override_template = build_manual_override_template(needs_review if len(needs_review) else out)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="All_Merchants", index=False)
        public_sample.to_excel(writer, sheet_name="Public_Analysis_Sample", index=False)
        direct_public.to_excel(writer, sheet_name="Direct_Public", index=False)
        parent_public.to_excel(writer, sheet_name="Public_Parent", index=False)
        manual_overrides_applied.to_excel(writer, sheet_name="Manual_Overrides_Applied", index=False)
        needs_review.to_excel(writer, sheet_name="Needs_Review", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        override_template.to_excel(writer, sheet_name="Manual_Override_Template", index=False)

    print(f"Saved output workbook: {output_path}")


# ============================================================
# Main
# ============================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Add public/private listing status fields to the BNPL merchant dataset."
    )
    parser.add_argument("--input", default=str(DEFAULT_INPUT), help="Input Excel workbook.")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="Input sheet name.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output Excel workbook.")
    parser.add_argument("--cache", default=str(DEFAULT_CACHE), help="Result cache JSONL path.")
    parser.add_argument("--log", default=str(DEFAULT_LOG), help="Run log CSV path.")
    parser.add_argument("--manual-overrides", default=str(DEFAULT_MANUAL_OVERRIDES), help="Manual overrides CSV.")
    parser.add_argument("--reference-cache", default=str(DEFAULT_REFERENCE_CACHE), help="Reference-data cache directory.")
    parser.add_argument("--start", type=int, default=0, help="Start row index.")
    parser.add_argument("--limit", type=int, default=0, help="Rows to process. 0 means all rows.")
    parser.add_argument("--sleep", type=float, default=0.2, help="Seconds to sleep before source requests.")
    parser.add_argument("--save-every", type=int, default=100, help="Save after every N new classifications.")
    parser.add_argument("--offline", action="store_true", help="Use cached/reference data only; make no web requests.")
    parser.add_argument("--no-yahoo", action="store_true", help="Disable Yahoo Finance search.")
    parser.add_argument("--no-wikidata", action="store_true", help="Disable Wikidata direct/parent lookup.")
    parser.add_argument(
        "--classify-no-match-as-private",
        action="store_true",
        help="Label no-match rows as private_no_public_match and remove review flag.",
    )
    parser.add_argument("--force", action="store_true", help="Delete old result cache/log and reprocess.")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    cache_path = Path(args.cache).expanduser().resolve()
    log_path = Path(args.log).expanduser().resolve()
    manual_path = Path(args.manual_overrides).expanduser().resolve()
    reference_cache = Path(args.reference_cache).expanduser().resolve()

    df = read_input_table(input_path, args.sheet)
    print(f"Loaded input: {input_path}")
    print(f"Rows: {len(df)}")
    print(f"Manual overrides: {manual_path if manual_path.exists() else 'not found'}")

    if args.force:
        if cache_path.exists():
            cache_path.unlink()
        if log_path.exists():
            log_path.unlink()
        print("Force mode: old result cache/log removed.")

    manual_overrides = load_manual_overrides(manual_path)
    print(f"Manual overrides loaded: {len(manual_overrides)}")

    sec_records = load_sec_records(reference_cache, offline=args.offline, sleep=args.sleep)
    sec_index = PublicReferenceIndex(sec_records)
    print(f"SEC public-company reference records loaded: {len(sec_records)}")

    result_cache = load_result_cache(cache_path)
    print(f"Cached merchant listing results loaded: {len(result_cache)}")
    applied_overrides = apply_manual_overrides_to_results(df, result_cache, manual_overrides)
    if applied_overrides:
        print(f"Manual overrides applied on top of cache: {applied_overrides}")

    indices = list(df.index)
    if args.start:
        indices = [i for i in indices if i >= args.start]
    if args.limit and args.limit > 0:
        indices = indices[: args.limit]

    print(f"Rows scheduled this run: {len(indices)}")

    newly_processed = 0
    for n, idx in enumerate(indices, start=1):
        row = df.loc[idx]
        merchant_key = get_merchant_key(row)
        merchant_name = get_merchant_name(row)
        if merchant_key in manual_overrides:
            if n % 25 == 0:
                result = result_cache[merchant_key]
                print(
                    f"[{n}/{len(indices)}] manual override: {merchant_name} -> "
                    f"{result.listing_status} | {result.ticker or '-'}"
                )
            continue
        if merchant_key in result_cache:
            if n % 100 == 0:
                print(f"[{n}/{len(indices)}] cached: {merchant_name}")
            continue

        print(f"[{n}/{len(indices)}] Listing lookup: {merchant_name}")
        try:
            result = classify_one_merchant(
                row=row,
                sec_index=sec_index,
                manual_overrides=manual_overrides,
                reference_cache=reference_cache,
                offline=args.offline,
                sleep=args.sleep,
                use_yahoo=not args.no_yahoo,
                use_wikidata=not args.no_wikidata,
                classify_no_match_as_private=args.classify_no_match_as_private,
            )
        except Exception as exc:
            result = ListingResult(
                merchant_key=merchant_key,
                merchant_name=merchant_name,
                listing_status="error",
                financial_data_available="unknown",
                financial_data_level="unknown",
                listing_match_method="exception",
                listing_notes="Exception during classification.",
                needs_listing_review=True,
                error_message=str(exc),
            )

        result_cache[merchant_key] = result
        append_result_cache(cache_path, result)
        write_log_row(log_path, result)
        newly_processed += 1

        print(
            "    -> "
            f"{result.listing_status} | {result.ticker or '-'} | "
            f"{result.listing_match_confidence} | score={result.listing_match_score} | "
            f"review={result.needs_listing_review}"
        )

        if newly_processed % max(args.save_every, 1) == 0:
            save_output_workbook(df, result_cache, output_path)
            print(f"Progress saved after {newly_processed} newly processed merchants.")

    save_output_workbook(df, result_cache, output_path)

    final_df = apply_results(df, result_cache)
    print("\nDone.")
    print(f"Output: {output_path}")
    print(f"Cache: {cache_path}")
    print(f"Log: {log_path}")
    if "listing_status" in final_df.columns:
        print("\nListing status counts:")
        print(final_df["listing_status"].fillna("").astype(str).value_counts().to_string())
    if "financial_data_level" in final_df.columns:
        print("\nFinancial data level counts:")
        print(final_df["financial_data_level"].fillna("").astype(str).value_counts().to_string())
    if "needs_listing_review" in final_df.columns:
        print("\nNeeds listing review counts:")
        print(final_df["needs_listing_review"].fillna("").astype(str).value_counts().to_string())


if __name__ == "__main__":
    main()
