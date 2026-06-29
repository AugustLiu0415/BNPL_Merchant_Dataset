#!/usr/bin/env python3
"""
classify_merchants_from_web_only_calibrated_v7.py

Seventh-round calibrated BNPL merchant classifier with exact-brand override corrections.

Core principles
---------------
1. Merchant name is used only to discover official website / search evidence.
2. Industry classification is based on official website text and search evidence.
3. Existing broad_industry / sub_industry columns are preserved as original_* fields.
4. BNPL provider internal pages and navigation/locale pages are excluded before classification.
5. Confidence levels reflect evidence quality and margin, not merely whether a label exists.

Default input:
    Data_Clean/bnpl_merchant_master_long.xlsx

Default output:
    Data_Clean/bnpl_merchant_master_web_only_classified_v2.xlsx

Recommended test:
    python classify_merchants_from_web_only_calibrated_v7.py --start 13 --limit 50 --sleep 0.8 --force

Then inspect known cases:
    Apple, Audeze, Australia, Bellroy, DHGate, DavidYurman, Diversity&inclusion,
    DoMyOwnPestControl.com, Dyson, Eufy, Fordevelopers, HannaAndersson, Howitworks, Instacart.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import random
import re
import time
from collections import defaultdict
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup


# ============================================================
# Paths
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent
DATA_CLEAN = PROJECT_ROOT / "Data_Clean"

DEFAULT_INPUT = DATA_CLEAN / "bnpl_merchant_master_long.xlsx"
DEFAULT_OUTPUT = DATA_CLEAN / "bnpl_merchant_master_web_only_classified_v7_parse_safe.xlsx"
DEFAULT_CACHE = DATA_CLEAN / "web_only_classification_v7_cache.jsonl"
DEFAULT_LOG = DATA_CLEAN / "web_only_classification_v7_parse_safe_run_log.csv"


# ============================================================
# Network settings
# ============================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

REQUEST_TIMEOUT = 18

BNPL_DOMAINS = {
    "affirm.com",
    "klarna.com",
    "afterpay.com",
    "clearpay.co.uk",
    "zip.co",
}

BAD_RESULT_DOMAIN_PARTS = [
    "facebook.",
    "instagram.",
    "linkedin.",
    "twitter.",
    "x.com",
    "tiktok.",
    "pinterest.",
    "youtube.",
    "reddit.",
    "wikipedia.",
    "wikidata.",
    "yelp.",
    "trustpilot.",
    "bbb.",
    "glassdoor.",
    "indeed.",
    "crunchbase.",
    "retailmenot.",
    "coupon",
    "shop.app",
    "linktr.ee",
    "linktree",
    "amazon.com",
    "ebay.com",
]


# ============================================================
# Taxonomy
# ============================================================

TAXONOMY_TERMS: dict[str, dict[str, list[str]]] = {
    "Sports & Outdoor": {
        "Athletic footwear & apparel": [
            "running shoes", "athletic shoes", "sneakers", "trainers", "sportswear",
            "sports apparel", "athletic apparel", "activewear", "cleats",
            "basketball shoes", "tennis shoes", "trail running shoes",
            "wrestling shoes", "volleyball shoes", "golf shoes", "soccer shoes",
            "running apparel", "sports bras", "tights", "leggings",
        ],
        "Outdoor apparel & gear": [
            "outdoor apparel", "outdoor gear", "hiking", "climbing", "skiing",
            "ski", "snowboard", "shell jacket", "insulated jacket", "backcountry",
            "camping", "trail running", "gore-tex", "technical apparel",
            "arc'teryx", "arcteryx",
        ],
        "Sporting goods / outdoor": [
            "sporting goods", "camping", "hiking", "climbing", "fishing", "hunting",
            "golf", "cycling", "bike", "bicycle", "ski", "snowboard", "surf",
            "skateboard", "fitness equipment", "gym equipment", "barbell", "dumbbell",
            "tactical gear", "tactical apparel",
        ],
    },
    "Apparel & Fashion": {
        "Activewear / athleisure": [
            "leggings", "sports bra", "sports bras", "yoga pants", "activewear",
            "athleisure", "workout clothes", "training clothes", "performance apparel",
            "biker shorts", "sweatpants", "hoodie", "tank top", "running shorts",
            "yoga clothes", "studio to street",
        ],
        "Footwear": [
            "shoes", "sneakers", "boots", "sandals", "heels", "flats", "footwear",
            "running shoes", "athletic shoes", "dress shoes", "loafers",
        ],
        "Bags & accessories": [
            "wallet", "wallets", "bags", "backpack", "backpacks", "tote", "totes",
            "crossbody", "briefcase", "duffel", "pouch", "pouches", "phone case",
            "phone cases", "card holder", "cardholders", "passport wallet",
            "travel wallet", "key cover", "accessories",
        ],
        "Jewelry & watches": [
            "jewelry", "jewellery", "bracelets", "bracelet", "rings", "ring",
            "necklaces", "necklace", "earrings", "earring", "watches", "watch",
            "timepieces", "diamonds", "gold", "sterling silver", "david yurman",
        ],
        "Designer fashion & accessories": [
            "designer handbags", "handbags", "designer clothing", "luxury fashion",
            "watches", "wallets", "purses", "jewelry", "jewellery", "sunglasses",
            "accessories", "menswear", "womenswear", "ready-to-wear", "luxury resale",
            "designer resale", "consignment",
        ],
        "Fashion / apparel": [
            "apparel", "clothing", "fashion", "dresses", "tops", "bottoms", "denim",
            "jeans", "shirts", "t-shirts", "outerwear", "coats", "jackets", "swimwear",
            "lingerie", "underwear", "boutique", "streetwear", "women's clothing",
            "men's clothing",
        ],
    },
    "Electronics": {
        "Audio electronics": [
            "headphones", "headphone", "earbuds", "earphones", "speakers", "speaker",
            "audio", "hi-fi", "hifi", "audiophile", "planar magnetic", "dac",
            "amplifier", "amplifiers", "soundbar", "noise cancelling", "bluetooth speaker",
            "wireless headphones",
        ],
        "Smart home electronics": [
            "smart home", "security camera", "security cameras", "video doorbell",
            "doorbell", "robot vacuum", "robot vacuums", "smart lock", "smart locks",
            "home security", "baby monitor", "smart lighting", "eufy", "anker innovations",
        ],
        "Consumer electronics": [
            "iphone", "ipad", "macbook", "mac computer", "imac", "apple watch",
            "airpods", "smartphone", "tablet", "laptop", "computer", "desktop",
            "consumer electronics", "electronics", "charger", "charging cable",
            "camera", "drone", "monitor", "keyboard", "mouse", "gaming pc",
            "router", "smartwatch", "power station", "solar generator", "portable power",
        ],
    },
    "Marketplace / General Retail": {
        "Department store / omnichannel retail": [
            "department store", "department stores", "omnichannel retailer", "retail chain",
            "clothing, shoes, home", "home, bedding, furniture", "kohl's", "kohls",
            "macy's", "macys", "nordstrom", "jcpenney", "j.c. penney", "bloomingdale",
            "dillard", "belk",
        ],
        "Big-box / general merchandise": [
            "big box", "big-box", "general merchandise", "household essentials",
            "walmart", "target", "costco", "sam's club", "sams club", "one-stop shop",
        ],
        "Marketplace": [
            "marketplace", "third-party sellers", "online marketplace",
            "millions of products", "amazon", "ebay", "etsy",
        ],
        "Cross-border marketplace": [
            "cross-border", "cross border", "b2b marketplace", "b2c marketplace",
            "global wholesale", "wholesale marketplace", "international marketplace",
            "china wholesale", "dhgate",
        ],
    },
    "Beauty & Personal Care": {
        "Cosmetics / makeup": [
            "makeup", "cosmetics", "brows", "eyebrow", "eyeshadow", "lipstick",
            "foundation", "mascara", "concealer", "blush", "palette",
        ],
        "Skincare": [
            "skincare", "skin care", "serum", "moisturizer", "cleanser", "spf", "sunscreen",
            "exfoliant", "dermalogica", "estee lauder", "estée lauder",
        ],
        "Haircare": [
            "haircare", "hair care", "shampoo", "conditioner", "salon", "hair dryer",
            "hair straightener", "airwrap", "styler",
        ],
        "Fragrance": ["fragrance", "perfume", "cologne"],
        "Grooming / body care": [
            "grooming", "shaving", "razor", "body care", "deodorant", "electric shaver",
            "beard trimmer", "oral-b", "braun",
        ],
    },
    "Home & Furniture": {
        "Furniture": [
            "furniture", "sofa", "couch", "chair", "table", "desk", "dresser",
            "bed frame", "sectional", "dining table", "article", "castlery", "ikea",
        ],
        "Cookware / kitchenware": [
            "cookware", "kitchenware", "pots", "pans", "knife", "knives",
            "kitchen essentials",
        ],
        "Mattress / bedding": [
            "mattress", "bedding", "sheets", "duvet", "pillow", "weighted blanket",
            "blanket", "bearaby", "avocado mattress",
        ],
        "Home decor": [
            "home decor", "home décor", "decor", "rugs", "wall art", "lighting",
            "candles", "crate and barrel",
        ],
        "Home appliances": [
            "vacuum cleaner", "vacuum cleaners", "vacuum", "air purifier", "air purifiers",
            "fan", "fans", "heater", "humidifier", "dyson", "appliances",
        ],
        "Outdoor living & grills": [
            "grill", "grills", "bbq", "barbecue", "outdoor kitchen", "patio furniture",
            "outdoor living", "smokers", "fire pit", "bbq guys", "bbqguys",
        ],
        "Pest control & lawn care": [
            "pest control", "insecticide", "pesticide", "weed killer", "lawn care",
            "lawn & garden", "termite", "mosquito control", "do my own pest control",
        ],
        "Home improvement / hardware": [
            "hardware store", "tools", "home improvement", "paint", "plumbing",
            "electrical", "ace hardware", "lawn and garden",
        ],
        "Home goods": ["home goods", "housewares", "bath", "patio", "garden"],
    },
    "Travel & Ticketing": {
        "Airline": [
            "airline", "airlines", "flights", "flight", "airfare", "book a flight",
            "delta air lines", "jetblue", "united airlines",
        ],
        "Flight booking": [
            "flight booking", "book flights", "cheap flights", "expedia flights",
            "compare flights", "airline tickets",
        ],
        "Travel booking / lodging marketplace": [
            "hotel", "hotels", "booking", "vacation", "resort", "travel",
            "vacation rental", "lodging", "airbnb", "hotels.com", "expedia",
        ],
        "Event tickets": ["tickets", "ticketing", "concert tickets", "event tickets", "ticketmaster"],
        "Luggage / travel accessories": ["luggage", "suitcase", "carry-on", "travel bag"],
    },
    "Food & Grocery": {
        "Grocery delivery": [
            "grocery delivery", "grocery pickup", "same-day grocery", "personal shopper",
            "deliver groceries", "instacart", "food delivery from stores",
        ],
        "Grocery": ["grocery", "groceries", "supermarket", "food market"],
        "Restaurant / food delivery": ["restaurant", "food delivery", "meal delivery", "order food"],
        "Coffee / tea": ["coffee", "tea", "espresso"],
        "Snacks / sweets": ["snacks", "chocolate", "candy", "cookies", "bakery"],
        "Food & beverage": ["food", "beverage", "drink", "meal kit"],
    },
    "Baby & Kids": {
        "Kids clothing": [
            "kids clothing", "children's clothing", "baby clothes", "kids pajamas",
            "pajamas", "hannas", "hanna andersson", "children's pajamas",
        ],
        "Baby gear / nursery": ["baby", "nursery", "stroller", "car seat", "crib", "infant", "toddler"],
        "Toys": ["toys", "toy store", "dolls", "games for kids"],
    },
    "Automotive": {
        "Auto parts": ["auto parts", "car parts", "truck parts", "vehicle parts"],
        "Tires / wheels": ["tires", "wheels", "rims"],
        "Motorcycle accessories": ["motorcycle", "motorbike"],
        "Off-road accessories": ["off-road", "4x4", "jeep accessories"],
    },
    "Pet": {
        "Pet supplies": ["pet supplies", "dog", "cat", "pets", "pet toys", "pet accessories", "chewy"],
        "Pet food": ["pet food", "dog food", "cat food"],
        "Pet healthcare": ["pet health", "veterinary", "flea", "tick"],
    },
    "Health & Wellness": {
        "Dental care": ["dental", "oral care", "floss", "toothbrush", "cocofloss"],
        "Nutrition / supplements": [
            "supplements", "vitamins", "protein powder", "pre-workout",
            "nutrition supplements",
        ],
        "Pharmacy / medical": ["pharmacy", "medical", "healthcare", "prescription"],
        "Vision care": ["contact lenses", "contacts", "vision care", "eyeglasses"],
        "Health / wellness products": ["wellness products", "health products", "sleep aid", "therapy products"],
    },
    "Education / Books": {
        "Books / textbooks": ["books", "bookstore", "textbooks", "audiobooks"],
        "Online training / certification": ["online courses", "training", "certification", "learning"],
        "Education / learning": ["education", "school supplies", "tutoring"],
    },
    "Entertainment / Music": {
        "Music artist merchandise / fan shop": [
            "artist merchandise", "band merchandise", "fan shop", "official store",
            "tour merch", "music merch",
        ],
        "Records / vinyl / music media": ["vinyl", "records", "albums", "cds", "cassettes", "music media"],
    },
    "Entertainment / Gaming": {
        "Video games / consoles": ["video games", "gaming", "console", "playstation", "xbox", "nintendo"],
        "Collectibles / comics": ["collectibles", "trading cards", "anime", "comics", "comic books"],
        "Entertainment merchandise": ["fan merchandise", "movie merchandise", "tv merchandise"],
    },
    "Gifts & Flowers": {
        "Flowers": ["flowers", "florist", "bouquets", "flower delivery"],
        "Gift baskets": ["gift baskets"],
        "Gifts": ["gifts", "personalized gifts", "greeting cards"],
    },
    "Arts & Crafts": {
        "Art & craft supplies": ["craft supplies", "art supplies", "scrapbooking"],
        "Fabric / sewing / yarn": ["fabric", "sewing", "yarn", "knitting"],
        "Stationery": ["stationery", "paper goods"],
    },
    "Services / Insurance": {
        "Insurance": ["insurance"],
        "Software service": ["software", "saas", "app service"],
        "Services": ["service", "membership", "subscription"],
    },
    "Nonprofit / Public Service": {
        "Nonprofit / public service": ["nonprofit", "charity", "foundation", "donate"],
    },
}


# ============================================================
# V5 taxonomy extensions from 100-row validation sample
# ============================================================

def apply_taxonomy_v5_extensions() -> None:
    """
    Add product/service categories revealed by the 100-row validation batch.
    This keeps the main taxonomy readable and makes calibration changes auditable.
    """
    def add_terms(broad: str, sub: str, terms: list[str]) -> None:
        TAXONOMY_TERMS.setdefault(broad, {}).setdefault(sub, [])
        existing = set(TAXONOMY_TERMS[broad][sub])
        for term in terms:
            if term not in existing:
                TAXONOMY_TERMS[broad][sub].append(term)
                existing.add(term)

    add_terms(
        "Sports & Outdoor",
        "Cycling / bike resale",
        [
            "cycling", "bike", "bikes", "bicycle", "bicycles",
            "used bikes", "bike marketplace", "bike resale", "cycling gear",
            "the pro's closet", "the pros closet",
        ],
    )

    add_terms(
        "Electronics",
        "Wearable health tech",
        [
            "smart ring", "oura ring", "wearable", "wearables",
            "sleep tracking", "sleep tracker", "health tracking",
            "heart rate", "activity tracking",
        ],
    )

    add_terms(
        "Electronics",
        "Phone accessories",
        [
            "phone accessories", "phone grip", "phone grips", "phone stand",
            "phone stands", "phone case", "phone cases", "popsockets",
            "popgrip", "mobile accessories",
        ],
    )

    add_terms(
        "Food & Grocery",
        "Coffee / tea",
        [
            "nespresso", "coffee capsules", "coffee pods", "espresso machine",
            "espresso machines", "coffee machine", "coffee machines",
        ],
    )

    add_terms(
        "Gifts & Flowers",
        "Photo printing & personalized gifts",
        [
            "photo books", "photo book", "photo printing", "personalized gifts",
            "custom cards", "photo cards", "shutterfly", "prints",
            "calendars", "wedding invitations",
        ],
    )

    add_terms(
        "Travel & Ticketing",
        "Theme parks / attractions",
        [
            "theme park", "theme parks", "amusement park", "amusement parks",
            "roller coasters", "six flags", "season pass", "park tickets",
            "attractions",
        ],
    )

    add_terms(
        "Food & Grocery",
        "Restaurant / food delivery",
        [
            "doordash", "restaurant delivery", "restaurants", "delivery app",
            "food delivery", "meal delivery", "dashpass", "order food",
        ],
    )

    add_terms(
        "Services / Insurance",
        "Rideshare / local transportation",
        [
            "rideshare", "ride share", "ride-hailing", "ride hailing",
            "lyft", "taxi", "driver", "rides", "transportation network",
        ],
    )

    add_terms(
        "Telecom",
        "Wireless carrier",
        [
            "wireless carrier", "mobile carrier", "cell phone plans",
            "phone plans", "5g", "t-mobile", "tmobile",
            "wireless service", "mobile network",
        ],
    )

    add_terms(
        "Office & Business Supplies",
        "Office supplies",
        [
            "office supplies", "office depot", "officedepot", "paper",
            "printer ink", "toner", "office furniture", "school supplies",
            "business supplies", "printing services",
        ],
    )

    add_terms(
        "Automotive",
        "Auto parts / garage",
        [
            "garage", "auto parts", "car parts", "performance parts",
            "truck parts", "automotive parts", "tps garage",
        ],
    )


apply_taxonomy_v5_extensions()

# ============================================================
# V6 taxonomy extensions and exact-brand calibration targets
# ============================================================

def apply_taxonomy_v6_extensions() -> None:
    """
    Add small, auditable categories and terms revealed by the second 100-row validation batch.
    """
    def add_terms(broad: str, sub: str, terms: list[str]) -> None:
        TAXONOMY_TERMS.setdefault(broad, {}).setdefault(sub, [])
        existing = set(TAXONOMY_TERMS[broad][sub])
        for term in terms:
            if term not in existing:
                TAXONOMY_TERMS[broad][sub].append(term)
                existing.add(term)

    add_terms(
        "Beauty & Personal Care",
        "Fragrance",
        [
            "fragrances", "perfumes", "eau de parfum", "eau de toilette",
            "scent", "scents", "le labo", "le labo fragrances",
        ],
    )

    add_terms(
        "Baby & Kids",
        "Toys",
        [
            "lego", "lego sets", "lego bricks", "building toys", "construction toys",
            "toy bricks", "play sets", "children's toys", "kids toys",
        ],
    )

    add_terms(
        "Marketplace / General Retail",
        "Collectibles / trading cards marketplace",
        [
            "tcgplayer", "trading cards", "collectible cards", "card marketplace",
            "pokemon cards", "magic the gathering", "yugioh", "sports cards",
        ],
    )

    add_terms(
        "Marketplace / General Retail",
        "Shopping platform / app",
        [
            "shop app", "shop.app", "shopping app", "shopify shop app", "package tracking app",
            "discover brands", "online shopping app",
        ],
    )

    add_terms(
        "Apparel & Fashion",
        "Fashion / apparel",
        ["quince", "sustainable clothing", "cashmere sweater", "silk clothing"],
    )


apply_taxonomy_v6_extensions()


# ============================================================
# Data model
# ============================================================

@dataclass
class ClassificationResult:
    cache_key: str
    merchant_name: str
    official_url: str
    final_url: str
    url_discovery_method: str
    evidence_title: str
    evidence_meta_description: str
    evidence_schema_text: str
    evidence_nav_text: str
    evidence_heading_text: str
    evidence_body_snippet: str
    evidence_search_snippet: str
    verified_broad_industry: str
    verified_sub_industry: str
    classification_confidence: str
    verification_status: str
    classification_reason: str
    needs_manual_review: bool
    error_message: str


# ============================================================
# Basic utilities
# ============================================================

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value)
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_for_cache(value: str) -> str:
    value = clean_text(value).lower()
    value = value.replace("&", " and ")
    value = value.replace("™", "").replace("®", "")
    value = re.sub(r"\b(official store|official|online store|store|shop)\b", "", value)
    value = re.sub(r"\b(inc|llc|ltd|limited|co|company|corp|corporation|plc)\b", "", value)
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


def normalize_for_filter(value: str) -> str:
    value = clean_text(value).lower()
    value = value.replace("™", "").replace("®", "")
    value = value.replace("&", " and ")
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


def get_domain(url: str) -> str:
    try:
        domain = urlparse(url).netloc.lower()
        if domain.startswith("www."):
            domain = domain[4:]
        return domain
    except Exception:
        return ""


def is_http_url(url: str) -> bool:
    return isinstance(url, str) and url.lower().startswith(("http://", "https://"))


def is_bnpl_url(url: str) -> bool:
    domain = get_domain(url)
    return any(domain == d or domain.endswith("." + d) for d in BNPL_DOMAINS)


def is_bad_result_url(url: str) -> bool:
    domain = get_domain(url)
    if not domain:
        return True
    return any(part in domain for part in BAD_RESULT_DOMAIN_PARTS)


def merchant_cache_key(row: pd.Series) -> str:
    name = ""
    for col in ["normalized_company_name", "canonical_company_name", "raw_company_name", "company_name", "merchant_name", "name"]:
        if col in row.index and clean_text(row.get(col)):
            name = clean_text(row.get(col))
            break
    country = clean_text(row.get("country", ""))
    raw_key = f"{normalize_for_cache(name)}||{country.lower()}"
    return hashlib.sha1(raw_key.encode("utf-8")).hexdigest()


def get_merchant_name(row: pd.Series) -> str:
    for col in [
        "canonical_company_name", "raw_company_name", "company_name",
        "merchant_name", "name", "normalized_company_name",
    ]:
        if col in row.index:
            value = clean_text(row.get(col))
            if value:
                return value
    return ""


def make_search_friendly_name(name: str) -> str:
    """
    Convert compact scraped merchant names into search-friendly names.

    This is used only for web search / official URL discovery.
    It is NOT used as direct industry evidence by itself.

    Examples:
        HannaAndersson -> Hanna Andersson
        DeltaAirLines -> Delta Air Lines
        B&HPhoto -> B&H Photo
        Crate&Barrel -> Crate and Barrel
    """
    text = clean_text(name)
    if not text:
        return ""

    replacements = {
        "B&HPhoto": "B&H Photo",
        "Crate&Barrel": "Crate and Barrel",
        "DeltaAirLines": "Delta Air Lines",
        "HannaAndersson": "Hanna Andersson",
        "DavidYurman": "David Yurman",
        "HolabirdSports": "Holabird Sports",
        "AcademySports+Outdoors": "Academy Sports Outdoors",
        "BestBuy": "Best Buy",
        "FootLocker": "Foot Locker",
        "GameStop": "GameStop",
        "EsteeLauder": "Estee Lauder",
        "EcoFlowTech": "EcoFlow Tech",
        "DoMyOwnPestControl.com": "Do My Own Pest Control",
        "BBQGuys": "BBQ Guys",
        "FARFETCH": "Farfetch",
        "LeLaboFragrances": "Le Labo Fragrances",
        "L.L.Bean": "L.L.Bean",
        "Lowe'sHomeImprovement": "Lowe's Home Improvement",
        "Madewell": "Madewell",
        "Moreinformationaboutourcookiepolicy": "More information about our cookie policy",
        "Nespresso": "Nespresso",
        "NewBalanceRunning": "New Balance Running",
        "OfficeDepot": "Office Depot",
        "OuraRing": "Oura Ring",
        "PopSockets": "PopSockets",
        "RoomsToGoInc.": "Rooms To Go",
        "SaksFifthAvenue": "Saks Fifth Avenue",
        "SaksOFF5TH": "Saks OFF 5TH",
        "Shopnow": "Shop now",
        "Shutterfly": "Shutterfly",
        "SixFlags": "Six Flags",
        "Skiptomaincontent": "Skip to main content",
        "SouthwestAirlines": "Southwest Airlines",
        "SurLaTable": "Sur La Table",
        "TPSGARAGELLC": "TPS Garage LLC",
        "TheLEGO": "LEGO",
        "ThePro'sCloset,Inc.": "The Pro's Closet",
        "TOMSShoes": "TOMS Shoes",
        "UnitedAirlines": "United Airlines",
        "WilliamsSonoma": "Williams Sonoma",
        "WilsonSportingGoods": "Wilson Sporting Goods",
        "DoorDash": "DoorDash",
        "GoogleStore": "Google Store",
        "Booking.com": "Booking.com",
        "StubHub": "StubHub",
        "Shop.app": "Shop app",
        "T-Mobile": "T-Mobile",
        "Zara": "Zara",
    }
    if text in replacements:
        return replacements[text]

    text = text.replace("&", " and ")
    text = text.replace("+", " ")
    text = re.sub(r"\.com$", "", text, flags=re.IGNORECASE)
    text = text.replace("™", "").replace("®", "")

    # Split CamelCase and common alpha/number boundaries.
    text = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", text)
    text = re.sub(r"(?<=[A-Z])(?=[A-Z][a-z])", " ", text)
    text = re.sub(r"(?<=[A-Za-z])(?=[0-9])", " ", text)
    text = re.sub(r"(?<=[0-9])(?=[A-Za-z])", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def get_provider_name(row: pd.Series) -> str:
    if "provider" in row.index:
        return clean_text(row.get("provider"))
    if "bnpl_provider" in row.index:
        return clean_text(row.get("bnpl_provider"))
    return ""


# ============================================================
# Pre-classification filters
# ============================================================

def get_name_candidates(row: pd.Series) -> list[str]:
    candidates = []
    for col in ["normalized_company_name", "canonical_company_name", "company_name", "merchant_name", "name"]:
        if col in row.index and clean_text(row.get(col)):
            candidates.append(normalize_for_filter(clean_text(row.get(col))))
    if not candidates and "raw_company_name" in row.index:
        candidates.append(normalize_for_filter(clean_text(row.get("raw_company_name"))))
    return list(dict.fromkeys(candidates))


def is_bnpl_provider_internal_page(row: pd.Series) -> tuple[bool, str]:
    provider = normalize_for_filter(get_provider_name(row))
    name_candidates = get_name_candidates(row)
    if not provider or not name_candidates:
        return False, ""

    provider_internal_names = {
        "affirm": {
            "affirm", "affirmcard", "affirmcares", "affirmmoney", "affirmapp",
            "affirmsavings", "affirmshop", "shopaffirm",
        },
        "klarna": {
            "klarna", "klarnacard", "klarnaapp", "klarnashop", "shopklarna",
        },
        "afterpay": {
            "afterpay", "afterpaycard", "afterpayapp", "afterpayshop", "shopafterpay",
        },
        "clearpay": {
            "clearpay", "clearpaycard", "clearpayapp", "clearpayshop", "shopclearpay",
        },
        "zip": {
            "zip", "zipco", "zipcard", "zipapp", "zipshop", "shopzip",
        },
    }

    internal_set = provider_internal_names.get(provider, set())
    for name in name_candidates:
        if name in internal_set:
            return True, f"Provider internal page/product detected: provider={provider}, merchant_name={name}"
        if provider in {"affirm", "klarna", "afterpay", "clearpay"}:
            if name.startswith(provider) and len(name) > len(provider):
                return True, f"Provider internal page/product detected by provider prefix: provider={provider}, merchant_name={name}"
        if provider == "zip" and name in {"zipcard", "zipapp", "zipshop", "zipco"}:
            return True, f"Zip internal page/product detected: merchant_name={name}"
    return False, ""


def is_non_merchant_navigation_page(row: pd.Series) -> tuple[bool, str]:
    """
    Exclude locale/navigation/help/corporate pages that are not retailer merchants.
    This filter is intentionally conservative and relies on exact normalized labels.
    """
    names = set(get_name_candidates(row))

    exact_non_merchant = {
        "australia", "canadaen", "canadafr", "canada", "unitedstates", "usa",
        "unitedkingdom", "uk", "diversityandinclusion", "fordevelopers",
        "developers", "howitworks", "howitworks", "learnmore", "support",
        "help", "faq", "contactus", "careers", "press", "about", "aboutus",
        "moreinformationaboutourcookiepolicy", "cookiepolicy", "privacypolicy",
        "privacychoices", "skiptomaincontent", "skipnavigation", "skiptonavigation",
        "shopnow", "learnmore", "readmore", "findastore",
    }

    for name in names:
        if name in exact_non_merchant:
            return True, f"Non-merchant navigation/locale/corporate page detected: merchant_name={name}"

    # Common page labels with provider URL evidence are almost surely internal pages.
    raw_all = " ".join(clean_text(row.get(c, "")) for c in row.index).lower()
    if any(name in names for name in {"howitworks", "fordevelopers", "diversityandinclusion"}):
        return True, "Non-merchant internal informational page detected."

    # Keep this conservative: do not exclude real merchants with country words in their names.
    return False, ""


def make_exclusion_result(
    row: pd.Series,
    broad: str,
    sub: str,
    status: str,
    reason: str,
) -> ClassificationResult:
    return ClassificationResult(
        cache_key=merchant_cache_key(row),
        merchant_name=get_merchant_name(row),
        official_url="",
        final_url="",
        url_discovery_method=status,
        evidence_title="",
        evidence_meta_description="",
        evidence_schema_text="",
        evidence_nav_text="",
        evidence_heading_text="",
        evidence_body_snippet="",
        evidence_search_snippet="",
        verified_broad_industry=broad,
        verified_sub_industry=sub,
        classification_confidence="high",
        verification_status=status,
        classification_reason=reason,
        needs_manual_review=False,
        error_message="",
    )


# ============================================================
# URL handling
# ============================================================

def get_existing_urls(row: pd.Series) -> list[str]:
    urls: list[str] = []
    for col in ["official_url", "merchant_url", "source_page", "classification_source_url", "search_result_url"]:
        if col not in row.index:
            continue
        value = clean_text(row.get(col))
        if not value:
            continue
        parts = value.split("|") if "|" in value else [value]
        for part in parts:
            part = clean_text(part)
            if is_http_url(part):
                urls.append(part)
    out = []
    seen = set()
    for u in urls:
        if u not in seen:
            out.append(u)
            seen.add(u)
    return out


def extract_redirect_dest(url: str) -> str:
    try:
        parsed = urlparse(url)
        qs = parse_qs(parsed.query)
        for key in ["dest_url", "url", "redirect_url", "merchant_url", "target"]:
            if key in qs and qs[key]:
                candidate = unquote(qs[key][0])
                if is_http_url(candidate):
                    return candidate
    except Exception:
        pass
    return ""


def domain_name_match_score(merchant_name: str, url: str) -> int:
    name = normalize_for_cache(merchant_name)
    domain = get_domain(url)
    domain_core = re.sub(r"[^a-z0-9]+", "", domain)
    domain_core = domain_core.replace("com", "").replace("co", "").replace("net", "")
    if not name or not domain_core:
        return 0
    if name == domain_core:
        return 5
    if name in domain_core:
        return 4
    if domain_core in name and len(domain_core) >= 4:
        return 3
    tokens = [t for t in re.split(r"[^a-z0-9]+", merchant_name.lower()) if len(t) >= 3]
    return sum(1 for t in tokens if t in domain_core)


def fetch_url(url: str, timeout: int = REQUEST_TIMEOUT) -> tuple[bool, str, str, str]:
    if not is_http_url(url):
        return False, url, "", "not an HTTP URL"
    try:
        response = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        final_url = response.url
        if response.status_code >= 400:
            return False, final_url, "", f"HTTP {response.status_code}"
        html = response.text or ""
        content_type = response.headers.get("content-type", "").lower()
        if "text/html" not in content_type and "<html" not in html[:500].lower():
            return False, final_url, "", f"non-html content type: {content_type}"
        return True, final_url, html, ""
    except requests.exceptions.Timeout:
        return False, url, "", "timeout"
    except requests.exceptions.RequestException as exc:
        return False, url, "", f"request error: {exc}"
    except Exception as exc:
        return False, url, "", f"unexpected error: {exc}"


# ============================================================
# Search
# ============================================================

def serpapi_search(query: str, max_results: int = 5) -> list[dict[str, str]]:
    api_key = os.getenv("SERPAPI_API_KEY", "").strip()
    if not api_key:
        return []
    try:
        response = requests.get(
            "https://serpapi.com/search",
            params={"engine": "google", "q": query, "api_key": api_key, "num": max_results},
            headers=HEADERS,
            timeout=20,
        )
        response.raise_for_status()
        data = response.json()
        return [
            {
                "title": clean_text(item.get("title")),
                "href": clean_text(item.get("link")),
                "body": clean_text(item.get("snippet")),
            }
            for item in data.get("organic_results", [])[:max_results]
        ]
    except Exception as exc:
        print(f"SerpApi search failed: {exc}")
        return []


def ddg_search(query: str, max_results: int = 5) -> list[dict[str, str]]:
    try:
        try:
            from ddgs import DDGS  # type: ignore
        except Exception:
            from duckduckgo_search import DDGS  # type: ignore
        with DDGS(timeout=20) as ddgs:
            raw_results = list(ddgs.text(query, max_results=max_results))
        return [
            {
                "title": clean_text(item.get("title")),
                "href": clean_text(item.get("href") or item.get("url")),
                "body": clean_text(item.get("body")),
            }
            for item in raw_results
        ]
    except Exception as exc:
        print(f"DDG search failed: {exc}")
        return []


def web_search(query: str, max_results: int = 5) -> list[dict[str, str]]:
    results = serpapi_search(query, max_results=max_results)
    if results:
        return results
    return ddg_search(query, max_results=max_results)


def build_search_evidence_bundle(merchant_name: str, country: str, max_results: int = 5) -> str:
    country_hint = ""
    if "united kingdom" in country.lower():
        country_hint = " UK"
    elif "united states" in country.lower():
        country_hint = " US"

    search_names = []
    for candidate in [merchant_name, make_search_friendly_name(merchant_name)]:
        candidate = clean_text(candidate)
        if candidate and candidate.lower() not in {x.lower() for x in search_names}:
            search_names.append(candidate)

    queries = []
    for search_name in search_names:
        queries.extend([
            f'"{search_name}" official website store{country_hint}',
            f'"{search_name}" products',
            f'"{search_name}" product categories',
            f'"{search_name}" what does it sell',
            f'"{search_name}" retailer brand products',
        ])

    # Targeted search hints for recurring edge cases where compact scraped names
    # or JavaScript-heavy official websites produce weak website evidence.
    # These hints improve evidence retrieval; the query text itself is NOT used
    # as classification evidence below.
    normalized_search_name = normalize_for_filter(merchant_name)
    special_queries_by_name = {
        "hannaandersson": [
            '"Hanna Andersson" kids pajamas children clothing baby toddler',
            '"Hanna Andersson" childrens clothing pajamas official',
        ],
        "bhphoto": [
            '"B&H Photo" cameras video audio computers electronics',
        ],
        "farfetch": [
            '"Farfetch" luxury fashion designer clothing shoes accessories',
        ],
        "bestbuy": [
            '"Best Buy" consumer electronics computers appliances cameras',
        ],
        "lelabofragrances": [
            '"Le Labo" fragrances perfume cologne candles',
        ],
        "lyft": [
            '"Lyft" rideshare ride hailing transportation app',
        ],
        "nespresso": [
            '"Nespresso" coffee capsules espresso machines',
        ],
        "officedepot": [
            '"Office Depot" office supplies printer ink business supplies',
        ],
        "ouraring": [
            '"Oura Ring" smart ring sleep tracking wearable health tech',
        ],
        "popsockets": [
            '"PopSockets" phone grip phone stand mobile accessories',
        ],
        "shutterfly": [
            '"Shutterfly" photo books photo printing personalized gifts',
        ],
        "sixflags": [
            '"Six Flags" theme park amusement park tickets season pass',
        ],
        "surlatable": [
            '"Sur La Table" cookware kitchenware kitchen tools',
        ],
        "tpsgaragellc": [
            '"TPS Garage" automotive parts performance parts garage',
        ],
        "theprosclosetinc": [
            "\"The Pro\'s Closet\" used bikes cycling marketplace",
        ],
        "williamsomona": [
            '"Williams Sonoma" cookware kitchenware home goods',
        ],
        "williamssonoma": [
            '"Williams Sonoma" cookware kitchenware home goods',
        ],
        "zara": [
            '"Zara" fashion apparel clothing official store',
        ],
        "doordash": [
            '"DoorDash" food delivery restaurant delivery order food',
        ],
        "tmobile": [
            '"T-Mobile" wireless carrier phone plans mobile network',
        ],
        "llbean": [
            '"L.L.Bean" outdoor clothing outdoor gear boots',
        ],
        "rei": [
            '"REI" outdoor gear camping hiking cycling apparel',
        ],
        "patagonia": [
            '"Patagonia" outdoor clothing gear jackets',
        ],
        "newbalancerunning": [
            '"New Balance" running shoes athletic footwear apparel',
        ],
        "roomstogoinc": [
            '"Rooms To Go" furniture store sofas bedroom furniture',
        ],
        "skiptomaincontent": [
            '"Skip to main content" navigation link',
        ],
        "moreinformationaboutourcookiepolicy": [
            '"cookie policy" website privacy informational page',
        ],
    }
    queries.extend(special_queries_by_name.get(normalized_search_name, []))

    snippets = []
    for query in queries:
        results = web_search(query, max_results=max_results)
        for item in results:
            href = clean_text(item.get("href"))
            title = clean_text(item.get("title"))
            body = clean_text(item.get("body"))
            if href and is_bad_result_url(href):
                continue
            snippet = clean_text(f"TITLE: {title} | SNIPPET: {body} | URL: {href}")
            if snippet:
                snippets.append(snippet)
        time.sleep(0.4)
    return clean_text(" ".join(snippets))[:10000]

def choose_official_search_result(results: list[dict[str, str]], merchant_name: str) -> tuple[str, dict[str, str], str]:
    valid = []
    for r in results:
        href = clean_text(r.get("href"))
        if not is_http_url(href):
            continue
        if is_bad_result_url(href):
            continue
        if is_bnpl_url(href):
            continue
        valid.append(r)
    if not valid:
        return "", {}, "no valid search result"

    scored = []
    for r in valid:
        href = clean_text(r.get("href"))
        score = domain_name_match_score(merchant_name, href)
        title_body = f"{clean_text(r.get('title'))} {clean_text(r.get('body'))}".lower()
        if "official" in title_body:
            score += 2
        if "shop" in title_body or "store" in title_body:
            score += 1
        scored.append((score, r))
    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_result = scored[0]
    return clean_text(best_result.get("href")), best_result, f"selected search result with official-site score={best_score}"


# ============================================================
# HTML evidence extraction
# ============================================================

def safe_json_loads(raw: str) -> list[Any]:
    raw = clean_text(raw)
    if not raw:
        return []
    try:
        obj = json.loads(raw)
        return obj if isinstance(obj, list) else [obj]
    except Exception:
        return []


def flatten_json(obj: Any) -> list[Any]:
    out: list[Any] = []

    def visit(x: Any) -> None:
        if isinstance(x, dict):
            out.append(x)
            for value in x.values():
                if isinstance(value, (dict, list)):
                    visit(value)
        elif isinstance(x, list):
            for item in x:
                visit(item)

    visit(obj)
    return out


def extract_schema_text(soup: BeautifulSoup) -> str:
    pieces = []
    scripts = soup.find_all("script", attrs={"type": re.compile(r"application/ld\+json", re.I)})
    for script in scripts:
        raw = script.string or script.get_text() or ""
        for obj in safe_json_loads(raw):
            for node in flatten_json(obj):
                if not isinstance(node, dict):
                    continue
                for key in ["@type", "name", "category", "description", "brand"]:
                    value = node.get(key)
                    if isinstance(value, str):
                        pieces.append(value)
                    elif isinstance(value, list):
                        pieces.extend([clean_text(v) for v in value if isinstance(v, str)])
                    elif isinstance(value, dict):
                        pieces.append(clean_text(value.get("name", "")))
    return clean_text(" | ".join([p for p in pieces if clean_text(p)]))[:4000]


def extract_page_evidence(html: str, base_url: str) -> dict[str, Any]:
    soup = BeautifulSoup(html, "html.parser")
    title = clean_text(soup.title.string if soup.title else "")

    meta_description = ""
    for meta_attrs in [
        {"name": "description"},
        {"property": "og:description"},
        {"name": "twitter:description"},
    ]:
        tag = soup.find("meta", attrs=meta_attrs)
        if tag is not None:
            content = tag.get("content")
            if content:
                meta_description = clean_text(content)
                break

    schema_text = extract_schema_text(soup)

    nav_texts = []
    internal_links = []
    external_links = []
    base_domain = get_domain(base_url)

    for a in soup.find_all("a", href=True):
        text = clean_text(a.get_text(" ", strip=True))
        href = clean_text(a.get("href"))

        if not href:
            continue

        try:
            full_url = urljoin(base_url or "", href)
        except Exception:
            # Some long-tail merchant pages contain malformed href values.
            # Skip only the bad link instead of crashing the full classifier.
            continue

        domain = get_domain(full_url)

        if text and 2 <= len(text) <= 80:
            nav_texts.append(text)

        if is_http_url(full_url):
            if domain and base_domain and domain != base_domain and not is_bad_result_url(full_url):
                external_links.append(full_url)
            if domain == base_domain:
                try:
                    path = urlparse(full_url).path.lower()
                except Exception:
                    continue

                if any(token in path for token in [
                    "shop", "collections", "collection", "category", "categories",
                    "products", "product", "men", "women", "kids", "shoes",
                    "apparel", "beauty", "home", "electronics", "jewelry", "bags",
                    "wallet", "travel", "grill", "pest", "grocery", "audio",
                ]):
                    internal_links.append(full_url)

    for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
        tag.decompose()

    heading_text = " | ".join(
        clean_text(h.get_text(" ", strip=True))
        for h in soup.find_all(["h1", "h2", "h3"], limit=30)
        if clean_text(h.get_text(" ", strip=True))
    )
    body_text = clean_text(soup.get_text(" ", strip=True))[:5000]

    return {
        "title": title,
        "meta_description": meta_description,
        "schema_text": schema_text,
        "nav_text": clean_text(" | ".join(nav_texts[:150]))[:4000],
        "heading_text": clean_text(heading_text)[:2500],
        "body_text": body_text,
        "internal_links": list(dict.fromkeys(internal_links))[:8],
        "external_links": list(dict.fromkeys(external_links))[:20],
    }


def get_extra_internal_evidence(internal_links: list[str], max_pages: int, sleep: float) -> str:
    snippets = []
    bad_parts = [
        "account", "login", "cart", "checkout", "privacy", "terms", "return",
        "shipping", "contact", "help", "support",
    ]

    chosen = []
    for link in internal_links:
        try:
            path = urlparse(str(link)).path.lower()
        except Exception:
            continue

        if any(bad in path for bad in bad_parts):
            continue

        chosen.append(str(link))

        if len(chosen) >= max_pages:
            break

    for link in chosen:
        ok, final_url, html, err = fetch_url(link)

        if sleep:
            time.sleep(sleep + random.uniform(0, 0.2))

        if not ok:
            continue

        try:
            ev = extract_page_evidence(html, final_url)
        except Exception:
            # Internal pages are supplemental evidence only.
            # A malformed internal page should not stop the main merchant classifier.
            continue

        snippets.append(
            clean_text(
                " ".join(
                    [
                        ev.get("title", ""),
                        ev.get("meta_description", ""),
                        ev.get("schema_text", ""),
                        ev.get("nav_text", ""),
                        ev.get("heading_text", ""),
                        ev.get("body_text", "")[:1200],
                    ]
                )
            )
        )

    return clean_text(" ".join(snippets))[:5000]


# ============================================================
# Official URL discovery
# ============================================================

def discover_from_existing_urls(row: pd.Series, merchant_name: str, sleep: float) -> tuple[str, str, str, str]:
    urls = get_existing_urls(row)
    for url in urls:
        redirect = extract_redirect_dest(url)
        if redirect and not is_bad_result_url(redirect):
            return redirect, "existing_redirect_dest_url", "", ""
        if is_bad_result_url(url):
            continue
        if not is_bnpl_url(url):
            return url, "existing_non_bnpl_url", "", ""

        ok, final_url, html, err = fetch_url(url)
        if sleep:
            time.sleep(sleep)
        if not ok:
            continue
        try:
            ev = extract_page_evidence(html, final_url)
        except Exception:
            continue

        external_links = ev.get("external_links", [])
        if external_links:
            ranked = sorted(
                external_links,
                key=lambda u: domain_name_match_score(merchant_name, u),
                reverse=True,
            )
            return ranked[0], "external_link_from_bnpl_page", "", ""

    return "", "", "", "no usable URL from existing columns"


def discover_by_search(merchant_name: str, country: str, max_results: int, sleep: float) -> tuple[str, str, str, str]:
    country_hint = ""
    if "united kingdom" in country.lower():
        country_hint = " UK"
    elif "united states" in country.lower():
        country_hint = " US"

    search_name = make_search_friendly_name(merchant_name) or merchant_name
    query = f'"{search_name}" official website store{country_hint}'
    results = web_search(query, max_results=max_results)
    if sleep:
        time.sleep(sleep)

    official_url, selected, note = choose_official_search_result(results, search_name)
    search_snippet = clean_text(" ".join([
        clean_text(selected.get("title")),
        clean_text(selected.get("body")),
        clean_text(selected.get("href")),
    ]))
    if not official_url:
        return "", search_snippet, "search_engine_failed", note
    return official_url, search_snippet, "search_engine_result", note


# ============================================================
# Classification
# ============================================================

def phrase_count(text: str, phrase: str) -> int:
    text_l = text.lower()
    phrase_l = phrase.lower().strip()
    if not phrase_l:
        return 0
    if " " in phrase_l or "-" in phrase_l or "'" in phrase_l or "&" in phrase_l:
        return text_l.count(phrase_l)
    return len(re.findall(rf"(?<![a-z0-9]){re.escape(phrase_l)}(?![a-z0-9])", text_l))


def score_evidence(evidence_sources: dict[str, str]) -> dict[tuple[str, str], int]:
    source_weights = {
        "schema": 8,
        "title_meta": 5,
        "nav": 5,
        "heading": 4,
        "extra": 4,
        "body": 1,
        "search_snippet": 3,
    }
    scores: dict[tuple[str, str], int] = defaultdict(int)
    for broad, sub_map in TAXONOMY_TERMS.items():
        for sub, terms in sub_map.items():
            for source_name, source_text in evidence_sources.items():
                weight = source_weights.get(source_name, 1)
                for term in terms:
                    matches = phrase_count(source_text, term)
                    if matches:
                        term_bonus = 3 if " " in term else 1
                        scores[(broad, sub)] += matches * weight * term_bonus
    return scores


def apply_product_first_guardrails(scores: dict[tuple[str, str], int], combined_evidence: str) -> None:
    t = combined_evidence.lower()

    def contains_any(terms: list[str]) -> bool:
        return any(term in t for term in terms)

    # Sixth-round calibration guardrails from second 100-row validation sample.
    # These are targeted corrections for high-confidence false positives remaining after v5.
    if contains_any(["le labo", "lelabofragrances", "le labo fragrances", "fragrances", "perfumes", "eau de parfum"]):
        scores[("Beauty & Personal Care", "Fragrance")] += 140
        scores[("Gifts & Flowers", "Photo printing & personalized gifts")] -= 120
        scores[("Gifts & Flowers", "Gifts")] -= 80

    if contains_any(["lego", "lego sets", "lego bricks", "building toys", "toy bricks"]):
        scores[("Baby & Kids", "Toys")] += 140
        scores[("Office & Business Supplies", "Office supplies")] -= 120
        scores[("Marketplace / General Retail", "Big-box / general merchandise")] -= 30

    if contains_any(["booking.com", "booking com", "booking holdings", "book hotels", "hotel booking", "accommodations", "lodging"]):
        scores[("Travel & Ticketing", "Travel booking / lodging marketplace")] += 140
        scores[("Services / Insurance", "Rideshare / local transportation")] -= 130

    if contains_any(["google store", "googlestore", "pixel phone", "pixel phones", "google pixel", "nest", "fitbit", "google electronics"]):
        scores[("Electronics", "Consumer electronics")] += 140
        scores[("Apparel & Fashion", "Jewelry & watches")] -= 130

    if contains_any(["stubhub", "event tickets", "concert tickets", "sports tickets", "ticket marketplace", "buy tickets", "sell tickets"]):
        scores[("Travel & Ticketing", "Event tickets")] += 140
        scores[("Sports & Outdoor", "Outdoor apparel & gear")] -= 130

    if contains_any(["shop.app", "shop app", "shopify shop app", "package tracking", "shopping app"]):
        scores[("Marketplace / General Retail", "Shopping platform / app")] += 110
        scores[("Marketplace / General Retail", "Marketplace")] += 40
        scores[("Services / Insurance", "Software service")] -= 80

    if contains_any(["tcgplayer", "trading cards", "collectible cards", "card marketplace", "pokemon cards", "magic the gathering"]):
        scores[("Marketplace / General Retail", "Collectibles / trading cards marketplace")] += 120
        scores[("Entertainment / Gaming", "Collectibles / comics")] += 30

    if contains_any(["publix", "grocery supermarket", "supermarket", "publix supermarket"]):
        scores[("Food & Grocery", "Grocery")] += 120
        scores[("Food & Grocery", "Grocery delivery")] -= 90

    if contains_any(["quince", "cashmere", "apparel", "clothing", "silk", "linen"]):
        scores[("Apparel & Fashion", "Fashion / apparel")] += 80
        scores[("Apparel & Fashion", "Jewelry & watches")] -= 60

    if contains_any(["sephora", "beauty retailer", "makeup", "cosmetics", "skincare", "beauty products"]):
        scores[("Beauty & Personal Care", "Cosmetics / makeup")] += 95
        # Sephora sells fragrance, but the merchant-level category is better treated as beauty retail.
        scores[("Beauty & Personal Care", "Fragrance")] -= 35

    if contains_any(["l.l.bean", "ll bean", "outdoor clothing", "outdoor gear", "boots", "hiking", "camping"]):
        scores[("Sports & Outdoor", "Outdoor apparel & gear")] += 100
        scores[("Apparel & Fashion", "Fashion / apparel")] -= 35
        scores[("Apparel & Fashion", "Bags & accessories")] -= 50

    if contains_any(["new balance", "newbalance", "nike", "nike.com", "nobull", "no bull", "athletic footwear", "running shoes", "training shoes"]):
        scores[("Sports & Outdoor", "Athletic footwear & apparel")] += 95
        scores[("Apparel & Fashion", "Fashion / apparel")] -= 25
        scores[("Apparel & Fashion", "Footwear")] -= 10

    # Fifth-round calibration guardrails from 100-row validation sample.
    # These correct high-confidence false positives caused by broad terms like bags, rings, books, or marketplace.
    if contains_any(["le labo", "lelabofragrances", "fine fragrances", "perfume", "cologne", "eau de parfum"]):
        scores[("Beauty & Personal Care", "Fragrance")] += 85
        scores[("Gifts & Flowers", "Gifts")] -= 45

    if contains_any(["lyft", "rideshare", "ride share", "ride-hailing", "ride hailing", "taxi", "book a ride"]):
        scores[("Services / Insurance", "Rideshare / local transportation")] += 90
        scores[("Services / Insurance", "Insurance")] -= 70

    if contains_any(["nespresso", "coffee capsules", "coffee pods", "espresso machine", "espresso machines"]):
        scores[("Food & Grocery", "Coffee / tea")] += 90
        scores[("Apparel & Fashion", "Jewelry & watches")] -= 90
        scores[("Gifts & Flowers", "Gifts")] -= 30

    if contains_any(["office depot", "officedepot", "office supplies", "printer ink", "toner", "business supplies"]):
        scores[("Office & Business Supplies", "Office supplies")] += 95
        scores[("Apparel & Fashion", "Bags & accessories")] -= 70

    if contains_any(["oura ring", "smart ring", "sleep tracking", "health tracking", "readiness score"]):
        scores[("Electronics", "Wearable health tech")] += 95
        scores[("Apparel & Fashion", "Jewelry & watches")] -= 85

    if contains_any(["popsockets", "popgrip", "phone grip", "phone stand", "mobile accessories"]):
        scores[("Electronics", "Phone accessories")] += 90
        scores[("Apparel & Fashion", "Bags & accessories")] -= 75

    if contains_any(["shutterfly", "photo books", "photo printing", "photo cards", "personalized gifts", "custom cards"]):
        scores[("Gifts & Flowers", "Photo printing & personalized gifts")] += 95
        scores[("Education / Books", "Books / textbooks")] -= 70

    if contains_any(["six flags", "theme park", "amusement park", "roller coasters", "park tickets", "season pass"]):
        scores[("Travel & Ticketing", "Theme parks / attractions")] += 95
        scores[("Apparel & Fashion", "Bags & accessories")] -= 80

    if contains_any(["sur la table", "williams sonoma", "cookware", "kitchenware", "kitchen tools", "dutch oven", "cutlery"]):
        scores[("Home & Furniture", "Cookware / kitchenware")] += 90
        scores[("Apparel & Fashion", "Bags & accessories")] -= 60
        scores[("Home & Furniture", "Furniture")] -= 25

    if contains_any(["tps garage", "garage llc", "performance parts", "automotive parts", "truck parts", "car parts"]):
        scores[("Automotive", "Auto parts / garage")] += 80
        scores[("Apparel & Fashion", "Jewelry & watches")] -= 80

    if contains_any(["the pro's closet", "the pros closet", "used bikes", "bike marketplace", "bicycle", "cycling"]):
        scores[("Sports & Outdoor", "Cycling / bike resale")] += 90
        scores[("Automotive", "Tires / wheels")] -= 80

    if contains_any(["zara", "zara official", "women's clothing", "men's clothing", "zara clothing"]):
        scores[("Apparel & Fashion", "Fashion / apparel")] += 100
        scores[("Electronics", "Audio electronics")] -= 100

    if contains_any(["doordash", "restaurant delivery", "food delivery", "order food", "dashpass"]):
        scores[("Food & Grocery", "Restaurant / food delivery")] += 100
        scores[("Apparel & Fashion", "Bags & accessories")] -= 90

    if contains_any(["t-mobile", "tmobile", "wireless carrier", "phone plans", "cell phone plans", "mobile network"]):
        scores[("Telecom", "Wireless carrier")] += 100
        scores[("Marketplace / General Retail", "Big-box / general merchandise")] -= 80

    if contains_any(["l.l.bean", "ll bean", "rei", "patagonia", "outdoor clothing", "outdoor gear", "hiking gear"]):
        scores[("Sports & Outdoor", "Outdoor apparel & gear")] += 65
        scores[("Apparel & Fashion", "Bags & accessories")] -= 40

    if contains_any(["new balance", "newbalance", "nike", "nike.com", "athletic shoes", "running shoes"]):
        scores[("Sports & Outdoor", "Athletic footwear & apparel")] += 65
        scores[("Apparel & Fashion", "Fashion / apparel")] -= 15

    if contains_any(["rooms to go", "roomstogo", "furniture store", "living room furniture", "bedroom furniture"]):
        scores[("Home & Furniture", "Furniture")] += 85
        scores[("Home & Furniture", "Home decor")] -= 35

    if contains_any(["sephora", "beauty products", "makeup", "cosmetics", "skincare"]):
        scores[("Beauty & Personal Care", "Cosmetics / makeup")] += 55

    # Third-round calibration guardrails from validation sample.
    if contains_any(["best buy", "bestbuy"]):
        scores[("Electronics", "Consumer electronics")] += 80
        scores[("Marketplace / General Retail", "Marketplace")] -= 45
        scores[("Marketplace / General Retail", "Department store / omnichannel retail")] -= 35

    if contains_any(["farfetch", "luxury fashion marketplace", "designer fashion", "luxury fashion", "boutiques worldwide"]):
        scores[("Apparel & Fashion", "Designer fashion & accessories")] += 75
        scores[("Marketplace / General Retail", "Department store / omnichannel retail")] -= 45

    if contains_any(["champs sports", "champs", "foot locker family", "sports footwear", "athletic footwear"]):
        scores[("Sports & Outdoor", "Athletic footwear & apparel")] += 60
        scores[("Apparel & Fashion", "Footwear")] -= 15

    if contains_any(["holabird sports", "holabird", "running shoes", "tennis racquets", "racquets", "running gear"]):
        scores[("Sports & Outdoor", "Athletic footwear & apparel")] += 55
        scores[("Apparel & Fashion", "Footwear")] -= 15

    if contains_any(["b&h photo", "bh photo", "photo video", "camera equipment", "photography equipment"]):
        scores[("Electronics", "Consumer electronics")] += 65

    if contains_any(["bombas", "socks", "underwear", "t-shirts", "basics"]):
        scores[("Apparel & Fashion", "Fashion / apparel")] += 45

    if contains_any(["leggings", "sports bra", "sports bras", "activewear", "athleisure", "yoga pants", "workout clothes", "running shorts"]):
        scores[("Apparel & Fashion", "Activewear / athleisure")] += 35
        scores[("Sports & Outdoor", "Athletic footwear & apparel")] += 20

    if sum(1 for term in ["iphone", "ipad", "macbook", "imac", "apple watch", "airpods"] if term in t) >= 2:
        scores[("Electronics", "Consumer electronics")] += 50

    if contains_any(["running shoes", "athletic shoes", "tennis shoes", "trail running shoes", "sports apparel", "sportswear"]):
        scores[("Sports & Outdoor", "Athletic footwear & apparel")] += 40

    if contains_any(["headphones", "planar magnetic", "audiophile", "wireless headphones", "speaker", "speakers", "dac", "amplifier"]):
        scores[("Electronics", "Audio electronics")] += 55
        # Prevent audio brands from being pulled into gaming only because a product is game-compatible.
        scores[("Entertainment / Gaming", "Video games / consoles")] -= 25

    if contains_any(["security camera", "security cameras", "robot vacuum", "video doorbell", "smart home", "eufy", "smart lock"]):
        scores[("Electronics", "Smart home electronics")] += 55

    if contains_any(["wallets", "wallet", "bags", "backpack", "pouches", "phone cases", "card holder", "bellroy"]):
        scores[("Apparel & Fashion", "Bags & accessories")] += 55
        scores[("Travel & Ticketing", "Travel booking / lodging marketplace")] -= 30

    if contains_any(["jewelry", "bracelets", "rings", "necklaces", "earrings", "david yurman"]):
        scores[("Apparel & Fashion", "Jewelry & watches")] += 60
        scores[("Gifts & Flowers", "Gifts")] -= 35

    if contains_any(["cross-border", "cross border", "b2b marketplace", "b2c marketplace", "global wholesale", "dhgate"]):
        scores[("Marketplace / General Retail", "Cross-border marketplace")] += 70

    if contains_any(["vacuum cleaner", "vacuum cleaners", "air purifier", "air purifiers", "dyson"]):
        scores[("Home & Furniture", "Home appliances")] += 60
        scores[("Beauty & Personal Care", "Haircare")] -= 20

    if contains_any(["grocery delivery", "grocery pickup", "same-day grocery", "personal shopper", "instacart"]):
        scores[("Food & Grocery", "Grocery delivery")] += 65

    if contains_any(["grill", "grills", "bbq", "barbecue", "outdoor kitchen", "bbqguys"]):
        scores[("Home & Furniture", "Outdoor living & grills")] += 60

    if contains_any(["pest control", "insecticide", "pesticide", "weed killer", "do my own pest control"]):
        scores[("Home & Furniture", "Pest control & lawn care")] += 60
        scores[("Pet", "Pet healthcare")] -= 25

    if contains_any(["hanna andersson", "kids pajamas", "children's pajamas", "kids clothing"]):
        scores[("Baby & Kids", "Kids clothing")] += 60

    if contains_any(["flight booking", "book flights", "expedia flights", "compare flights", "airline tickets"]):
        scores[("Travel & Ticketing", "Flight booking")] += 45

    if contains_any(["airline", "delta air lines", "jetblue", "book a flight"]):
        scores[("Travel & Ticketing", "Airline")] += 45

    if contains_any(["vinyl", "albums", "cds", "cassettes", "artist merchandise", "band merchandise", "tour merch"]):
        scores[("Entertainment / Music", "Music artist merchandise / fan shop")] += 35
        scores[("Entertainment / Music", "Records / vinyl / music media")] += 20

    if contains_any(["department store", "department stores", "omnichannel retailer", "clothing, shoes, home", "home, bedding, furniture"]):
        scores[("Marketplace / General Retail", "Department store / omnichannel retail")] += 50


def classify_from_evidence(
    evidence_sources: dict[str, str],
    website_fetched: bool,
    used_search_snippet_only: bool,
) -> tuple[str, str, str, str, bool]:
    combined = clean_text(" ".join(evidence_sources.values()))[:10000]
    scores = score_evidence(evidence_sources)
    apply_product_first_guardrails(scores, combined)

    if not scores:
        return ("Other / Unclassified", "Needs manual review", "low", "No taxonomy terms matched the website/search evidence.", True)

    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    (top_broad, top_sub), top_score = ranked[0]
    (second_broad, second_sub), second_score = ranked[1] if len(ranked) > 1 else (("None", "None"), 0)

    if top_score <= 0:
        return ("Other / Unclassified", "Needs manual review", "low", "No positive taxonomy score.", True)

    margin = top_score - second_score
    has_structured_or_nav = bool(
        clean_text(evidence_sources.get("schema"))
        or clean_text(evidence_sources.get("nav"))
        or clean_text(evidence_sources.get("heading"))
    )

    if used_search_snippet_only:
        if top_score >= 18 and margin >= 5:
            confidence = "medium"
        else:
            confidence = "low"
    elif website_fetched and has_structured_or_nav and top_score >= 35 and margin >= 8:
        confidence = "high"
    elif website_fetched and top_score >= 16 and margin >= 4:
        confidence = "medium"
    else:
        confidence = "low"

    if second_score > 0 and top_broad != second_broad and margin < 6:
        confidence = "low"

    needs_review = confidence == "low" or top_broad == "Other / Unclassified"
    reason = (
        f"Classified from website/search evidence only. "
        f"Top={top_broad} / {top_sub} score={top_score}; "
        f"Second={second_broad} / {second_sub} score={second_score}; "
        f"Margin={margin}; website_fetched={website_fetched}; "
        f"search_snippet_only={used_search_snippet_only}; "
        f"structured_or_nav_evidence={has_structured_or_nav}."
    )
    return top_broad, top_sub, confidence, reason, needs_review


# ============================================================
# Exact-brand overrides
# ============================================================

EXACT_BRAND_OVERRIDES: dict[str, tuple[str, str, str]] = {
    # Regression fixes from the 100-row v6 validation sample.
    "mercari": (
        "Marketplace / General Retail",
        "Marketplace",
        "Mercari is a marketplace; prevent apparel terms from pulling it into Fashion.",
    ),
    "saksfifthavenue": (
        "Marketplace / General Retail",
        "Department store / omnichannel retail",
        "Saks Fifth Avenue is best treated as a department-store / designer retail merchant.",
    ),
    "saksoff5th": (
        "Marketplace / General Retail",
        "Department store / omnichannel retail",
        "Saks OFF 5TH is best treated as an off-price department-store / designer retail merchant.",
    ),
    "sephora": (
        "Beauty & Personal Care",
        "Cosmetics / makeup",
        "Sephora is a beauty retailer; fragrance is only one product category.",
    ),
    "pureology": (
        "Beauty & Personal Care",
        "Haircare",
        "Pureology is a haircare brand; prevent generic beauty terms from pulling it into cosmetics.",
    ),
    "therollingstonesofficialstore": (
        "Entertainment / Music",
        "Music artist merchandise / fan shop",
        "Official artist stores are music merchandise / fan shops, even if they sell apparel.",
    ),
    "zales": (
        "Apparel & Fashion",
        "Jewelry & watches",
        "Zales is a jewelry retailer; prevent electronics/product-card noise from dominating.",
    ),
    "googlestore": (
        "Electronics",
        "Consumer electronics",
        "Google Store sells Google hardware / consumer electronics; prevent Pixel/Nest/Toy noise from dominating.",
    ),
    "shopapp": (
        "Marketplace / General Retail",
        "Shopping platform / app",
        "Shop.app is a shopping / package-tracking platform app, not a product vertical retailer.",
    ),
    "stubhub": (
        "Travel & Ticketing",
        "Event tickets",
        "StubHub is an event-ticket marketplace; prevent sports terms from pulling it into sporting goods.",
    ),
    "bookingcom": (
        "Travel & Ticketing",
        "Travel booking / lodging marketplace",
        "Booking.com is a travel/lodging booking platform; prevent rideshare/local-transport terms from dominating.",
    ),
    "thelego": (
        "Baby & Kids",
        "Toys",
        "LEGO is a toy/building-set brand; prevent generic office/product terms from dominating.",
    ),
    "lego": (
        "Baby & Kids",
        "Toys",
        "LEGO is a toy/building-set brand; prevent generic office/product terms from dominating.",
    ),
    # Sports / outdoor calibration.
    "llbean": (
        "Sports & Outdoor",
        "Outdoor apparel & gear",
        "L.L.Bean is best treated as outdoor apparel/gear rather than generic fashion.",
    ),
    "rei": (
        "Sports & Outdoor",
        "Outdoor apparel & gear",
        "REI is best treated as outdoor gear/apparel retail.",
    ),
    "patagonia": (
        "Sports & Outdoor",
        "Outdoor apparel & gear",
        "Patagonia is best treated as outdoor apparel/gear.",
    ),
    "newbalancerunning": (
        "Sports & Outdoor",
        "Athletic footwear & apparel",
        "New Balance is best treated as athletic footwear/apparel.",
    ),
    "newbalance": (
        "Sports & Outdoor",
        "Athletic footwear & apparel",
        "New Balance is best treated as athletic footwear/apparel.",
    ),
    "nike": (
        "Sports & Outdoor",
        "Athletic footwear & apparel",
        "Nike is best treated as athletic footwear/apparel for merchant-industry analysis.",
    ),
    "nobull": (
        "Sports & Outdoor",
        "Athletic footwear & apparel",
        "NOBULL is best treated as athletic footwear/apparel.",
    ),
    "lululemon": (
        "Sports & Outdoor",
        "Athletic footwear & apparel",
        "lululemon is treated as athletic/activewear for merchant-industry analysis.",
    ),
    "vans": (
        "Apparel & Fashion",
        "Footwear",
        "Vans is best treated as footwear rather than generic fashion.",
    ),
    # Retailers with recurring subindustry drift.
    "macys": (
        "Marketplace / General Retail",
        "Department store / omnichannel retail",
        "Macy's is a department-store retailer.",
    ),
    "lowes": (
        "Home & Furniture",
        "Home improvement / hardware",
        "Lowe's is a home-improvement / hardware retailer.",
    ),
    "loweshomeimprovement": (
        "Home & Furniture",
        "Home improvement / hardware",
        "Lowe's is a home-improvement / hardware retailer.",
    ),
    "publix": (
        "Food & Grocery",
        "Grocery",
        "Publix is a grocery/supermarket retailer, not primarily grocery delivery.",
    ),
    "tcgplayer": (
        "Marketplace / General Retail",
        "Collectibles / trading cards marketplace",
        "TCGplayer is a trading-card / collectibles marketplace.",
    ),
    "lelabo": (
        "Beauty & Personal Care",
        "Fragrance",
        "Le Labo is a fragrance/perfume brand.",
    ),
    "lelabofragrances": (
        "Beauty & Personal Care",
        "Fragrance",
        "Le Labo is a fragrance/perfume brand.",
    ),
}


def exact_brand_override_key(row: pd.Series, merchant_name: str) -> str:
    """
    Find an exact-brand override key from row-level merchant names.
    This is intentionally exact/conservative and only used for recurring validation errors.
    """
    candidates = set(get_name_candidates(row))
    candidates.add(normalize_for_filter(merchant_name))
    friendly = make_search_friendly_name(merchant_name)
    if friendly:
        candidates.add(normalize_for_filter(friendly))

    for candidate in candidates:
        if candidate in EXACT_BRAND_OVERRIDES:
            return candidate
    return ""


def apply_exact_brand_override(
    row: pd.Series,
    merchant_name: str,
    broad: str,
    sub: str,
    confidence: str,
    reason: str,
    needs_review: bool,
) -> tuple[str, str, str, str, bool]:
    key = exact_brand_override_key(row, merchant_name)
    if not key:
        return broad, sub, confidence, reason, needs_review

    override_broad, override_sub, override_reason = EXACT_BRAND_OVERRIDES[key]

    if broad == override_broad and sub == override_sub and confidence == "high" and not needs_review:
        return broad, sub, confidence, reason, needs_review

    updated_reason = (
        f"Exact-brand override applied for key={key}: {override_reason} "
        f"Original classifier output was {broad} / {sub} with confidence={confidence}. "
        f"Original reason: {reason}"
    )

    return override_broad, override_sub, "high", updated_reason, False


# ============================================================
# One merchant classification
# ============================================================

def classify_one_merchant(row: pd.Series, sleep: float, search_results: int, extra_pages: int) -> ClassificationResult:
    cache_key = merchant_cache_key(row)
    merchant_name = get_merchant_name(row)
    country = clean_text(row.get("country", ""))

    is_internal_page, internal_reason = is_bnpl_provider_internal_page(row)
    if is_internal_page:
        return make_exclusion_result(
            row,
            broad="Exclude / BNPL Internal Page",
            sub="Provider financial product / internal page",
            status="exclude_provider_internal_page",
            reason=internal_reason,
        )

    is_nav_page, nav_reason = is_non_merchant_navigation_page(row)
    if is_nav_page:
        return make_exclusion_result(
            row,
            broad="Exclude / Non-Merchant Page",
            sub="Navigation / locale / corporate informational page",
            status="exclude_non_merchant_page",
            reason=nav_reason,
        )

    if not merchant_name:
        return ClassificationResult(
            cache_key=cache_key,
            merchant_name="",
            official_url="",
            final_url="",
            url_discovery_method="missing_merchant_name",
            evidence_title="",
            evidence_meta_description="",
            evidence_schema_text="",
            evidence_nav_text="",
            evidence_heading_text="",
            evidence_body_snippet="",
            evidence_search_snippet="",
            verified_broad_industry="Other / Unclassified",
            verified_sub_industry="Needs manual review",
            classification_confidence="low",
            verification_status="unable_to_verify",
            classification_reason="Missing merchant name.",
            needs_manual_review=True,
            error_message="missing merchant name",
        )

    official_url, method, search_snippet, err = discover_from_existing_urls(row=row, merchant_name=merchant_name, sleep=sleep)

    if not official_url:
        official_url, search_snippet, method, err = discover_by_search(
            merchant_name=merchant_name,
            country=country,
            max_results=search_results,
            sleep=sleep,
        )

    if not official_url:
        richer_search_evidence = build_search_evidence_bundle(
            merchant_name=merchant_name,
            country=country,
            max_results=search_results,
        )
        combined_search_evidence = clean_text(f"{search_snippet} {richer_search_evidence}")
        evidence_sources = {"search_snippet": combined_search_evidence}
        broad, sub, confidence, reason, needs_review = classify_from_evidence(
            evidence_sources=evidence_sources,
            website_fetched=False,
            used_search_snippet_only=True,
        )
        broad, sub, confidence, reason, needs_review = apply_exact_brand_override(
            row, merchant_name, broad, sub, confidence, reason, needs_review
        )
        return ClassificationResult(
            cache_key=cache_key,
            merchant_name=merchant_name,
            official_url="",
            final_url="",
            url_discovery_method=method,
            evidence_title="",
            evidence_meta_description="",
            evidence_schema_text="",
            evidence_nav_text="",
            evidence_heading_text="",
            evidence_body_snippet="",
            evidence_search_snippet=combined_search_evidence,
            verified_broad_industry=broad,
            verified_sub_industry=sub,
            classification_confidence=confidence,
            verification_status="search_only_no_official_url",
            classification_reason=reason,
            needs_manual_review=needs_review,
            error_message=err,
        )

    ok, final_url, html, fetch_err = fetch_url(official_url)
    if sleep:
        time.sleep(sleep + random.uniform(0, 0.2))

    if not ok:
        richer_search_evidence = build_search_evidence_bundle(
            merchant_name=merchant_name,
            country=country,
            max_results=search_results,
        )
        combined_search_evidence = clean_text(f"{search_snippet} {richer_search_evidence}")
        evidence_sources = {"search_snippet": combined_search_evidence}
        broad, sub, confidence, reason, needs_review = classify_from_evidence(
            evidence_sources=evidence_sources,
            website_fetched=False,
            used_search_snippet_only=True,
        )
        broad, sub, confidence, reason, needs_review = apply_exact_brand_override(
            row, merchant_name, broad, sub, confidence, reason, needs_review
        )
        return ClassificationResult(
            cache_key=cache_key,
            merchant_name=merchant_name,
            official_url=official_url,
            final_url=official_url,
            url_discovery_method=method,
            evidence_title="",
            evidence_meta_description="",
            evidence_schema_text="",
            evidence_nav_text="",
            evidence_heading_text="",
            evidence_body_snippet="",
            evidence_search_snippet=combined_search_evidence,
            verified_broad_industry=broad,
            verified_sub_industry=sub,
            classification_confidence=confidence,
            verification_status="search_verified_website_unreachable",
            classification_reason=reason,
            needs_manual_review=needs_review,
            error_message=fetch_err,
        )

    try:
        ev = extract_page_evidence(html, final_url)
    except Exception as exc:
        richer_search_evidence = build_search_evidence_bundle(
            merchant_name=merchant_name,
            country=country,
            max_results=search_results,
        )

        combined_search_evidence = clean_text(
            f"{search_snippet} {richer_search_evidence}"
        )

        evidence_sources = {
            "search_snippet": combined_search_evidence,
        }

        broad, sub, confidence, reason, needs_review = classify_from_evidence(
            evidence_sources=evidence_sources,
            website_fetched=False,
            used_search_snippet_only=True,
        )

        broad, sub, confidence, reason, needs_review = apply_exact_brand_override(
            row, merchant_name, broad, sub, confidence, reason, needs_review
        )

        return ClassificationResult(
            cache_key=cache_key,
            merchant_name=merchant_name,
            official_url=official_url,
            final_url=final_url,
            url_discovery_method=method,
            evidence_title="",
            evidence_meta_description="",
            evidence_schema_text="",
            evidence_nav_text="",
            evidence_heading_text="",
            evidence_body_snippet="",
            evidence_search_snippet=combined_search_evidence,
            verified_broad_industry=broad,
            verified_sub_industry=sub,
            classification_confidence=confidence,
            verification_status="website_parse_failed_search_fallback",
            classification_reason=f"{reason} Website parsing failed: {exc}",
            needs_manual_review=True,
            error_message=f"website parse error: {exc}",
        )

    extra_evidence = get_extra_internal_evidence(
        internal_links=ev.get("internal_links", []),
        max_pages=extra_pages,
        sleep=sleep,
    )
    evidence_sources = {
        "schema": ev.get("schema_text", ""),
        "title_meta": clean_text(f"{ev.get('title', '')} {ev.get('meta_description', '')}"),
        "nav": ev.get("nav_text", ""),
        "heading": ev.get("heading_text", ""),
        "body": ev.get("body_text", "")[:2500],
        "extra": extra_evidence,
        "search_snippet": search_snippet,
    }
    broad, sub, confidence, reason, needs_review = classify_from_evidence(
        evidence_sources=evidence_sources,
        website_fetched=True,
        used_search_snippet_only=False,
    )
    broad, sub, confidence, reason, needs_review = apply_exact_brand_override(
        row, merchant_name, broad, sub, confidence, reason, needs_review
    )

    # If the official page was technically fetched but produced weak/no usable
    # product evidence, run a richer search-evidence fallback. This handles
    # JavaScript-heavy pages where requests.get() returns a shell without
    # meaningful product/category text.
    used_post_fetch_search_fallback = False
    if broad == "Other / Unclassified" or confidence == "low":
        richer_search_evidence = build_search_evidence_bundle(
            merchant_name=merchant_name,
            country=country,
            max_results=search_results,
        )
        combined_search_evidence = clean_text(f"{search_snippet} {richer_search_evidence}")
        fallback_sources = dict(evidence_sources)
        fallback_sources["search_snippet"] = combined_search_evidence

        fb_broad, fb_sub, fb_confidence, fb_reason, fb_needs_review = classify_from_evidence(
            evidence_sources=fallback_sources,
            website_fetched=True,
            used_search_snippet_only=False,
        )

        # Accept fallback only when it improves the classification.
        if fb_broad != "Other / Unclassified" and (
            broad == "Other / Unclassified" or fb_confidence in {"medium", "high"}
        ):
            broad, sub, confidence, reason, needs_review = (
                fb_broad,
                fb_sub,
                fb_confidence,
                fb_reason + " Post-fetch search fallback used because official page evidence was weak.",
                fb_needs_review,
            )
            broad, sub, confidence, reason, needs_review = apply_exact_brand_override(
                row, merchant_name, broad, sub, confidence, reason, needs_review
            )
            search_snippet = combined_search_evidence
            used_post_fetch_search_fallback = True

    status = "official_website_verified" if broad != "Other / Unclassified" else "needs_manual_review"
    if used_post_fetch_search_fallback and broad != "Other / Unclassified":
        status = "official_website_verified_plus_search_fallback"

    return ClassificationResult(
        cache_key=cache_key,
        merchant_name=merchant_name,
        official_url=official_url,
        final_url=final_url,
        url_discovery_method=method,
        evidence_title=ev.get("title", ""),
        evidence_meta_description=ev.get("meta_description", ""),
        evidence_schema_text=ev.get("schema_text", ""),
        evidence_nav_text=ev.get("nav_text", ""),
        evidence_heading_text=ev.get("heading_text", ""),
        evidence_body_snippet=ev.get("body_text", "")[:2500],
        evidence_search_snippet=search_snippet,
        verified_broad_industry=broad,
        verified_sub_industry=sub,
        classification_confidence=confidence,
        verification_status=status,
        classification_reason=reason,
        needs_manual_review=needs_review,
        error_message="",
    )


# ============================================================
# Cache and output
# ============================================================

def load_cache(cache_path: Path) -> dict[str, ClassificationResult]:
    cache: dict[str, ClassificationResult] = {}
    if not cache_path.exists():
        return cache
    with cache_path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                data = json.loads(line)
                result = ClassificationResult(**data)
                cache[result.cache_key] = result
            except Exception:
                continue
    return cache


def append_cache(cache_path: Path, result: ClassificationResult) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with cache_path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(asdict(result), ensure_ascii=False) + "\n")


def write_log(log_path: Path, result: ClassificationResult) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    file_exists = log_path.exists()
    with log_path.open("a", encoding="utf-8", newline="") as f:
        fieldnames = [
            "merchant_name", "official_url", "final_url", "verified_broad_industry",
            "verified_sub_industry", "classification_confidence", "verification_status",
            "needs_manual_review", "error_message",
        ]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerow({field: getattr(result, field) for field in fieldnames})


def read_input_table(input_path: Path) -> tuple[pd.DataFrame, str]:
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    xls = pd.ExcelFile(input_path)
    if not xls.sheet_names:
        raise ValueError(f"No sheets found in Excel file: {input_path}")
    sheet_name: str = str(xls.sheet_names[0])
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    return df, sheet_name


def apply_results_to_dataframe(df: pd.DataFrame, cache: dict[str, ClassificationResult]) -> pd.DataFrame:
    out = df.copy()

    if "broad_industry" in out.columns and "original_broad_industry" not in out.columns:
        out["original_broad_industry"] = out["broad_industry"]
    if "sub_industry" in out.columns and "original_sub_industry" not in out.columns:
        out["original_sub_industry"] = out["sub_industry"]

    new_cols = [
        "official_url", "final_url", "url_discovery_method", "evidence_title",
        "evidence_meta_description", "evidence_schema_text", "evidence_nav_text",
        "evidence_heading_text", "evidence_body_snippet", "evidence_search_snippet",
        "verified_broad_industry", "verified_sub_industry", "classification_confidence",
        "verification_status", "classification_reason", "needs_manual_review",
        "error_message",
    ]
    for col in new_cols:
        if col not in out.columns:
            out[col] = ""

    for idx, row in out.iterrows():
        key = merchant_cache_key(row)
        result = cache.get(key)
        if not result:
            continue
        for col in new_cols:
            value = getattr(result, col)
            if col == "needs_manual_review":
                value = "TRUE" if bool(value) else "FALSE"
            out.at[idx, col] = value

    return out


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []

    def add(metric: str, value: Any) -> None:
        rows.append({"metric": metric, "value": value})

    add("total_rows", len(df))
    for col in ["verified_broad_industry", "classification_confidence", "verification_status", "needs_manual_review"]:
        if col in df.columns:
            counts = df[col].fillna("").astype(str).value_counts(dropna=False)
            for key, value in counts.items():
                add(f"{col}::{key}", int(value))
    return pd.DataFrame(rows)


def save_output(df: pd.DataFrame, output_path: Path) -> None:
    summary = build_summary(df)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Web_Only_Classified_V4", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
    print(f"Saved output workbook: {output_path}")


# ============================================================
# Main
# ============================================================

def main() -> None:
    parser = argparse.ArgumentParser(description="Classify BNPL merchants from official website/search evidence only.")
    parser.add_argument("--input", default=str(DEFAULT_INPUT), help="Input Excel file.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output Excel file.")
    parser.add_argument("--cache", default=str(DEFAULT_CACHE), help="Cache JSONL file.")
    parser.add_argument("--log", default=str(DEFAULT_LOG), help="Run log CSV file.")
    parser.add_argument("--limit", type=int, default=0, help="Process only first N rows for testing. 0 means all rows.")
    parser.add_argument("--start", type=int, default=0, help="Start row index.")
    parser.add_argument("--sleep", type=float, default=0.8, help="Seconds to sleep between requests.")
    parser.add_argument("--search-results", type=int, default=5, help="Number of search results to inspect.")
    parser.add_argument("--extra-pages", type=int, default=2, help="Number of internal category/product pages to fetch.")
    parser.add_argument("--save-every", type=int, default=50, help="Save output after every N newly processed merchants.")
    parser.add_argument("--force", action="store_true", help="Delete existing cache and rerun from scratch.")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    cache_path = Path(args.cache).expanduser().resolve()
    log_path = Path(args.log).expanduser().resolve()

    df, sheet_name = read_input_table(input_path)
    print(f"Loaded input: {input_path}")
    print(f"Sheet: {sheet_name}")
    print(f"Rows: {len(df)}")

    if args.force:
        if cache_path.exists():
            cache_path.unlink()
        if log_path.exists():
            log_path.unlink()
        print("Force mode: old cache/log removed.")

    cache = load_cache(cache_path)
    print(f"Cached merchants loaded: {len(cache)}")

    indices = list(df.index)
    if args.start:
        indices = [i for i in indices if i >= args.start]
    if args.limit and args.limit > 0:
        indices = indices[: args.limit]

    print(f"Rows scheduled this run: {len(indices)}")

    newly_processed = 0
    for n, idx in enumerate(indices, start=1):
        row = df.loc[idx]
        key = merchant_cache_key(row)
        if key in cache:
            if n % 100 == 0:
                print(f"[{n}/{len(indices)}] cached row {idx}")
            continue

        merchant_name = get_merchant_name(row)
        print(f"[{n}/{len(indices)}] Classifying row {idx}: {merchant_name}")

        result = classify_one_merchant(
            row=row,
            sleep=args.sleep,
            search_results=args.search_results,
            extra_pages=args.extra_pages,
        )
        cache[key] = result
        append_cache(cache_path, result)
        write_log(log_path, result)
        newly_processed += 1

        print(
            f"    -> {result.verification_status} | "
            f"{result.classification_confidence} | "
            f"{result.verified_broad_industry} / {result.verified_sub_industry} | "
            f"review={result.needs_manual_review}"
        )

        if newly_processed % args.save_every == 0:
            partial_df = apply_results_to_dataframe(df, cache)
            save_output(partial_df, output_path)
            print(f"Progress saved after {newly_processed} newly processed merchants.")

    final_df = apply_results_to_dataframe(df, cache)
    save_output(final_df, output_path)

    print("\nDone.")
    print(f"Output: {output_path}")
    print(f"Cache: {cache_path}")
    print(f"Log: {log_path}")

    if "verified_broad_industry" in final_df.columns:
        print("\nVerified broad industry counts:")
        print(final_df["verified_broad_industry"].fillna("").astype(str).value_counts().to_string())
    if "classification_confidence" in final_df.columns:
        print("\nConfidence counts:")
        print(final_df["classification_confidence"].fillna("").astype(str).value_counts().to_string())
    if "verification_status" in final_df.columns:
        print("\nVerification status counts:")
        print(final_df["verification_status"].fillna("").astype(str).value_counts().to_string())


if __name__ == "__main__":
    main()
