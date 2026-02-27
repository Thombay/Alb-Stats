from __future__ import annotations

import argparse
import base64
import html as std_html
import io
import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from plotly.offline import get_plotlyjs_version
from dash import Dash, Input, Output, State, ctx, dash_table, dcc, html, no_update

SHEET_NAME = "Exportdaten"
HEADER_ROW_INDEX = 4  # Excel row 5
SEMESTER_PATTERN = re.compile(r"\((SS\s*\d{4}|WS\s*\d{4}(?:/\d{2})?)\)", re.IGNORECASE)
UNKNOWN_SEMESTER = "Ohne Semester"
DATE_TOKEN_PATTERN = re.compile(r"\d{1,2}\.\s*(?:\d{1,2}|[A-Za-zÄÖÜäöü]+)\.?\s*\d{4}", re.IGNORECASE)
UP_BP_STATUSES = ["UP", "BP", "EM"]
BU_FU_STATUSES = ["BU", "FU"]
STATUS_LABELS = {
    "UP": "UP (Urphilister)",
    "BP": "BP (Bandphilister)",
    "EM": "EM (Ehrenmitglied)",
    "BU": "BU (Bursch)",
    "FU": "FU (Fuchs)",
}
ACTIVE_CHARGEN_KEYWORDS = [
    "barwart",
    "bierkassier",
    "kassier",
    "consenior",
    "schriftfuehrer",
    "fuchsmajor",
    "fuchsmajor 2",
    "senior",
    "ov praesident",
    "ov prasident",
]
PHILISTER_CHARGEN_KEYWORDS = [
    "philistersenior",
    "philisterconsenior",
    "philisterschriftfuehrer",
    "philisterkassier",
]
FUNKTIONAER_KEYWORDS = [
    "chefredakteur",
    "chef red",
    "it beauftragter",
    "it bea",
    "verbindungsseelsorger",
    "verb seels",
    "vorsitzender des verbindungsgerichtes",
    "vg vors",
    "standesfuehrer",
    "standesfuhrer",
    "oecvnet",
    "oecv net",
    "archivar",
    "zirkel",
]
MANUAL_CLASS_OPTIONS = [
    ("Aktivenchargen", "aktiven"),
    ("Philisterchargen", "philister"),
    ("Verbandscharge (Aktiven)", "verband_aktiven"),
    ("Verbandscharge (Philister)", "verband_philister"),
    ("Funktionaere (nicht zaehlen)", "funktionaere"),
]
KIND_SORT_OPTIONS = [
    ("Gesamtchargen", "total"),
    ("Aktivenchargen", "active"),
    ("Philisterchargen", "philister"),
    ("Unklare Chargen", "unclear"),
    ("Semester (gewaehlt)", "semester"),
    ("Name A-Z", "name"),
]
OVERRIDES_FILE_NAME = "chargen_class_overrides.json"
CATEGORY_GRAPH_OPTIONS = [
    "Aktivenchargen",
    "Philisterchargen",
    "Verbandschargen (Aktiven)",
    "Verbandschargen (Philister)",
    "Funktionaere",
]
DEFAULT_CATEGORY_SELECTION = ["Aktivenchargen", "Verbandschargen (Aktiven)"]
PERSON_GROUP_OPTIONS = ["Aktive", "Philister"]
INTENSITY_PART_OPTIONS = [
    "Aktiven + Verbandschargen (Aktiven)",
    "Philister + Verbandschargen (Philister)",
]
MANDATORY_SLOT_ORDER = [
    "senior",
    "consenior",
    "schriftfuehrer",
    "fuchsmajor",
    "kassier",
    "barwart",
    "philistersenior",
    "philisterconsenior",
    "philisterschriftfuehrer",
    "philisterkassier",
]
MANDATORY_SLOT_LABELS = {
    "senior": "Senior",
    "consenior": "Consenior",
    "schriftfuehrer": "Schriftfuehrer",
    "fuchsmajor": "Fuchsmajor",
    "kassier": "Kassier",
    "barwart": "Barwart",
    "philistersenior": "Philistersenior",
    "philisterconsenior": "Philisterconsenior",
    "philisterschriftfuehrer": "Philisterschriftfuehrer",
    "philisterkassier": "Philisterkassier",
}


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def repair_mojibake(text: str) -> str:
    if "Ã" not in text and "â" not in text:
        return text
    try:
        return text.encode("latin1").decode("utf-8")
    except Exception:
        return text


def parse_chargen_entries(value: object) -> list[str]:
    text = normalize_text(value)
    if not text:
        return []
    unified = text.replace("\n", " | ")
    parts = [part.strip() for part in unified.split("|")]
    return [part for part in parts if part]


def extract_semester(chargen_entry: str) -> str:
    match = SEMESTER_PATTERN.search(chargen_entry)
    if not match:
        return UNKNOWN_SEMESTER
    semester = re.sub(r"\s+", " ", match.group(1)).upper()
    return semester


def semester_sort_key(semester: str) -> tuple[int, int, str]:
    if semester == UNKNOWN_SEMESTER:
        return (10_000, 9, semester)
    if semester.startswith("SS "):
        return (int(semester.split()[1]), 0, semester)
    if semester.startswith("WS "):
        return (int(semester.split()[1].split("/")[0]), 1, semester)
    return (9_999, 9, semester)


def parse_semester_parts(semester: str) -> tuple[str, int] | None:
    semester = normalize_text(semester).upper()
    if semester.startswith("SS "):
        try:
            return ("SS", int(semester.split()[1]))
        except Exception:  # noqa: BLE001
            return None
    if semester.startswith("WS "):
        try:
            return ("WS", int(semester.split()[1].split("/")[0]))
        except Exception:  # noqa: BLE001
            return None
    return None


def semester_label_from_parts(term: str, year: int) -> str:
    if term == "SS":
        return f"SS {year}"
    return f"WS {year}/{str((year + 1) % 100).zfill(2)}"


def semester_range_labels(start_semester: str, end_semester: str) -> list[str]:
    start_parts = parse_semester_parts(start_semester)
    end_parts = parse_semester_parts(end_semester)
    if start_parts is None or end_parts is None:
        return []

    season_idx = {"SS": 0, "WS": 1}
    inv_idx = {0: "SS", 1: "WS"}
    start_key = (start_parts[1], season_idx[start_parts[0]])
    end_key = (end_parts[1], season_idx[end_parts[0]])
    if start_key > end_key:
        start_key, end_key = end_key, start_key

    year, season = start_key
    out: list[str] = []
    while (year, season) <= end_key:
        out.append(semester_label_from_parts(inv_idx[season], year))
        if season == 0:
            season = 1
        else:
            season = 0
            year += 1
    return out


def normalize_for_match(text: str) -> str:
    value = normalize_text(text).casefold()
    value = (
        value.replace("ä", "ae")
        .replace("ö", "oe")
        .replace("ü", "ue")
        .replace("ß", "ss")
        .replace("ã¤", "ae")
        .replace("ã¶", "oe")
        .replace("ã¼", "ue")
        .replace("ãÿ", "ss")
    )
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def chargen_override_key(chargen_entry: str) -> str:
    text = repair_mojibake(normalize_text(chargen_entry))
    if not text:
        return ""
    if ":" in text:
        text = text.split(":", maxsplit=1)[1].strip()

    def clean_paren(match: re.Match[str]) -> str:
        inner = normalize_for_match(match.group(1))
        drop_keywords = [
            "ss ",
            "ws ",
            "von ",
            "bis ",
            "ab ",
            "heute",
            "jan",
            "feb",
            "maerz",
            "mrz",
            "apr",
            "mai",
            "jun",
            "jul",
            "aug",
            "sep",
            "okt",
            "nov",
            "dez",
        ]
        if any(keyword in inner for keyword in drop_keywords):
            return ""
        if re.search(r"\d{4}", inner):
            return ""
        return f" ({match.group(1).strip()})"

    text = re.sub(r"\(([^)]*)\)", clean_paren, text)
    text = re.sub(r"\s+", " ", text).strip(" -|")
    return text


def canonicalize_override_map(overrides: dict[str, str] | None) -> dict[str, str]:
    if not overrides:
        return {}
    allowed = {value for _, value in MANUAL_CLASS_OPTIONS}
    normalized: dict[str, str] = {}
    for key, value in overrides.items():
        key_text = normalize_text(key)
        value_text = normalize_text(value)
        if not key_text or value_text not in allowed:
            continue
        canon_key = chargen_override_key(key_text) or key_text
        normalized[canon_key] = value_text
    return normalized


def load_persistent_overrides(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(raw, dict):
        return {}
    return canonicalize_override_map(raw)


def save_persistent_overrides(path: Path, overrides: dict[str, str]) -> None:
    cleaned = canonicalize_override_map(overrides)
    path.write_text(json.dumps(cleaned, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


def parse_german_date_token(token: str) -> date | None:
    token = normalize_text(token)
    if not token:
        return None

    numeric = pd.to_datetime(token, dayfirst=True, errors="coerce")
    if pd.notna(numeric):
        return numeric.date()

    month_map = {
        "jan": 1,
        "jaen": 1,
        "jänner": 1,
        "januar": 1,
        "feb": 2,
        "maerz": 3,
        "märz": 3,
        "mrz": 3,
        "apr": 4,
        "mai": 5,
        "jun": 6,
        "juni": 6,
        "jul": 7,
        "juli": 7,
        "aug": 8,
        "sep": 9,
        "sept": 9,
        "september": 9,
        "okt": 10,
        "okto": 10,
        "nov": 11,
        "dez": 12,
        "dezember": 12,
    }
    compact = (
        token.casefold()
        .replace("ä", "ae")
        .replace("ö", "oe")
        .replace("ü", "ue")
        .replace("ß", "ss")
        .replace(".", "")
    )
    compact = re.sub(r"\s+", " ", compact).strip()
    m = re.match(r"^(\d{1,2})\s+([a-z]+)\s+(\d{4})$", compact)
    if not m:
        return None
    day = int(m.group(1))
    month_text = m.group(2)
    year = int(m.group(3))
    month = month_map.get(month_text[:4]) or month_map.get(month_text[:3])
    if not month:
        return None
    try:
        return date(year, month, day)
    except ValueError:
        return None


def semester_label_for_range_date(d: date) -> str:
    if d.month <= 6:
        return f"SS {d.year}"
    return f"WS {d.year}/{str((d.year + 1) % 100).zfill(2)}"


def semester_labels_for_date_range(start_date: date, end_date: date) -> list[str]:
    if end_date < start_date:
        start_date, end_date = end_date, start_date

    labels: list[str] = []
    cursor = date(start_date.year, start_date.month, 1)
    while cursor <= end_date:
        label = semester_label_for_range_date(cursor)
        if not labels or labels[-1] != label:
            labels.append(label)
        if cursor.month == 12:
            cursor = date(cursor.year + 1, 1, 1)
        else:
            cursor = date(cursor.year, cursor.month + 1, 1)
    return labels


def extract_entry_semesters(chargen_entry: str, today: date | None = None) -> list[str]:
    explicit = SEMESTER_PATTERN.findall(chargen_entry)
    if explicit:
        cleaned = [re.sub(r"\s+", " ", value).upper() for value in explicit]
        unique: list[str] = []
        for sem in cleaned:
            if sem not in unique:
                unique.append(sem)
        return unique

    today = today or date.today()
    low = chargen_entry.casefold()
    date_tokens = DATE_TOKEN_PATTERN.findall(chargen_entry)
    parsed_dates = [parse_german_date_token(token) for token in date_tokens]
    parsed_dates = [d for d in parsed_dates if d is not None]

    if "von" in low and "bis" in low and len(parsed_dates) >= 2:
        return semester_labels_for_date_range(parsed_dates[0], parsed_dates[1])
    if "ab" in low and parsed_dates:
        return semester_labels_for_date_range(parsed_dates[0], today)
    if "bis" in low and parsed_dates and "von" not in low:
        return [semester_label_for_range_date(parsed_dates[0])]

    return [UNKNOWN_SEMESTER]


def count_chargen_units(entries: list[str], today: date | None = None) -> int:
    total = 0
    for entry in entries:
        total += len(extract_entry_semesters(entry, today))
    return total


def is_active_chargen_name(chargen_entry: str) -> bool:
    normalized = normalize_for_match(chargen_entry)
    return any(keyword in normalized for keyword in ACTIVE_CHARGEN_KEYWORDS)


def is_philister_chargen_name(chargen_entry: str) -> bool:
    normalized = normalize_for_match(chargen_entry)
    return any(keyword in normalized for keyword in PHILISTER_CHARGEN_KEYWORDS)


def is_funktionaer_entry(chargen_entry: str) -> bool:
    normalized = normalize_for_match(chargen_entry)
    return any(keyword in normalized for keyword in FUNKTIONAER_KEYWORDS)


def filter_chargen_entries(entries: list[str]) -> list[str]:
    return [entry for entry in entries if not is_funktionaer_entry(entry)]


def special_count_group(couleurname: str, entry_key: str) -> str | None:
    name = normalize_for_match(couleurname)
    role = normalize_for_match(entry_key)
    if name == "gaius" and "ov" in role and ("kassier" in role or "vizepraesident" in role or "vizeprasident" in role):
        return "gaius_ov_kassier_vizepraesident"
    return None


def manual_class_meta(manual_class: str | None) -> tuple[str | None, str | None, bool | None]:
    if manual_class == "aktiven":
        return ("Aktivenchargen", "Aktivenchargen", True)
    if manual_class == "philister":
        return ("Philisterchargen", "Philisterchargen", True)
    if manual_class == "verband_aktiven":
        return ("Aktivenchargen", "Verbandschargen (Aktiven)", True)
    if manual_class == "verband_philister":
        return ("Philisterchargen", "Verbandschargen (Philister)", True)
    if manual_class == "funktionaere":
        return ("Funktionaere", "Funktionaere", False)
    return (None, None, None)


def default_entry_meta(chargen_entry: str, semester: str, philistrierung_date: pd.Timestamp | pd.NaT) -> tuple[str, str, bool]:
    if is_funktionaer_entry(chargen_entry):
        return ("Funktionaere", "Funktionaere", False)
    chargen_type = classify_chargen_kind(semester, philistrierung_date, chargen_entry)
    if chargen_type == "Aktivenchargen":
        return (chargen_type, "Aktivenchargen", True)
    if chargen_type == "Philisterchargen":
        return (chargen_type, "Philisterchargen", True)
    return (chargen_type, "Unklare Chargen", True)


def person_group_from_status(status: str) -> str:
    value = normalize_for_match(status)
    if value in {"bu", "fu"}:
        return "Aktive"
    if value in {"up", "bp", "em"}:
        return "Philister"
    return "Philister"


def build_person_chargen_details(semester_df: pd.DataFrame) -> pd.DataFrame:
    if semester_df.empty:
        return pd.DataFrame(columns=["Couleurname", "ChargenDetailsText", "ChargenDetailsHtml"])

    rows: list[dict[str, str]] = []
    for couleurname, group in semester_df.groupby("Couleurname"):
        ordered = sorted(
            group.to_dict("records"),
            key=lambda item: (semester_sort_key(item["Semester"]), normalize_text(item["ChargenEntry"])),
        )
        seen: set[str] = set()
        lines: list[str] = []
        for item in ordered:
            semester = item["Semester"]
            entry = normalize_text(item["ChargenEntry"])
            line = f"{semester}: {entry}" if semester != UNKNOWN_SEMESTER else entry
            if line in seen:
                continue
            seen.add(line)
            lines.append(line)

        max_lines = 20
        if len(lines) > max_lines:
            hidden = len(lines) - max_lines
            lines = lines[:max_lines] + [f"... (+{hidden} weitere)"]

        rows.append(
            {
                "Couleurname": couleurname,
                "ChargenDetailsText": "\n".join(lines) if lines else "Keine Chargen-Details",
                "ChargenDetailsHtml": "<br>".join(lines) if lines else "Keine Chargen-Details",
            }
        )

    return pd.DataFrame(rows)


def build_person_type_chargen_details(semester_df: pd.DataFrame) -> pd.DataFrame:
    if semester_df.empty:
        return pd.DataFrame(columns=["Couleurname", "ChargenTyp", "ChargenTypeDetailsText", "ChargenTypeDetailsHtml"])

    rows: list[dict[str, str]] = []
    for (couleurname, chargen_typ), group in semester_df.groupby(["Couleurname", "ChargenTyp"]):
        ordered = sorted(
            group.to_dict("records"),
            key=lambda item: (semester_sort_key(item["Semester"]), normalize_text(item["ChargenEntry"])),
        )
        seen: set[str] = set()
        lines: list[str] = []
        for item in ordered:
            semester = item["Semester"]
            entry = normalize_text(item["ChargenEntry"])
            line = f"{semester}: {entry}" if semester != UNKNOWN_SEMESTER else entry
            if line in seen:
                continue
            seen.add(line)
            lines.append(line)

        max_lines = 20
        if len(lines) > max_lines:
            hidden = len(lines) - max_lines
            lines = lines[:max_lines] + [f"... (+{hidden} weitere)"]

        rows.append(
            {
                "Couleurname": couleurname,
                "ChargenTyp": chargen_typ,
                "ChargenTypeDetailsText": "\n".join(lines) if lines else "Keine Chargen-Details",
                "ChargenTypeDetailsHtml": "<br>".join(lines) if lines else "Keine Chargen-Details",
            }
        )
    return pd.DataFrame(rows)


def build_person_category_details(semester_df: pd.DataFrame) -> pd.DataFrame:
    if semester_df.empty:
        return pd.DataFrame(columns=["Couleurname", "ChargenCategory", "CategoryDetailsText", "CategoryDetailsHtml"])

    rows: list[dict[str, str]] = []
    for (couleurname, category), group in semester_df.groupby(["Couleurname", "ChargenCategory"]):
        ordered = sorted(
            group.to_dict("records"),
            key=lambda item: (semester_sort_key(item["Semester"]), normalize_text(item["ChargenEntry"])),
        )
        seen: set[str] = set()
        lines: list[str] = []
        for item in ordered:
            semester = item["Semester"]
            entry = normalize_text(item["ChargenEntry"])
            line = f"{semester}: {entry}" if semester != UNKNOWN_SEMESTER else entry
            if line in seen:
                continue
            seen.add(line)
            lines.append(line)

        max_lines = 20
        if len(lines) > max_lines:
            hidden = len(lines) - max_lines
            lines = lines[:max_lines] + [f"... (+{hidden} weitere)"]

        rows.append(
            {
                "Couleurname": couleurname,
                "ChargenCategory": category,
                "CategoryDetailsText": "\n".join(lines) if lines else "Keine Chargen-Details",
                "CategoryDetailsHtml": "<br>".join(lines) if lines else "Keine Chargen-Details",
            }
        )
    return pd.DataFrame(rows)


def role_name_from_entry(chargen_entry: str) -> str:
    key = chargen_override_key(chargen_entry)
    return key if key else normalize_text(chargen_entry)


def mandatory_slot_for_role(role_name: str) -> str | None:
    norm = normalize_for_match(role_name)
    if not norm:
        return None
    tokens = norm.split()
    token_set = set(tokens)

    def has_prefix(prefix: str) -> bool:
        return any(token.startswith(prefix) for token in tokens)

    if has_prefix("philisterconsenior") or ("philister" in token_set and has_prefix("consenior")):
        return "philisterconsenior"
    if has_prefix("philisterschriftfuehrer") or ("philister" in token_set and has_prefix("schriftfuehrer")):
        return "philisterschriftfuehrer"
    if has_prefix("philisterkassier") or ("philister" in token_set and has_prefix("kassier")):
        return "philisterkassier"
    if has_prefix("philistersenior") or ("philister" in token_set and "senior" in token_set):
        return "philistersenior"

    if "consenior" in token_set:
        return "consenior"
    if "senior" in token_set:
        return "senior"
    if "schriftfuehrer" in token_set:
        return "schriftfuehrer"
    if "fuchsmajor" in token_set or "fm" in token_set or any(
        token.startswith("fm") and token[2:].isdigit() for token in tokens
    ):
        return "fuchsmajor"
    if "kassier" in token_set:
        return "kassier"
    if "barwart" in token_set:
        return "barwart"
    return None


def choose_reliable_semester_threshold(recorded_entries: pd.Series) -> int:
    nonzero = pd.to_numeric(recorded_entries, errors="coerce").fillna(0)
    nonzero = nonzero[nonzero > 0]
    if nonzero.empty:
        return 0
    median = float(nonzero.median())
    threshold = int(round(median * 0.75))
    return max(2, min(6, threshold))


def build_missing_mandatory_stats(counted_semester_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    empty_semester = pd.DataFrame(
        columns=[
            "Semester",
            "Year",
            "ExpectedSlots",
            "FilledSlots",
            "MissingSlots",
            "MissingPct",
            "RecordedEntries",
            "ReliableSemester",
            "MissingRoles",
            "FilledRoles",
            "FilledSlotIds",
        ]
    )
    empty_year = pd.DataFrame(
        columns=[
            "Year",
            "ExpectedSlots",
            "FilledSlots",
            "MissingSlots",
            "MissingPct",
            "SemestersInYear",
            "SemestersWithoutEntries",
            "ReliableSemesters",
            "FilledRolesYear",
        ]
    )

    if counted_semester_df.empty:
        return empty_semester, empty_year, 0

    source = counted_semester_df[counted_semester_df["Semester"] != UNKNOWN_SEMESTER].copy()
    if source.empty:
        return empty_semester, empty_year, 0
    all_semesters = sorted(source["Semester"].dropna().unique().tolist(), key=semester_sort_key)
    if not all_semesters:
        return empty_semester, empty_year, 0

    excluded_categories = {"Verbandschargen (Aktiven)", "Verbandschargen (Philister)", "Funktionaere"}
    work = source[~source["ChargenCategory"].isin(excluded_categories)].copy()
    if work.empty:
        return empty_semester, empty_year, 0

    full_semesters = semester_range_labels(all_semesters[0], all_semesters[-1])
    if not full_semesters:
        full_semesters = all_semesters

    work["MandatorySlot"] = work["ChargenRole"].apply(mandatory_slot_for_role)
    mandatory_df = work.dropna(subset=["MandatorySlot"]).drop_duplicates(subset=["Semester", "MandatorySlot"])

    recorded_entries = (
        work.groupby("Semester", as_index=False).size().rename(columns={"size": "RecordedEntries"})
    )
    slot_lists = (
        mandatory_df.groupby("Semester")["MandatorySlot"].apply(lambda s: sorted(set(s))).to_dict()
        if not mandatory_df.empty
        else {}
    )

    rows: list[dict[str, Any]] = []
    expected_slots = len(MANDATORY_SLOT_ORDER)
    recorded_lookup = dict(zip(recorded_entries["Semester"], recorded_entries["RecordedEntries"]))
    for semester in full_semesters:
        parsed = parse_semester_parts(semester)
        if parsed is None:
            continue
        year = parsed[1]
        present_slots = set(slot_lists.get(semester, []))
        filled = len(present_slots)
        missing_slot_ids = [slot for slot in MANDATORY_SLOT_ORDER if slot not in present_slots]
        missing_labels = ", ".join(MANDATORY_SLOT_LABELS[slot] for slot in missing_slot_ids)
        filled_slot_ids = [slot for slot in MANDATORY_SLOT_ORDER if slot in present_slots]
        filled_labels = ", ".join(MANDATORY_SLOT_LABELS[slot] for slot in filled_slot_ids)
        missing_slots = expected_slots - filled
        rows.append(
            {
                "Semester": semester,
                "Year": year,
                "ExpectedSlots": expected_slots,
                "FilledSlots": filled,
                "MissingSlots": missing_slots,
                "MissingPct": (missing_slots * 100.0 / expected_slots) if expected_slots else 0.0,
                "RecordedEntries": int(recorded_lookup.get(semester, 0)),
                "MissingRoles": missing_labels if missing_labels else "Keine",
                "FilledRoles": filled_labels if filled_labels else "Keine",
                "FilledSlotIds": "|".join(filled_slot_ids),
            }
        )

    semester_df = pd.DataFrame(rows)
    if semester_df.empty:
        return empty_semester, empty_year, 0
    threshold = choose_reliable_semester_threshold(semester_df["RecordedEntries"])
    semester_df["ReliableSemester"] = semester_df["RecordedEntries"] >= threshold
    semester_df["ReliableLabel"] = semester_df["ReliableSemester"].map(
        {True: "Zuverlaessig", False: "Datenluecke"}
    )
    semester_df = semester_df.sort_values("Semester", key=lambda s: s.map(semester_sort_key)).reset_index(drop=True)

    year_df = (
        semester_df.groupby("Year", as_index=False)
        .agg(
            ExpectedSlots=("ExpectedSlots", "sum"),
            FilledSlots=("FilledSlots", "sum"),
            MissingSlots=("MissingSlots", "sum"),
            SemestersInYear=("Semester", "size"),
            SemestersWithoutEntries=("RecordedEntries", lambda s: int((s == 0).sum())),
            ReliableSemesters=("ReliableSemester", "sum"),
            FilledRolesYear=(
                "FilledSlotIds",
                lambda s: ", ".join(
                    [
                        MANDATORY_SLOT_LABELS[slot]
                        for slot in MANDATORY_SLOT_ORDER
                        if any(
                            slot in {part for part in normalize_text(value).split("|") if part}
                            for value in s.dropna().tolist()
                        )
                    ]
                ),
            ),
        )
        .sort_values("Year")
    )
    year_df["FilledRolesYear"] = year_df["FilledRolesYear"].replace("", "Keine")
    year_df["MissingPct"] = year_df["MissingSlots"] * 100.0 / year_df["ExpectedSlots"]
    return semester_df, year_df, threshold


def age_bin_label(age_value: float | None) -> str | None:
    if age_value is None or pd.isna(age_value):
        return None
    age = float(age_value)
    if age < 25:
        return "<25"
    if age < 35:
        return "25-34"
    if age < 45:
        return "35-44"
    if age < 55:
        return "45-54"
    if age < 65:
        return "55-64"
    return "65+"


def build_intensity_per_person(
    included_semester_df: pd.DataFrame, reception_per_person: pd.DataFrame, today: date | None = None
) -> pd.DataFrame:
    if included_semester_df.empty:
        return pd.DataFrame(
            columns=[
                "Couleurname",
                "TotalChargen",
                "AktivVerbandAktivCount",
                "PhilVerbandPhilCount",
                "BasisYears",
                "BasisSource",
                "ReceptionDateDisplay",
                "AvgPerYear",
                "AvgPerYearAktivVerbandAktiv",
                "AvgPerYearPhilVerbandPhil",
            ]
        )

    today = today or date.today()
    work = included_semester_df.copy()
    if "ChargenCategory" not in work.columns:
        if "ChargenTyp" in work.columns:
            work["ChargenCategory"] = work["ChargenTyp"]
        else:
            work["ChargenCategory"] = "Unklare Chargen"
    work["SemesterStart"] = work["Semester"].apply(semester_to_start_date)
    grouped_rows: list[dict[str, Any]] = []
    reception_lookup = (
        reception_per_person.set_index("Couleurname")["ReceptionDate"]
        if not reception_per_person.empty
        else pd.Series(dtype="datetime64[ns]")
    )
    for couleurname, group in work.groupby("Couleurname"):
        total = int(len(group))
        aktiv_count = int(
            group["ChargenCategory"].isin(["Aktivenchargen", "Verbandschargen (Aktiven)"]).sum()
        )
        phil_count = int(
            group["ChargenCategory"].isin(["Philisterchargen", "Verbandschargen (Philister)"]).sum()
        )
        valid_dates = [d for d in group["SemesterStart"].tolist() if d is not None]
        if valid_dates:
            years = [d.year for d in valid_dates]
            fallback_years = float(max(years) - min(years) + 1)
        else:
            fallback_years = 1.0

        reception_dt = reception_lookup.get(couleurname) if not reception_lookup.empty else pd.NaT
        if pd.notna(reception_dt) and reception_dt.date() <= today:
            basis_years = max((today - reception_dt.date()).days / 365.2425, 1.0 / 365.2425)
            basis_source = "Reception"
            reception_display = reception_dt.strftime("%Y-%m-%d")
        else:
            basis_years = fallback_years
            basis_source = "SemesterSpan"
            reception_display = "n/a"

        avg_per_year = total / basis_years if basis_years > 0 else float(total)
        avg_aktiv = aktiv_count / basis_years if basis_years > 0 else float(aktiv_count)
        avg_phil = phil_count / basis_years if basis_years > 0 else float(phil_count)
        grouped_rows.append(
            {
                "Couleurname": couleurname,
                "TotalChargen": total,
                "AktivVerbandAktivCount": aktiv_count,
                "PhilVerbandPhilCount": phil_count,
                "BasisYears": basis_years,
                "BasisSource": basis_source,
                "ReceptionDateDisplay": reception_display,
                "AvgPerYear": avg_per_year,
                "AvgPerYearAktivVerbandAktiv": avg_aktiv,
                "AvgPerYearPhilVerbandPhil": avg_phil,
            }
        )

    out = pd.DataFrame(grouped_rows)
    return out.sort_values(["AvgPerYear", "TotalChargen"], ascending=False)


def semester_to_start_date(semester: str) -> date | None:
    if semester.startswith("SS "):
        try:
            return date(int(semester.split()[1]), 3, 1)
        except (ValueError, IndexError):
            return None
    if semester.startswith("WS "):
        try:
            year_text = semester.split()[1].split("/")[0]
            return date(int(year_text), 10, 1)
        except (ValueError, IndexError):
            return None
    return None


def classify_chargen_kind(
    semester: str, philistrierung_date: pd.Timestamp | pd.NaT, chargen_entry: str = ""
) -> str:
    if chargen_entry and is_philister_chargen_name(chargen_entry):
        return "Philisterchargen"
    if chargen_entry and is_active_chargen_name(chargen_entry):
        return "Aktivenchargen"
    semester_start = semester_to_start_date(semester)
    if semester_start is None or pd.isna(philistrierung_date):
        return "Unklare Chargen"
    if semester_start < philistrierung_date.date():
        return "Aktivenchargen"
    return "Philisterchargen"


def calculate_age_years(birth_date: pd.Timestamp, today: date) -> float:
    return (today - birth_date.date()).days / 365.2425


def average_age_for_statuses(df: pd.DataFrame, statuses: Iterable[str] | None = None) -> float | None:
    age_df = df.dropna(subset=["AgeYears"])
    if statuses is not None:
        status_set = {s for s in statuses if s}
        age_df = age_df[age_df["Mitgliedstatus"].isin(status_set)]
    if age_df.empty:
        return None
    return float(age_df["AgeYears"].mean())


def fmt_age(value: float | None) -> str:
    return "n/a" if value is None else f"{value:.1f} Jahre"


def build_member_records_from_excel_source(excel_source: Any) -> list[dict[str, Any]]:
    raw = pd.read_excel(excel_source, sheet_name=SHEET_NAME, header=HEADER_ROW_INDEX)
    required_columns = ["Couleurname", "Mitgliedstatus", "Geburtsdatum", "Chargen"]
    missing = set(required_columns) - set(raw.columns)
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(sorted(missing))}")

    df = raw[required_columns].copy()
    df["Philistrierung"] = raw["Philistrierung"] if "Philistrierung" in raw.columns else pd.NA
    df["Reception"] = raw["Reception"] if "Reception" in raw.columns else pd.NA
    df["Couleurname"] = df["Couleurname"].apply(normalize_text)
    df["Mitgliedstatus"] = df["Mitgliedstatus"].apply(normalize_text)
    df = df[df["Couleurname"] != ""].copy()

    df["Geburtsdatum"] = pd.to_datetime(df["Geburtsdatum"], errors="coerce", dayfirst=True)
    today = date.today()
    df["AgeYears"] = df["Geburtsdatum"].apply(
        lambda ts: calculate_age_years(ts, today) if pd.notna(ts) else float("nan")
    )
    df["AgeYears"] = df["AgeYears"].where(pd.notna(df["AgeYears"]), None)
    df["Philistrierung"] = pd.to_datetime(df["Philistrierung"], errors="coerce", dayfirst=True)
    df["PhilistrierungDate"] = df["Philistrierung"].dt.strftime("%Y-%m-%d")
    df["PhilistrierungDate"] = df["PhilistrierungDate"].where(df["Philistrierung"].notna(), None)
    df["Reception"] = pd.to_datetime(df["Reception"], errors="coerce", dayfirst=True)
    df["ReceptionDate"] = df["Reception"].dt.strftime("%Y-%m-%d")
    df["ReceptionDate"] = df["ReceptionDate"].where(df["Reception"].notna(), None)

    today = date.today()
    df["ChargenEntries"] = df["Chargen"].apply(parse_chargen_entries)
    df["TotalChargen"] = df["ChargenEntries"].apply(
        lambda entries: count_chargen_units([entry for entry in entries if not is_funktionaer_entry(entry)], today)
    ).astype(int)

    return df[
        [
            "Couleurname",
            "Mitgliedstatus",
            "AgeYears",
            "ReceptionDate",
            "PhilistrierungDate",
            "ChargenEntries",
            "TotalChargen",
        ]
    ].to_dict("records")


def dashboard_data_from_records(
    member_records: list[dict[str, Any]], manual_overrides: dict[str, str] | None = None
) -> dict[str, Any]:
    manual_overrides = canonicalize_override_map(manual_overrides)
    df = pd.DataFrame(member_records)
    if df.empty:
        df = pd.DataFrame(
            columns=[
                "Couleurname",
                "Mitgliedstatus",
                "AgeYears",
                "ReceptionDate",
                "PhilistrierungDate",
                "ChargenEntries",
                "TotalChargen",
            ]
        )

    df["Couleurname"] = df.get("Couleurname", "").apply(normalize_text)
    df["Mitgliedstatus"] = df.get("Mitgliedstatus", "").apply(normalize_text)
    df["AgeYears"] = pd.to_numeric(df.get("AgeYears"), errors="coerce")
    df["ReceptionDate"] = pd.to_datetime(df.get("ReceptionDate"), errors="coerce")
    df["PhilistrierungDate"] = pd.to_datetime(df.get("PhilistrierungDate"), errors="coerce")
    df["ChargenEntries"] = df.get("ChargenEntries", []).apply(
        lambda value: value if isinstance(value, list) else []
    )

    status_values = sorted([s for s in df["Mitgliedstatus"].dropna().unique() if s])
    person_groups = (
        df.groupby("Couleurname", as_index=False)["Mitgliedstatus"]
        .first()
        .assign(PersonGroup=lambda d: d["Mitgliedstatus"].apply(person_group_from_status))
    )

    semester_records: list[dict[str, Any]] = []
    unknown_counter: dict[str, int] = {}
    key_count: dict[str, int] = {}
    key_default_type_count: dict[str, dict[str, int]] = {}
    today = date.today()
    for row in df.itertuples(index=False):
        for entry in row.ChargenEntries:
            entry_key = chargen_override_key(entry) or entry
            entry_semesters = extract_entry_semesters(entry, today)
            default_meta = [default_entry_meta(entry, semester, row.PhilistrierungDate) for semester in entry_semesters]
            default_types = [meta[0] for meta in default_meta]
            key_count[entry_key] = key_count.get(entry_key, 0) + len(entry_semesters)
            type_counter = key_default_type_count.setdefault(entry_key, {})
            for typ in default_types:
                type_counter[typ] = type_counter.get(typ, 0) + 1

            has_unclear_default = any(t == "Unklare Chargen" for t in default_types)
            if has_unclear_default and entry_key not in manual_overrides:
                unknown_counter[entry_key] = unknown_counter.get(entry_key, 0) + len(entry_semesters)

            manual_class = manual_overrides.get(entry_key) or manual_overrides.get(entry)
            manual_type, manual_category, manual_include = manual_class_meta(manual_class)

            for semester, (default_type, default_category, default_include) in zip(entry_semesters, default_meta):
                chargen_type = manual_type if manual_type is not None else default_type
                chargen_category = manual_category if manual_category is not None else default_category
                include_in_chargen = manual_include if manual_include is not None else default_include
                count_group = special_count_group(row.Couleurname, entry_key) or entry_key
                semester_records.append(
                    {
                        "Couleurname": row.Couleurname,
                        "Semester": semester,
                        "ChargenTyp": chargen_type,
                        "ChargenCategory": chargen_category,
                        "IncludeInChargen": include_in_chargen,
                        "ChargenEntry": entry,
                        "ChargenEntryKey": entry_key,
                        "ChargenRole": role_name_from_entry(entry),
                        "CountGroup": count_group,
                    }
                )

    semester_df = pd.DataFrame(semester_records)
    if semester_df.empty:
        counted_semester_df = pd.DataFrame(
            columns=[
                "Couleurname",
                "Semester",
                "ChargenTyp",
                "ChargenCategory",
                "IncludeInChargen",
                "ChargenEntry",
                "ChargenEntryKey",
                "ChargenRole",
                "CountGroup",
            ]
        )
    else:
        counted_semester_df = semester_df[semester_df["IncludeInChargen"] == True].copy()  # noqa: E712
        counted_semester_df = counted_semester_df.drop_duplicates(
            subset=["Couleurname", "Semester", "CountGroup"]
        )

    role_person = (
        counted_semester_df.groupby(["ChargenRole", "Couleurname"], as_index=False)
        .size()
        .rename(columns={"size": "ChargenCount"})
        if not counted_semester_df.empty
        else pd.DataFrame(columns=["ChargenRole", "Couleurname", "ChargenCount"])
    )
    missing_semester_df, missing_year_df, missing_threshold = build_missing_mandatory_stats(counted_semester_df)
    role_totals = (
        role_person.groupby("ChargenRole", as_index=False)["ChargenCount"].sum().sort_values("ChargenCount", ascending=False)
        if not role_person.empty
        else pd.DataFrame(columns=["ChargenRole", "ChargenCount"])
    )
    role_values = role_totals["ChargenRole"].tolist() if not role_totals.empty else []

    age_df = df.dropna(subset=["AgeYears"]).copy()
    age_df["AgeBin"] = age_df["AgeYears"].apply(age_bin_label)
    age_distribution = (
        age_df.dropna(subset=["AgeBin"])
        .groupby("AgeBin", as_index=False)
        .size()
        .rename(columns={"size": "Count"})
        if not age_df.empty
        else pd.DataFrame(columns=["AgeBin", "Count"])
    )
    if not age_distribution.empty:
        age_order = {"<25": 0, "25-34": 1, "35-44": 2, "45-54": 3, "55-64": 4, "65+": 5}
        age_distribution["SortKey"] = age_distribution["AgeBin"].map(age_order).fillna(99)
        age_distribution = age_distribution.sort_values("SortKey").drop(columns=["SortKey"])

    status_distribution = (
        df.groupby("Mitgliedstatus", as_index=False)
        .size()
        .rename(columns={"size": "Count"})
        .sort_values("Count", ascending=False)
        if not df.empty
        else pd.DataFrame(columns=["Mitgliedstatus", "Count"])
    )
    if not status_distribution.empty:
        total_status = float(status_distribution["Count"].sum())
        status_distribution["Percent"] = status_distribution["Count"] * 100.0 / total_status

    reception_per_person = df.groupby("Couleurname", as_index=False)["ReceptionDate"].min()
    intensity_person = build_intensity_per_person(counted_semester_df, reception_per_person, today)

    per_person_total = (
        df[["Couleurname"]]
        .drop_duplicates()
        .merge(
            counted_semester_df.groupby("Couleurname", as_index=False).size().rename(columns={"size": "TotalChargen"})
            if not counted_semester_df.empty
            else pd.DataFrame(columns=["Couleurname", "TotalChargen"]),
            on="Couleurname",
            how="left",
        )
    )
    per_person_total["TotalChargen"] = pd.to_numeric(per_person_total["TotalChargen"], errors="coerce").fillna(0).astype(int)
    per_person_total = per_person_total.sort_values("TotalChargen", ascending=False)

    details_person = build_person_chargen_details(counted_semester_df)
    details_person_type = build_person_type_chargen_details(counted_semester_df)
    details_person_category = build_person_category_details(semester_df)
    if counted_semester_df.empty:
        semester_person = pd.DataFrame(columns=["Semester", "Couleurname", "ChargenCount"])
        kind_person = pd.DataFrame(
            columns=["Couleurname", "Aktivenchargen", "Philisterchargen", "Unklare Chargen"]
        )
        semester_values: list[str] = []
    else:
        semester_person = (
            counted_semester_df.groupby(["Semester", "Couleurname"], as_index=False)
            .size()
            .rename(columns={"size": "ChargenCount"})
        )
        kind_person = (
            counted_semester_df.groupby(["Couleurname", "ChargenTyp"]).size().unstack(fill_value=0).reset_index()
        )
        for col in ["Aktivenchargen", "Philisterchargen", "Unklare Chargen"]:
            if col not in kind_person.columns:
                kind_person[col] = 0
        kind_person = kind_person[["Couleurname", "Aktivenchargen", "Philisterchargen", "Unklare Chargen"]]
        semester_values = sorted(semester_person["Semester"].unique().tolist(), key=semester_sort_key)

    if semester_df.empty:
        category_person = pd.DataFrame(columns=["ChargenCategory", "Couleurname", "ChargenCount"])
    else:
        category_df = semester_df.drop_duplicates(
            subset=["ChargenCategory", "Couleurname", "Semester", "CountGroup"]
        )
        category_person = (
            category_df.groupby(["ChargenCategory", "Couleurname"], as_index=False)
            .size()
            .rename(columns={"size": "ChargenCount"})
        )

    unknown_entries = sorted(
        [{"entry": entry, "count": count} for entry, count in unknown_counter.items()],
        key=lambda item: (-item["count"], item["entry"]),
    )
    entry_candidates = []
    for entry_key, count in key_count.items():
        type_counts = key_default_type_count.get(entry_key, {})
        if type_counts:
            auto_type = sorted(type_counts.items(), key=lambda item: (-item[1], item[0]))[0][0]
        else:
            auto_type = "Unklare Chargen"
        entry_candidates.append(
            {
                "entry": entry_key,
                "count": count,
                "auto_type": auto_type,
                "is_unclear": entry_key in unknown_counter,
            }
        )
    entry_candidates = sorted(entry_candidates, key=lambda item: (-item["count"], item["entry"]))

    return {
        "df": df,
        "status_values": status_values,
        "person_groups": person_groups,
        "semester_values": semester_values,
        "per_person_total": per_person_total,
        "semester_person": semester_person,
        "kind_person": kind_person,
        "details_person": details_person,
        "details_person_type": details_person_type,
        "details_person_category": details_person_category,
        "role_person": role_person,
        "role_values": role_values,
        "missing_semester": missing_semester_df,
        "missing_year": missing_year_df,
        "missing_threshold": missing_threshold,
        "age_distribution": age_distribution,
        "status_distribution": status_distribution,
        "intensity_person": intensity_person,
        "category_person": category_person,
        "unknown_entries": unknown_entries,
        "entry_candidates": entry_candidates,
    }


def load_excel_data(excel_path: Path) -> dict[str, Any]:
    member_records = build_member_records_from_excel_source(excel_path)
    data = dashboard_data_from_records(member_records)
    data["member_records"] = member_records
    return data


def resolve_excel_path(explicit_path: str | None) -> Path:
    if explicit_path:
        path = Path(explicit_path).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"Excel file not found: {path}")
        return path

    candidates = sorted(Path.cwd().glob("Datenexport*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise FileNotFoundError("No Excel file found. Provide one with --file <path>.")
    return candidates[0].resolve()


def default_status_selection(status_values: list[str]) -> list[str]:
    _ = status_values
    return []


def status_option(status: str) -> dict[str, str]:
    return {"label": STATUS_LABELS.get(status, status), "value": status}


def default_role_selection(role_values: list[str]) -> list[str]:
    if not role_values:
        return []
    for role in role_values:
        norm = normalize_for_match(role)
        if re.match(r"^senior(\s|$)", norm):
            return [role]
    return role_values[:3]


def empty_bar_figure(
    title: str, xaxis_title: str = "Anzahl Chargen", yaxis_title: str = "Couleurname"
) -> go.Figure:
    fig = go.Figure()
    fig.update_layout(
        title=title,
        height=360,
        xaxis_title=xaxis_title,
        yaxis_title=yaxis_title,
        margin={"l": 10, "r": 10, "t": 56, "b": 36},
    )
    return fig


def compact_bar_height(
    item_count: int, min_height: int = 320, max_height: int = 560, row_height: int = 22, base_height: int = 140
) -> int:
    count = max(int(item_count), 1)
    return max(min_height, min(max_height, base_height + (count * row_height)))


def apply_compact_figure_layout(fig: go.Figure, height: int) -> go.Figure:
    fig.update_layout(
        height=height,
        margin={"l": 10, "r": 10, "t": 64, "b": 72},
        title={"x": 0.01, "xanchor": "left"},
        legend={"orientation": "h", "yanchor": "top", "y": -0.2, "xanchor": "left", "x": 0},
    )
    return fig


def decode_plotly_typed_arrays(value: Any) -> Any:
    if isinstance(value, dict):
        dtype = value.get("dtype")
        bdata = value.get("bdata")
        if isinstance(dtype, str) and isinstance(bdata, str) and set(value.keys()).issubset({"dtype", "bdata", "shape"}):
            try:
                raw = base64.b64decode(bdata)
                arr = np.frombuffer(raw, dtype=np.dtype(dtype))
                shape = value.get("shape")
                if isinstance(shape, (list, tuple)) and shape:
                    arr = arr.reshape(tuple(int(dim) for dim in shape))
                return arr.tolist()
            except Exception:  # noqa: BLE001
                return value
        return {k: decode_plotly_typed_arrays(v) for k, v in value.items()}
    if isinstance(value, list):
        return [decode_plotly_typed_arrays(v) for v in value]
    if isinstance(value, tuple):
        return tuple(decode_plotly_typed_arrays(v) for v in value)
    return value


def figure_dict_to_html_fragment(figure_obj: Any) -> str:
    if not figure_obj:
        return "<div class='empty-plot'>Keine Daten</div>"
    fig = figure_obj if isinstance(figure_obj, go.Figure) else go.Figure(figure_obj)
    fig_dict = decode_plotly_typed_arrays(fig.to_plotly_json())
    fig_clean = go.Figure(fig_dict)
    return pio.to_html(fig_clean, include_plotlyjs=False, full_html=False, config={"responsive": True})


def records_to_html_table(records: list[dict[str, Any]], title: str, max_rows: int = 500) -> str:
    if not records:
        return f"<section><h3>{std_html.escape(title)}</h3><p>Keine Daten</p></section>"
    rows = records[:max_rows]
    columns = list(rows[0].keys())
    header = "".join(f"<th>{std_html.escape(str(col))}</th>" for col in columns)
    body_rows: list[str] = []
    for row in rows:
        cells = "".join(f"<td>{std_html.escape(str(row.get(col, '')))}</td>" for col in columns)
        body_rows.append(f"<tr>{cells}</tr>")
    truncated_note = ""
    if len(records) > max_rows:
        truncated_note = f"<p class='muted'>Angezeigt: {max_rows} von {len(records)} Zeilen</p>"
    return (
        f"<section><h3>{std_html.escape(title)}</h3>{truncated_note}"
        f"<div class='table-wrap'><table><thead><tr>{header}</tr></thead><tbody>{''.join(body_rows)}</tbody></table></div></section>"
    )


def build_export_dashboard_html(
    source_label: str | None,
    stat_items: list[tuple[str, str]],
    chart_items: list[tuple[str, Any]],
    top20_rows: list[dict[str, Any]] | None,
    filtered_rows: list[dict[str, Any]] | None,
) -> str:
    chart_sections = []
    for title, fig_dict in chart_items:
        chart_html = figure_dict_to_html_fragment(fig_dict)
        chart_sections.append(
            f"<section class='chart'><h3>{std_html.escape(title)}</h3>{chart_html}</section>"
        )

    stats_html = "".join(
        (
            "<div class='stat-card'>"
            f"<div class='stat-label'>{std_html.escape(label)}</div>"
            f"<div class='stat-value'>{std_html.escape(value)}</div>"
            "</div>"
        )
        for label, value in stat_items
    )

    source_text = std_html.escape(source_label or "")
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    plotly_js_version = get_plotlyjs_version()

    return f"""<!doctype html>
<html lang="de">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Alb Stats Export</title>
  <script src="https://cdn.plot.ly/plotly-{plotly_js_version}.min.js"></script>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 14px; color: #222; background: #f5f6f8; }}
    .page {{ max-width: 1700px; margin: 0 auto; }}
    h1 {{ margin: 0 0 4px 0; font-size: 26px; }}
    h2 {{ margin: 16px 0 8px 0; font-size: 20px; }}
    h3 {{ margin: 8px 0 6px 0; font-size: 16px; }}
    .muted {{ color: #555; font-size: 13px; }}
    .stats-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 8px; margin: 8px 0 12px 0; }}
    .stat-card {{ border: 1px solid #ddd; border-radius: 4px; padding: 8px 10px; background: #fff; }}
    .stat-label {{ color: #555; font-size: 12px; }}
    .stat-value {{ font-size: 22px; font-weight: 600; margin-top: 3px; }}
    .chart-grid {{ display: flex; flex-direction: column; gap: 10px; }}
    .chart {{ border: 1px solid #ddd; border-radius: 4px; padding: 8px; background: #fff; }}
    .chart .js-plotly-plot, .chart .plot-container, .chart .plotly-graph-div {{ width: 100% !important; }}
    .table-wrap {{ overflow-x: auto; border: 1px solid #ddd; border-radius: 4px; }}
    table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
    th, td {{ border-bottom: 1px solid #eee; text-align: left; padding: 6px; white-space: nowrap; }}
    th {{ background: #f8f8f8; position: sticky; top: 0; }}
    .empty-plot {{ color: #666; font-style: italic; padding: 8px 4px; }}
  </style>
</head>
<body>
  <div class="page">
    <h1>Alb Stats Dashboard Export</h1>
    <div class="muted">{source_text}</div>
    <div class="muted">Exportiert am: {generated}</div>
    <h2>Kennzahlen</h2>
    <div class="stats-grid">{stats_html}</div>
    <h2>Diagramme</h2>
    <div class="chart-grid">{''.join(chart_sections)}</div>
    {records_to_html_table(top20_rows or [], "Top 20 Bundesbrueder")}
    {records_to_html_table(filtered_rows or [], "Chargen Tabelle (aktuelle Filterung)", max_rows=2000)}
  </div>
</body>
</html>"""


def build_app(data: dict[str, Any], excel_path: Path) -> Dash:
    initial_records: list[dict[str, Any]] = data["member_records"]
    initial_status_values: list[str] = data["status_values"]
    initial_semester_values: list[str] = data["semester_values"]
    initial_status_selection = default_status_selection(initial_status_values)
    overrides_file = Path.cwd() / OVERRIDES_FILE_NAME
    initial_overrides = load_persistent_overrides(overrides_file)
    if initial_overrides:
        save_persistent_overrides(overrides_file, initial_overrides)

    app = Dash(__name__)
    app.title = "Alb Stats"

    card_style = {
        "padding": "8px 10px",
        "border": "1px solid #ddd",
        "borderRadius": "4px",
        "background": "#fff",
        "minWidth": "160px",
    }
    section_style = {
        "border": "1px solid #ddd",
        "borderRadius": "4px",
        "padding": "10px",
        "marginBottom": "10px",
        "background": "#fff",
    }
    label_style = {"fontWeight": "600", "fontSize": "13px", "marginBottom": "4px", "display": "block"}

    app.layout = html.Div(
        style={
            "fontFamily": "Segoe UI, Arial, sans-serif",
            "padding": "10px 14px",
            "maxWidth": "1700px",
            "margin": "0 auto",
            "fontSize": "14px",
        },
        children=[
            dcc.Location(id="page-url", refresh=False),
            dcc.Store(id="members-store", data=initial_records),
            dcc.Store(id="manual-override-store", data=initial_overrides),
            dcc.Store(id="filtered-table-store", data=[]),
            dcc.Download(id="table-download"),
            dcc.Download(id="dashboard-html-download"),
            html.Div(
                style={
                    "display": "flex",
                    "justifyContent": "space-between",
                    "alignItems": "flex-end",
                    "flexWrap": "wrap",
                    "gap": "10px",
                    "marginBottom": "8px",
                },
                children=[
                    html.Div(
                        children=[
                            html.H1("Alb Stats", style={"margin": "0", "fontSize": "28px", "lineHeight": "1.1"}),
                            html.Div(f"Excel: {excel_path}", id="source-label", style={"color": "#555", "marginTop": "2px"}),
                        ]
                    ),
                    html.Div(
                        style={"display": "flex", "flexDirection": "column", "alignItems": "flex-end", "gap": "4px"},
                        children=[
                            dcc.Upload(
                                id="upload-data",
                                children=html.Button("Upload new Excel (.xlsx)"),
                                multiple=False,
                            ),
                            html.Div(id="upload-status", style={"color": "#444", "fontSize": "13px"}),
                        ],
                    ),
                ],
            ),
            html.Div(
                style={"display": "flex", "gap": "8px", "flexWrap": "wrap", "marginBottom": "10px"},
                children=[
                    html.Div(
                        style=card_style,
                        children=[
                            html.Div("Durchschnittsalter (Albertina)", style={"fontSize": "12px", "color": "#555"}),
                            html.H3(id="avg-all-value", style={"margin": "4px 0 0 0", "fontSize": "24px"}),
                        ],
                    ),
                    html.Div(
                        style=card_style,
                        children=[
                            html.Div("Durchschnittsalter (Philister)", style={"fontSize": "12px", "color": "#555"}),
                            html.H3(id="avg-up-bp-value", style={"margin": "4px 0 0 0", "fontSize": "24px"}),
                        ],
                    ),
                    html.Div(
                        style=card_style,
                        children=[
                            html.Div("Durchschnittsalter (Aktivitas)", style={"fontSize": "12px", "color": "#555"}),
                            html.H3(id="avg-bu-fu-value", style={"margin": "4px 0 0 0", "fontSize": "24px"}),
                        ],
                    ),
                    html.Div(
                        style=card_style,
                        children=[
                            html.Div(
                                "Durchschnittsalter (flexibel)",
                                id="avg-flex-label",
                                style={"fontSize": "12px", "color": "#555"},
                            ),
                            html.H3(id="avg-flex-value", style={"margin": "4px 0 0 0", "fontSize": "24px"}),
                        ],
                    ),
                    html.Div(
                        style=card_style,
                        children=[
                            html.Div("Medianalter", style={"fontSize": "12px", "color": "#555"}),
                            html.H3(id="median-age-value", style={"margin": "4px 0 0 0", "fontSize": "24px"}),
                        ],
                    ),
                    html.Div(
                        style=card_style,
                        children=[
                            html.Div("Durchschnitt Aktivenchargen / Person", style={"fontSize": "12px", "color": "#555"}),
                            html.H3(
                                id="avg-active-chargen-person-value",
                                style={"margin": "4px 0 0 0", "fontSize": "24px"},
                            ),
                        ],
                    ),
                    html.Div(
                        style=card_style,
                        children=[
                            html.Div(
                                "Durchschnitt Philisterchargen / Person",
                                style={"fontSize": "12px", "color": "#555"},
                            ),
                            html.H3(
                                id="avg-philister-chargen-person-value",
                                style={"margin": "4px 0 0 0", "fontSize": "24px"},
                            ),
                        ],
                    ),
                    html.Div(
                        style=card_style,
                        children=[
                            html.Div("Top 10% Cutoff (Chargen)", style={"fontSize": "12px", "color": "#555"}),
                            html.H3(id="top-percentile-value", style={"margin": "4px 0 0 0", "fontSize": "24px"}),
                        ],
                    ),
                ],
            ),
            html.Div(
                style=section_style,
                children=[
                    html.Div(
                        style={"display": "flex", "gap": "10px", "flexWrap": "wrap", "alignItems": "flex-end"},
                        children=[
                            html.Div(
                                style={"minWidth": "260px", "flex": "2"},
                                children=[
                                    html.Label("Statusfilter (flexibel):", style=label_style),
                                    dcc.Dropdown(
                                        id="status-filter",
                                        options=[status_option(status) for status in initial_status_values],
                                        value=initial_status_selection,
                                        multi=True,
                                        placeholder="Mitgliedstatus waehlen",
                                    ),
                                ],
                            ),
                            html.Div(
                                style={"minWidth": "220px", "flex": "1"},
                                children=[
                                    html.Label("Semesterfilter:", style=label_style),
                                    dcc.Dropdown(
                                        id="semester-filter",
                                        options=[{"label": "Alle Semester", "value": "__ALL__"}]
                                        + [{"label": value, "value": value} for value in initial_semester_values],
                                        value="__ALL__",
                                    ),
                                ],
                            ),
                            html.Div(
                                style={"minWidth": "240px", "flex": "1"},
                                children=[
                                    html.Label("Top N Bundesbrueder:", style=label_style),
                                    dcc.Slider(
                                        id="top-n-slider",
                                        min=5,
                                        max=50,
                                        step=5,
                                        value=20,
                                        marks={5: "5", 20: "20", 35: "35", 50: "50"},
                                    ),
                                ],
                            ),
                        ],
                    )
                ],
            ),
            html.Div(
                style={"display": "grid", "gridTemplateColumns": "repeat(auto-fit, minmax(360px, 1fr))", "gap": "10px"},
                children=[
                    html.Div(style=section_style, children=[dcc.Graph(id="age-distribution-chart")]),
                    html.Div(style=section_style, children=[dcc.Graph(id="status-distribution-chart")]),
                ],
            ),
            html.Div(
                style=section_style,
                children=[
                    html.Div(
                        style={
                            "display": "grid",
                            "gridTemplateColumns": "repeat(auto-fit, minmax(250px, 1fr))",
                            "gap": "10px",
                            "marginBottom": "6px",
                        },
                        children=[
                            html.Div(
                                children=[
                                    html.Label("Kategorie-Filter (1. Graph):", style=label_style),
                                    dcc.Dropdown(
                                        id="category-graph-select",
                                        options=[{"label": category, "value": category} for category in CATEGORY_GRAPH_OPTIONS],
                                        value=DEFAULT_CATEGORY_SELECTION.copy(),
                                        multi=True,
                                        clearable=True,
                                        placeholder="Kategorien auswaehlen",
                                    ),
                                ]
                            ),
                            html.Div(
                                children=[
                                    html.Label("Personengruppe (1. Graph):", style=label_style),
                                    dcc.Dropdown(
                                        id="category-person-group-filter",
                                        options=[{"label": g, "value": g} for g in PERSON_GROUP_OPTIONS],
                                        value=["Aktive", "Philister"],
                                        multi=True,
                                        clearable=False,
                                    ),
                                ]
                            ),
                        ],
                    ),
                    dcc.Graph(id="category-selectable-chart"),
                ],
            ),
            html.Div(
                style=section_style,
                children=[
                    html.Div(
                        style={"maxWidth": "520px", "marginBottom": "6px"},
                        children=[
                            html.Label("Sortierung 2. Graph:", style=label_style),
                            dcc.Dropdown(
                                id="kind-sort-mode",
                                options=[{"label": label, "value": value} for label, value in KIND_SORT_OPTIONS],
                                value="total",
                                clearable=False,
                            ),
                        ],
                    ),
                    dcc.Graph(id="chargen-kind-chart"),
                ],
            ),
            html.Div(
                style=section_style,
                children=[
                    html.Div(
                        style={"maxWidth": "760px", "marginBottom": "6px"},
                        children=[
                            html.Label("Chargen-Auswahl (Top Personen):", style=label_style),
                            dcc.Dropdown(
                                id="role-select-filter",
                                options=[],
                                value=[],
                                multi=True,
                                placeholder="Chargen waehlen",
                            ),
                        ],
                    ),
                    dcc.Graph(id="role-top-people-chart"),
                ],
            ),
            html.Div(
                style=section_style,
                children=[
                    html.Div(
                        style={"maxWidth": "760px", "marginBottom": "6px"},
                        children=[
                            html.Label("Intensitaet-Filter:", style=label_style),
                            dcc.Dropdown(
                                id="intensity-part-filter",
                                options=[{"label": value, "value": value} for value in INTENSITY_PART_OPTIONS],
                                value=INTENSITY_PART_OPTIONS.copy(),
                                multi=True,
                                clearable=True,
                            ),
                        ],
                    ),
                    dcc.Graph(id="chargen-intensity-chart"),
                ],
            ),
            html.Div(
                style=section_style,
                children=[
                    html.H4("Pflichtchargen-Datenluecken", style={"margin": "0 0 6px 0"}),
                    html.Div(id="missing-data-summary", style={"marginBottom": "8px", "color": "#444", "fontSize": "13px"}),
                    html.Div(
                        style={"display": "grid", "gridTemplateColumns": "repeat(auto-fit, minmax(360px, 1fr))", "gap": "10px"},
                        children=[
                            dcc.Graph(id="missing-semester-chart"),
                            dcc.Graph(id="missing-year-chart"),
                        ],
                    ),
                ],
            ),
            html.Div(
                style=section_style,
                children=[
                    html.H4("Chargen-Klassifikation manuell zuordnen", style={"margin": "0 0 8px 0"}),
                    html.Div(id="unclear-summary", style={"marginBottom": "8px", "color": "#444", "fontSize": "13px"}),
                    html.Div(
                        style={"display": "flex", "gap": "10px", "flexWrap": "wrap", "alignItems": "end"},
                        children=[
                            html.Div(
                                style={"minWidth": "420px", "flex": "2"},
                                children=[
                                    html.Label("Chargen-Position (ohne Datum):", style=label_style),
                                    dcc.Dropdown(
                                        id="unclear-entry-dropdown",
                                        options=[],
                                        value=None,
                                        placeholder="Eine Chargen-Position auswaehlen",
                                        maxHeight=520,
                                        optionHeight=38,
                                        style={"width": "100%"},
                                    ),
                                ],
                            ),
                            html.Div(
                                style={"minWidth": "260px", "flex": "1"},
                                children=[
                                    html.Label("Zuordnung:", style=label_style),
                                    dcc.Dropdown(
                                        id="unclear-class-dropdown",
                                        options=[{"label": label, "value": value} for label, value in MANUAL_CLASS_OPTIONS],
                                        value="aktiven",
                                        clearable=False,
                                    ),
                                ],
                            ),
                            html.Div(
                                style={"display": "flex", "gap": "6px"},
                                children=[
                                    html.Button("Zuordnung speichern", id="apply-unclear-btn"),
                                    html.Button("Zuordnung entfernen", id="remove-unclear-btn"),
                                ],
                            ),
                        ],
                    ),
                    html.Div(id="unclear-apply-status", style={"marginTop": "6px", "color": "#2f6f44", "fontSize": "13px"}),
                ],
            ),
            html.Div(
                style=section_style,
                children=[
                    html.Div(
                        style={"display": "flex", "gap": "8px", "flexWrap": "wrap"},
                        children=[
                            html.Button("Export CSV", id="export-csv-btn"),
                            html.Button("Export Excel", id="export-xlsx-btn"),
                            html.Button("Export HTML", id="export-html-btn"),
                        ],
                    ),
                    html.Div(id="export-status", style={"marginTop": "6px", "color": "#444", "fontSize": "13px"}),
                ],
            ),
            html.Div(
                style={"display": "none"},
                children=[
                    dash_table.DataTable(
                        id="top20-table",
                        page_size=20,
                        columns=[
                            {"name": "Couleurname", "id": "Couleurname"},
                            {"name": "Total Chargen", "id": "TotalChargen"},
                            {"name": "Aktivenchargen", "id": "Aktivenchargen"},
                            {"name": "Philisterchargen", "id": "Philisterchargen"},
                        ],
                    ),
                    dash_table.DataTable(
                        id="chargen-table",
                        page_size=20,
                        columns=[
                            {"name": "Couleurname", "id": "Couleurname"},
                            {"name": "Total Chargen", "id": "TotalChargen"},
                            {"name": "Aktivenchargen", "id": "Aktivenchargen"},
                            {"name": "Philisterchargen", "id": "Philisterchargen"},
                            {"name": "Unklare Chargen", "id": "Unklare Chargen"},
                            {"name": "Chargen im gewaehlten Semester", "id": "SemesterChargen"},
                        ],
                    ),
                ],
            ),
        ],
    )

    @app.callback(
        Output("members-store", "data"),
        Output("source-label", "children"),
        Output("upload-status", "children"),
        Input("upload-data", "contents"),
        State("upload-data", "filename"),
        prevent_initial_call=True,
    )
    def handle_upload(contents: str | None, filename: str | None):
        if not contents:
            return no_update, no_update, no_update

        try:
            _, content_string = contents.split(",", maxsplit=1)
            decoded = base64.b64decode(content_string)
            member_records = build_member_records_from_excel_source(io.BytesIO(decoded))
            file_label = filename or "uploaded file"
            return member_records, f"Excel: {file_label}", f"Loaded {len(member_records)} rows."
        except Exception as exc:  # noqa: BLE001
            return no_update, no_update, f"Upload failed: {exc}"

    @app.callback(
        Output("manual-override-store", "data"),
        Output("unclear-apply-status", "children"),
        Input("page-url", "pathname"),
        Input("apply-unclear-btn", "n_clicks"),
        Input("remove-unclear-btn", "n_clicks"),
        State("unclear-entry-dropdown", "value"),
        State("unclear-class-dropdown", "value"),
        State("manual-override-store", "data"),
    )
    def apply_unclear_override(
        pathname: str | None,
        apply_clicks: int | None,
        remove_clicks: int | None,
        unclear_entry: str | None,
        unclear_class: str | None,
        manual_overrides: dict[str, str] | None,
    ):
        _ = pathname, apply_clicks, remove_clicks
        if ctx.triggered_id == "page-url":
            disk_overrides = load_persistent_overrides(overrides_file)
            if disk_overrides:
                save_persistent_overrides(overrides_file, disk_overrides)
            return disk_overrides, f"Zuordnungen geladen: {len(disk_overrides)}"

        if not unclear_entry:
            return no_update, "Bitte zuerst eine Chargen-Position auswaehlen."
        updated = dict(manual_overrides or {})
        if ctx.triggered_id == "remove-unclear-btn":
            if unclear_entry in updated:
                del updated[unclear_entry]
                save_persistent_overrides(overrides_file, updated)
                return updated, f"Zuordnung entfernt: '{unclear_entry}'"
            return no_update, "Keine gespeicherte Zuordnung fuer diese Position vorhanden."

        if not unclear_class:
            return no_update, "Bitte zuerst eine Zuordnung waehlen."

        updated[unclear_entry] = unclear_class
        save_persistent_overrides(overrides_file, updated)
        label = next((label for label, value in MANUAL_CLASS_OPTIONS if value == unclear_class), unclear_class)
        return updated, f"Gespeichert: '{unclear_entry}' -> {label}"

    @app.callback(
        Output("status-filter", "options"),
        Output("status-filter", "value"),
        Output("semester-filter", "options"),
        Output("semester-filter", "value"),
        Output("role-select-filter", "options"),
        Output("role-select-filter", "value"),
        Input("members-store", "data"),
        Input("manual-override-store", "data"),
        State("role-select-filter", "value"),
    )
    def update_filter_options(
        member_records: list[dict[str, Any]] | None,
        manual_overrides: dict[str, str] | None,
        selected_roles: list[str] | None,
    ):
        data = dashboard_data_from_records(member_records or [], manual_overrides or {})
        status_values: list[str] = data["status_values"]
        semester_values: list[str] = data["semester_values"]
        role_values: list[str] = data["role_values"]

        status_options = [status_option(status) for status in status_values]
        status_value = default_status_selection(status_values)

        semester_options = [{"label": "Alle Semester", "value": "__ALL__"}] + [
            {"label": semester, "value": semester} for semester in semester_values
        ]
        semester_value = "__ALL__"
        role_options = [{"label": role, "value": role} for role in role_values]
        selected_set = set(selected_roles or [])
        valid_selected = [role for role in role_values if role in selected_set]
        if not valid_selected:
            valid_selected = default_role_selection(role_values)
        return status_options, status_value, semester_options, semester_value, role_options, valid_selected

    @app.callback(
        Output("avg-all-value", "children"),
        Output("avg-up-bp-value", "children"),
        Output("avg-bu-fu-value", "children"),
        Output("avg-flex-label", "children"),
        Output("avg-flex-value", "children"),
        Output("median-age-value", "children"),
        Output("avg-active-chargen-person-value", "children"),
        Output("avg-philister-chargen-person-value", "children"),
        Output("top-percentile-value", "children"),
        Output("unclear-entry-dropdown", "options"),
        Output("unclear-entry-dropdown", "value"),
        Output("unclear-summary", "children"),
        Output("age-distribution-chart", "figure"),
        Output("status-distribution-chart", "figure"),
        Output("missing-data-summary", "children"),
        Output("missing-semester-chart", "figure"),
        Output("missing-year-chart", "figure"),
        Output("chargen-intensity-chart", "figure"),
        Output("role-top-people-chart", "figure"),
        Output("category-selectable-chart", "figure"),
        Output("chargen-kind-chart", "figure"),
        Output("top20-table", "data"),
        Output("top20-table", "tooltip_data"),
        Output("chargen-table", "data"),
        Output("chargen-table", "tooltip_data"),
        Output("filtered-table-store", "data"),
        Input("members-store", "data"),
        Input("manual-override-store", "data"),
        Input("status-filter", "value"),
        Input("semester-filter", "value"),
        Input("top-n-slider", "value"),
        Input("role-select-filter", "value"),
        Input("category-graph-select", "value"),
        Input("category-person-group-filter", "value"),
        Input("intensity-part-filter", "value"),
        Input("kind-sort-mode", "value"),
        State("unclear-entry-dropdown", "value"),
    )
    def update_dashboard(
        member_records: list[dict[str, Any]] | None,
        manual_overrides: dict[str, str] | None,
        selected_statuses: list[str] | None,
        selected_semester: str | None,
        top_n: int | None,
        selected_roles: list[str] | None,
        selected_categories: list[str] | str | None,
        selected_person_groups: list[str] | str | None,
        selected_intensity_parts: list[str] | str | None,
        kind_sort_mode: str | None,
        current_selected_entry: str | None,
    ):
        data = dashboard_data_from_records(member_records or [], manual_overrides or {})
        df: pd.DataFrame = data["df"]
        person_groups: pd.DataFrame = data["person_groups"]
        per_person_total: pd.DataFrame = data["per_person_total"]
        semester_person: pd.DataFrame = data["semester_person"]
        category_person: pd.DataFrame = data["category_person"]
        role_person: pd.DataFrame = data["role_person"]
        age_distribution: pd.DataFrame = data["age_distribution"]
        status_distribution: pd.DataFrame = data["status_distribution"]
        missing_semester: pd.DataFrame = data["missing_semester"]
        missing_year: pd.DataFrame = data["missing_year"]
        missing_threshold: int = int(data["missing_threshold"])
        intensity_person: pd.DataFrame = data["intensity_person"]
        kind_person: pd.DataFrame = data["kind_person"]
        details_person: pd.DataFrame = data["details_person"]
        details_person_type: pd.DataFrame = data["details_person_type"]
        details_person_category: pd.DataFrame = data["details_person_category"]
        unknown_entries: list[dict[str, Any]] = data["unknown_entries"]
        entry_candidates: list[dict[str, Any]] = data["entry_candidates"]

        avg_all = fmt_age(average_age_for_statuses(df))
        avg_up_bp = fmt_age(average_age_for_statuses(df, UP_BP_STATUSES))
        avg_bu_fu = fmt_age(average_age_for_statuses(df, BU_FU_STATUSES))
        selected_status_codes = [normalize_text(s) for s in (selected_statuses or []) if normalize_text(s)]
        selected_status_codes = [s for s in selected_status_codes if s in set(data["status_values"])]
        if selected_status_codes:
            avg_flex_label = f"Durchschnittsalter ({' + '.join(selected_status_codes)})"
        else:
            avg_flex_label = "Durchschnittsalter (keine Auswahl)"
        avg_flex = fmt_age(average_age_for_statuses(df, selected_statuses or []))
        median_age = fmt_age(float(df["AgeYears"].dropna().median()) if not df["AgeYears"].dropna().empty else None)
        if per_person_total.empty:
            avg_active_chargen_person = "n/a"
            avg_philister_chargen_person = "n/a"
        else:
            people_df = per_person_total[["Couleurname"]].drop_duplicates()
            category_source = (
                category_person
                if not category_person.empty
                else pd.DataFrame(columns=["ChargenCategory", "Couleurname", "ChargenCount"])
            )

            active_avg_df = people_df.merge(
                category_source[
                    category_source["ChargenCategory"].isin(["Aktivenchargen", "Verbandschargen (Aktiven)"])
                ]
                .groupby("Couleurname", as_index=False)["ChargenCount"]
                .sum()
                .rename(columns={"ChargenCount": "ActivePlusCount"}),
                on="Couleurname",
                how="left",
            )
            active_avg_df["ActivePlusCount"] = active_avg_df["ActivePlusCount"].fillna(0)
            avg_active_chargen_person = f"{float(active_avg_df['ActivePlusCount'].mean()):.2f}"

            philister_avg_df = people_df.merge(
                category_source[
                    category_source["ChargenCategory"].isin(["Philisterchargen", "Verbandschargen (Philister)"])
                ]
                .groupby("Couleurname", as_index=False)["ChargenCount"]
                .sum()
                .rename(columns={"ChargenCount": "PhilisterPlusCount"}),
                on="Couleurname",
                how="left",
            )
            philister_avg_df["PhilisterPlusCount"] = philister_avg_df["PhilisterPlusCount"].fillna(0)
            avg_philister_chargen_person = f"{float(philister_avg_df['PhilisterPlusCount'].mean()):.2f}"
        if per_person_total.empty:
            top_percentile_value = "n/a"
        else:
            positive = per_person_total[per_person_total["TotalChargen"] > 0]["TotalChargen"]
            if positive.empty:
                top_percentile_value = "0"
            else:
                cutoff = float(positive.quantile(0.9))
                top_percentile_value = f">= {cutoff:.1f}"
        overrides = manual_overrides or {}
        unclear_options = []
        for item in entry_candidates:
            key = item["entry"]
            auto_type = item["auto_type"]
            count = item["count"]
            marker = " [unklar]" if item["is_unclear"] else ""
            override_text = f" [override: {overrides[key]}]" if key in overrides else ""
            label = f"{key} ({count}x, auto: {auto_type}){marker}{override_text}"
            unclear_options.append({"label": label, "value": key})

        option_values = {opt["value"] for opt in unclear_options}
        if current_selected_entry in option_values:
            unclear_value = current_selected_entry
        elif unclear_options:
            unclear_value = unclear_options[0]["value"]
        else:
            unclear_value = None
        unclear_summary = (
            f"Unklare offen: {len(unknown_entries)} | Gespeicherte Zuordnungen: {len(overrides)} | Datei: {OVERRIDES_FILE_NAME}"
        )

        limit = int(top_n or 20)

        if age_distribution.empty:
            age_fig = empty_bar_figure(
                "Keine Altersdaten vorhanden", xaxis_title="Altersgruppe", yaxis_title="Anzahl Personen"
            )
        else:
            age_fig = px.bar(
                age_distribution,
                x="AgeBin",
                y="Count",
                title="Altersverteilung",
                labels={"AgeBin": "Altersgruppe", "Count": "Anzahl Personen"},
                text="Count",
            )
            apply_compact_figure_layout(age_fig, 320)

        if status_distribution.empty:
            status_fig = empty_bar_figure(
                "Keine Statusdaten vorhanden", xaxis_title="Status", yaxis_title="Anzahl Personen"
            )
        else:
            status_df = status_distribution.copy()
            status_df["PercentLabel"] = status_df["Percent"].map(lambda x: f"{x:.1f}%")
            status_fig = px.bar(
                status_df,
                x="Mitgliedstatus",
                y="Count",
                text="PercentLabel",
                title="Statusverteilung (Anzahl und Anteil)",
                labels={"Mitgliedstatus": "Status", "Count": "Anzahl"},
            )
            apply_compact_figure_layout(status_fig, 320)

        if missing_semester.empty:
            missing_summary = "Keine Semesterdaten fuer Pflichtchargen verfuegbar."
            missing_semester_fig = empty_bar_figure(
                "Keine Pflichtchargen-Daten", xaxis_title="Semester", yaxis_title="Fehlende Pflichtslots"
            )
            missing_year_fig = empty_bar_figure(
                "Keine Jahresdaten", xaxis_title="Jahr", yaxis_title="Fehlende Pflichtslots"
            )
        else:
            missing_summary = (
                f"Standard: 10 Pflichtchargen/Semester (Aktive 6 + Philister 4). "
                f"Verbandschargen/Funktionaere sind optional und ausgeschlossen. "
                f"Start: {missing_semester.iloc[0]['Semester']}. "
                f"Zuverlaessig ab >= {missing_threshold} Eintraegen/Semester "
                f"(automatisch aus Median abgeleitet)."
            )
            semester_plot = missing_semester.copy()
            semester_order = semester_plot["Semester"].tolist()
            semester_plot["Semester"] = pd.Categorical(
                semester_plot["Semester"], categories=semester_order, ordered=True
            )
            missing_semester_fig = px.bar(
                semester_plot,
                x="Semester",
                y="MissingSlots",
                color="ReliableLabel",
                title="Fehlende Pflichtchargen je Semester",
                labels={"MissingSlots": "Fehlende Pflichtslots", "Semester": "Semester", "ReliableLabel": "Datenqualitaet"},
                custom_data=["FilledSlots", "ExpectedSlots", "RecordedEntries", "MissingPct", "FilledRoles", "MissingRoles"],
                color_discrete_map={"Zuverlaessig": "#4472C4", "Datenluecke": "#D9534F"},
            )
            missing_semester_fig.update_traces(
                hovertemplate="<b>%{x}</b><br>Fehlend: %{y}<br>Gefuellt: %{customdata[0]}/%{customdata[1]}<br>Erfasste Eintraege: %{customdata[2]}<br>Fehlend %%: %{customdata[3]:.1f}<br>Erfuellte Rollen: %{customdata[4]}<br>Fehlende Rollen: %{customdata[5]}<extra></extra>"
            )
            missing_semester_fig.update_xaxes(tickangle=-60)
            apply_compact_figure_layout(missing_semester_fig, 380)

            if missing_year.empty:
                missing_year_fig = empty_bar_figure(
                    "Keine Jahresdaten", xaxis_title="Jahr", yaxis_title="Fehlende Pflichtslots"
                )
            else:
                year_plot = missing_year.copy()
                year_plot["Year"] = year_plot["Year"].astype(str)
                year_plot["MissingPctLabel"] = year_plot["MissingPct"].map(lambda v: f"{v:.1f}%")
                missing_year_fig = px.bar(
                    year_plot,
                    x="Year",
                    y="MissingSlots",
                    text="MissingPctLabel",
                    title="Fehlende Pflichtchargen je Jahr",
                    labels={"Year": "Jahr", "MissingSlots": "Fehlende Pflichtslots"},
                    custom_data=[
                        "FilledSlots",
                        "ExpectedSlots",
                        "SemestersWithoutEntries",
                        "ReliableSemesters",
                        "SemestersInYear",
                        "FilledRolesYear",
                    ],
                )
                missing_year_fig.update_traces(
                    hovertemplate="<b>%{x}</b><br>Fehlend: %{y}<br>Gefuellt: %{customdata[0]}/%{customdata[1]}<br>Semester ohne Eintraege: %{customdata[2]}<br>Zuverlaessige Semester: %{customdata[3]}/%{customdata[4]}<br>Erfuellte Rollen (mind. 1x): %{customdata[5]}<extra></extra>"
                )
                apply_compact_figure_layout(missing_year_fig, 380)
        if isinstance(selected_intensity_parts, str):
            selected_intensity_parts = [selected_intensity_parts]
        selected_intensity_parts = [p for p in (selected_intensity_parts or []) if p in INTENSITY_PART_OPTIONS]

        intensity_key_to_label = {
            "AvgPerYearAktivVerbandAktiv": "Aktiven + Verbandschargen (Aktiven)",
            "AvgPerYearPhilVerbandPhil": "Philister + Verbandschargen (Philister)",
        }
        selected_value_cols = [
            key for key, label in intensity_key_to_label.items() if label in selected_intensity_parts
        ]

        if intensity_person.empty:
            intensity_fig = empty_bar_figure("Keine Intensitaetsdaten vorhanden")
        elif not selected_value_cols:
            intensity_fig = empty_bar_figure("Keine Intensitaets-Kategorien ausgewaehlt")
        else:
            intensity_rank = intensity_person.copy()
            intensity_rank["SelectedAvgPerYear"] = intensity_rank[selected_value_cols].sum(axis=1)
            intensity_rank = intensity_rank.sort_values(
                ["SelectedAvgPerYear", "TotalChargen"], ascending=[False, False]
            ).head(limit)
            order = intensity_rank.sort_values("SelectedAvgPerYear", ascending=True)["Couleurname"].tolist()

            intensity_long = intensity_rank.melt(
                id_vars=[
                    "Couleurname",
                    "TotalChargen",
                    "AktivVerbandAktivCount",
                    "PhilVerbandPhilCount",
                    "BasisYears",
                    "BasisSource",
                    "ReceptionDateDisplay",
                    "SelectedAvgPerYear",
                ],
                value_vars=selected_value_cols,
                var_name="IntensityPart",
                value_name="AvgPerYearPart",
            )
            intensity_long["IntensityPart"] = intensity_long["IntensityPart"].map(intensity_key_to_label)
            intensity_long["PartCount"] = intensity_long.apply(
                lambda row: row["AktivVerbandAktivCount"]
                if row["IntensityPart"] == "Aktiven + Verbandschargen (Aktiven)"
                else row["PhilVerbandPhilCount"],
                axis=1,
            )
            intensity_long["Couleurname"] = pd.Categorical(
                intensity_long["Couleurname"], categories=order, ordered=True
            )
            intensity_long = intensity_long.sort_values(["Couleurname", "IntensityPart"])

            intensity_fig = px.bar(
                intensity_long,
                x="AvgPerYearPart",
                y="Couleurname",
                color="IntensityPart",
                orientation="h",
                barmode="stack",
                title=f"Chargen-Intensitaet pro Person (Top {limit})",
                labels={"AvgPerYearPart": "Durchschnitt Chargen/Jahr", "Couleurname": "Couleurname"},
                custom_data=[
                    "PartCount",
                    "TotalChargen",
                    "BasisYears",
                    "BasisSource",
                    "ReceptionDateDisplay",
                    "SelectedAvgPerYear",
                ],
            )
            intensity_fig.update_traces(
                hovertemplate="<b>%{y} (Anteil ?/Jahr: %{x:.2f})</b><br>Kategorie: %{fullData.name}<br>Kategorie-Count: %{customdata[0]}<br>Total: %{customdata[1]}<br>Total (gewaehlte Teile) ?/Jahr: %{customdata[5]:.2f}<br>Basis: %{customdata[3]}<br>Reception: %{customdata[4]}<br>Jahre seit Reception/Basis: %{customdata[2]:.2f}<extra></extra>"
            )
            apply_compact_figure_layout(
                intensity_fig,
                compact_bar_height(len(order), min_height=360, max_height=560, row_height=19, base_height=190),
            )

        selected_roles = [r for r in (selected_roles or []) if normalize_text(r) != ""]
        if role_person.empty:
            role_fig = empty_bar_figure("Keine Chargen-Rollen vorhanden")
        else:
            role_filtered = role_person[role_person["ChargenRole"].isin(selected_roles)] if selected_roles else role_person
            role_counts = (
                role_filtered.groupby("Couleurname", as_index=False)["ChargenCount"].sum()
                if not role_filtered.empty
                else pd.DataFrame(columns=["Couleurname", "ChargenCount"])
            )
            if role_counts.empty:
                role_fig = empty_bar_figure("Keine Daten fuer ausgewaehlte Chargen")
            else:
                role_counts = role_counts.merge(details_person, on="Couleurname", how="left")
                role_counts["ChargenDetailsHtml"] = role_counts["ChargenDetailsHtml"].fillna("Keine Chargen-Details")
                role_counts = role_counts.sort_values("ChargenCount", ascending=False).head(limit)
                role_counts = role_counts.sort_values("ChargenCount", ascending=True)
                role_title = ", ".join(selected_roles) if selected_roles else "alle Chargen"
                role_fig = px.bar(
                    role_counts,
                    x="ChargenCount",
                    y="Couleurname",
                    orientation="h",
                    title=f"Top {limit} Personen fuer {role_title}",
                    labels={"ChargenCount": "Anzahl", "Couleurname": "Couleurname"},
                    custom_data=["ChargenDetailsHtml"],
                )
                role_fig.update_traces(
                    hovertemplate="<b>%{y} (Anzahl: %{x})</b><br>%{customdata[0]}<extra></extra>",
                )
                apply_compact_figure_layout(
                    role_fig,
                    compact_bar_height(len(role_counts), min_height=340, max_height=560, row_height=22, base_height=160),
                )
        if isinstance(selected_categories, str):
            selected_categories = [selected_categories]
        selected_categories = [c for c in (selected_categories or []) if c in CATEGORY_GRAPH_OPTIONS]
        if not selected_categories:
            selected_categories = DEFAULT_CATEGORY_SELECTION.copy()
        if isinstance(selected_person_groups, str):
            selected_person_groups = [selected_person_groups]
        selected_person_groups = [g for g in (selected_person_groups or []) if g in PERSON_GROUP_OPTIONS]
        if not selected_person_groups:
            selected_person_groups = PERSON_GROUP_OPTIONS.copy()

        category_counts = (
            category_person[category_person["ChargenCategory"].isin(selected_categories)]
            if not category_person.empty
            else pd.DataFrame(columns=["ChargenCategory", "Couleurname", "ChargenCount"])
        )
        if not category_counts.empty and not person_groups.empty:
            category_counts = category_counts.merge(
                person_groups[["Couleurname", "PersonGroup"]],
                on="Couleurname",
                how="left",
            )
            category_counts = category_counts[category_counts["PersonGroup"].isin(selected_person_groups)].copy()
        if category_counts.empty:
            category_fig = empty_bar_figure("Keine Daten fuer die ausgewaehlten Kategorien")
        else:
            category_df = category_counts.merge(
                details_person_category[
                    ["Couleurname", "ChargenCategory", "CategoryDetailsHtml", "CategoryDetailsText"]
                ],
                on=["Couleurname", "ChargenCategory"],
                how="left",
            )
            category_df["CategoryDetailsHtml"] = category_df["CategoryDetailsHtml"].fillna("Keine Chargen-Details")
            ranking = (
                category_df.groupby("Couleurname", as_index=False)["ChargenCount"]
                .sum()
                .sort_values("ChargenCount", ascending=False)
                .head(limit)
            )
            category_df = category_df[category_df["Couleurname"].isin(ranking["Couleurname"])]
            order = ranking.sort_values("ChargenCount", ascending=True)["Couleurname"].tolist()
            category_df["Couleurname"] = pd.Categorical(category_df["Couleurname"], categories=order, ordered=True)
            category_df = category_df.sort_values(["Couleurname", "ChargenCategory"])
            category_fig = px.bar(
                category_df,
                x="ChargenCount",
                y="Couleurname",
                color="ChargenCategory",
                orientation="h",
                barmode="stack",
                title=f"Top {limit} Bundesbrueder nach Kategorien",
                labels={"ChargenCount": "Anzahl", "Couleurname": "Couleurname"},
                custom_data=["CategoryDetailsHtml"],
            )
            category_fig.update_traces(
                hovertemplate="<b>%{y} (Anzahl: %{x})</b><br>Kategorie: %{fullData.name}<br>%{customdata[0]}<extra></extra>",
            )
            apply_compact_figure_layout(
                category_fig,
                compact_bar_height(len(order), min_height=340, max_height=560, row_height=22, base_height=165),
            )

        if selected_semester == "__ALL__" or not selected_semester:
            sem_counts = (
                semester_person.groupby("Couleurname", as_index=False)["ChargenCount"].sum()
                if not semester_person.empty
                else pd.DataFrame(columns=["Couleurname", "ChargenCount"])
            )
            semester_title = "Alle Semester"
        else:
            sem_counts = (
                semester_person[semester_person["Semester"] == selected_semester][["Couleurname", "ChargenCount"]]
                if not semester_person.empty
                else pd.DataFrame(columns=["Couleurname", "ChargenCount"])
            )
            semester_title = selected_semester

        table_df = per_person_total.merge(
            kind_person,
            on="Couleurname",
            how="left",
        ).merge(
            sem_counts.rename(columns={"ChargenCount": "SemesterChargen"}),
            on="Couleurname",
            how="left",
        ).merge(
            details_person,
            on="Couleurname",
            how="left",
        )
        if not table_df.empty:
            for col in ["Aktivenchargen", "Philisterchargen", "Unklare Chargen"]:
                table_df[col] = table_df[col].fillna(0).astype(int)
            table_df["SemesterChargen"] = table_df["SemesterChargen"].fillna(0).astype(int)
            table_df["ChargenDetailsText"] = table_df["ChargenDetailsText"].fillna("Keine Chargen-Details")
            table_df["ChargenDetailsHtml"] = table_df["ChargenDetailsHtml"].fillna("Keine Chargen-Details")
            table_df = table_df.sort_values(["TotalChargen", "SemesterChargen"], ascending=False)

        if table_df.empty:
            kind_fig = empty_bar_figure("Keine Chargen-Typen vorhanden")
            top20_data: list[dict[str, Any]] = []
            top20_tooltips: list[dict[str, Any]] = []
        else:
            sort_map = {
                "total": "TotalChargen",
                "active": "Aktivenchargen",
                "philister": "Philisterchargen",
                "unclear": "Unklare Chargen",
                "semester": "SemesterChargen",
                "name": "Couleurname",
            }
            sort_label_map = {
                "total": "Gesamtchargen",
                "active": "Aktivenchargen",
                "philister": "Philisterchargen",
                "unclear": "Unklare Chargen",
                "semester": f"Semester ({semester_title})",
                "name": "Name A-Z",
            }
            sort_col = sort_map.get(kind_sort_mode or "total", "TotalChargen")
            if sort_col == "Couleurname":
                ranked_df = table_df.sort_values(["Couleurname"], ascending=True)
            else:
                ranked_df = table_df.sort_values([sort_col, "TotalChargen"], ascending=[False, False])

            top_kind_df = ranked_df.head(limit)[
                [
                    "Couleurname",
                    "Aktivenchargen",
                    "Philisterchargen",
                    "Unklare Chargen",
                    "ChargenDetailsHtml",
                ]
            ].copy()
            if sort_col == "Couleurname":
                top_kind_df = top_kind_df.sort_values("Couleurname", ascending=False)
            else:
                sort_values = ranked_df.head(limit)[sort_col]
                top_kind_df["__SortValue"] = sort_values.values
                top_kind_df = top_kind_df.sort_values("__SortValue", ascending=True).drop(columns=["__SortValue"])
            top_kind_long = top_kind_df.melt(
                id_vars=["Couleurname", "ChargenDetailsHtml"], var_name="ChargenTyp", value_name="Anzahl"
            )
            top_kind_long = top_kind_long.merge(
                details_person_type[["Couleurname", "ChargenTyp", "ChargenTypeDetailsHtml"]],
                on=["Couleurname", "ChargenTyp"],
                how="left",
            )
            top_kind_long["ChargenTypeDetailsHtml"] = top_kind_long["ChargenTypeDetailsHtml"].fillna(
                top_kind_long["ChargenDetailsHtml"]
            )
            kind_fig = px.bar(
                top_kind_long,
                x="Anzahl",
                y="Couleurname",
                color="ChargenTyp",
                orientation="h",
                barmode="stack",
                title=f"Top {limit} Bundesbrueder nach Chargen-Typen ({sort_label_map.get(kind_sort_mode or 'total', 'Gesamtchargen')})",
                labels={"Anzahl": "Anzahl Chargen", "Couleurname": "Couleurname"},
                custom_data=["ChargenTypeDetailsHtml"],
            )
            kind_fig.update_traces(
                hovertemplate="<b>%{y} (Anzahl: %{x})</b><br>Typ: %{fullData.name}<br>%{customdata[0]}<extra></extra>",
            )
            apply_compact_figure_layout(
                kind_fig,
                compact_bar_height(len(top_kind_df), min_height=340, max_height=560, row_height=22, base_height=165),
            )
            top20_df = table_df.head(20)[
                ["Couleurname", "TotalChargen", "Aktivenchargen", "Philisterchargen", "ChargenDetailsText"]
            ].copy()
            top20_data = top20_df.drop(columns=["ChargenDetailsText"]).to_dict("records")
            top20_tooltips = [
                {"Couleurname": {"value": row["ChargenDetailsText"], "type": "markdown"}}
                for row in top20_df.to_dict("records")
            ]

        if table_df.empty:
            table_data: list[dict[str, Any]] = []
            table_tooltips: list[dict[str, Any]] = []
        else:
            display_cols = [
                "Couleurname",
                "TotalChargen",
                "Aktivenchargen",
                "Philisterchargen",
                "Unklare Chargen",
                "SemesterChargen",
            ]
            table_display_df = table_df[display_cols + ["ChargenDetailsText"]].copy()
            table_data = table_display_df[display_cols].to_dict("records")
            table_tooltips = [
                {"Couleurname": {"value": row["ChargenDetailsText"], "type": "markdown"}}
                for row in table_display_df.to_dict("records")
            ]

        return (
            avg_all,
            avg_up_bp,
            avg_bu_fu,
            avg_flex_label,
            avg_flex,
            median_age,
            avg_active_chargen_person,
            avg_philister_chargen_person,
            top_percentile_value,
            unclear_options,
            unclear_value,
            unclear_summary,
            age_fig,
            status_fig,
            missing_summary,
            missing_semester_fig,
            missing_year_fig,
            intensity_fig,
            role_fig,
            category_fig,
            kind_fig,
            top20_data,
            top20_tooltips,
            table_data,
            table_tooltips,
            table_data,
        )

    @app.callback(
        Output("table-download", "data"),
        Output("dashboard-html-download", "data"),
        Output("export-status", "children"),
        Input("export-csv-btn", "n_clicks"),
        Input("export-xlsx-btn", "n_clicks"),
        Input("export-html-btn", "n_clicks"),
        State("filtered-table-store", "data"),
        State("members-store", "data"),
        State("manual-override-store", "data"),
        State("status-filter", "value"),
        State("semester-filter", "value"),
        State("top-n-slider", "value"),
        State("role-select-filter", "value"),
        State("category-graph-select", "value"),
        State("category-person-group-filter", "value"),
        State("intensity-part-filter", "value"),
        State("kind-sort-mode", "value"),
        State("unclear-entry-dropdown", "value"),
        State("source-label", "children"),
        prevent_initial_call=True,
    )
    def export_filtered_data(
        csv_clicks: int | None,
        xlsx_clicks: int | None,
        html_clicks: int | None,
        filtered_rows: list[dict[str, Any]] | None,
        member_records: list[dict[str, Any]] | None,
        manual_overrides: dict[str, str] | None,
        selected_statuses: list[str] | None,
        selected_semester: str | None,
        top_n: int | None,
        selected_roles: list[str] | None,
        selected_categories: list[str] | str | None,
        selected_person_groups: list[str] | str | None,
        selected_intensity_parts: list[str] | str | None,
        kind_sort_mode: str | None,
        current_selected_entry: str | None,
        source_label: str | None,
    ):
        _ = csv_clicks, xlsx_clicks, html_clicks
        if not filtered_rows:
            if ctx.triggered_id != "export-html-btn":
                return no_update, no_update, "No rows to export."

        table_df = pd.DataFrame(filtered_rows)
        if table_df.empty and ctx.triggered_id != "export-html-btn":
            return no_update, no_update, "No rows to export."

        if ctx.triggered_id == "export-csv-btn":
            return dcc.send_data_frame(table_df.to_csv, "chargen_filtered.csv", index=False), no_update, (
                f"Exported {len(table_df)} rows to chargen_filtered.csv"
            )
        if ctx.triggered_id == "export-xlsx-btn":
            return dcc.send_data_frame(table_df.to_excel, "chargen_filtered.xlsx", index=False), no_update, (
                f"Exported {len(table_df)} rows to chargen_filtered.xlsx"
            )
        if ctx.triggered_id == "export-html-btn":
            compute_dashboard = getattr(update_dashboard, "__wrapped__", update_dashboard)
            dashboard_out = compute_dashboard(
                member_records,
                manual_overrides,
                selected_statuses,
                selected_semester,
                top_n,
                selected_roles,
                selected_categories,
                selected_person_groups,
                selected_intensity_parts,
                kind_sort_mode,
                current_selected_entry,
            )
            avg_all = dashboard_out[0]
            avg_phil = dashboard_out[1]
            avg_aktivitas = dashboard_out[2]
            avg_flex_label = dashboard_out[3]
            avg_flex = dashboard_out[4]
            median_age = dashboard_out[5]
            avg_active_chargen = dashboard_out[6]
            avg_phil_chargen = dashboard_out[7]
            top_percentile = dashboard_out[8]
            age_fig = dashboard_out[12]
            status_fig = dashboard_out[13]
            missing_semester_fig = dashboard_out[15]
            missing_year_fig = dashboard_out[16]
            intensity_fig = dashboard_out[17]
            role_fig = dashboard_out[18]
            category_fig = dashboard_out[19]
            kind_fig = dashboard_out[20]
            top20_rows = dashboard_out[21]
            exported_filtered_rows = dashboard_out[25]

            stat_items = [
                ("Durchschnittsalter (Albertina)", str(avg_all)),
                ("Durchschnittsalter (Philister)", str(avg_phil)),
                ("Durchschnittsalter (Aktivitas)", str(avg_aktivitas)),
                (str(avg_flex_label or "Durchschnittsalter"), str(avg_flex)),
                ("Medianalter", str(median_age)),
                ("Durchschnitt Aktivenchargen / Person", str(avg_active_chargen)),
                ("Durchschnitt Philisterchargen / Person", str(avg_phil_chargen)),
                ("Top 10% Cutoff (Chargen)", str(top_percentile)),
            ]
            chart_items = [
                ("Altersverteilung", age_fig),
                ("Statusverteilung", status_fig),
                ("Kategorien (1. Graph)", category_fig),
                ("Chargen-Typen (2. Graph)", kind_fig),
                ("Top Personen nach Charge", role_fig),
                ("Chargen-Intensitaet", intensity_fig),
                ("Fehlende Pflichtchargen je Semester", missing_semester_fig),
                ("Fehlende Pflichtchargen je Jahr", missing_year_fig),
            ]
            html_doc = build_export_dashboard_html(
                source_label=source_label,
                stat_items=stat_items,
                chart_items=chart_items,
                top20_rows=top20_rows or [],
                filtered_rows=exported_filtered_rows or [],
            )
            return (
                no_update,
                {
                    "content": html_doc,
                    "filename": "alb_stats_dashboard_export.html",
                    "type": "text/html",
                },
                "Exported interactive dashboard to alb_stats_dashboard_export.html",
            )
        return no_update, no_update, no_update

    return app


def main() -> None:
    parser = argparse.ArgumentParser(description="Alb Excel stats dashboard")
    parser.add_argument("--file", type=str, default=None, help="Path to Excel export (.xlsx)")
    parser.add_argument("--host", type=str, default="127.0.0.1", help="Host for Dash app")
    parser.add_argument("--port", type=int, default=8050, help="Port for Dash app")
    parser.add_argument("--check-only", action="store_true", help="Only validate and print parsed summary")
    args = parser.parse_args()

    excel_path = resolve_excel_path(args.file)
    data = load_excel_data(excel_path)

    if args.check_only:
        df: pd.DataFrame = data["df"]
        semesters: list[str] = data["semester_values"]
        print(f"Excel: {excel_path}")
        print(f"Rows with Couleurname: {len(df)}")
        print(f"Statuses: {sorted(df['Mitgliedstatus'].dropna().unique().tolist())}")
        print(f"Semesters in chargen: {len(semesters)}")
        print(f"Average age all: {fmt_age(average_age_for_statuses(df))}")
        return

    app = build_app(data, excel_path)
    app.run(host=args.host, port=args.port, debug=False)


if __name__ == "__main__":
    main()
