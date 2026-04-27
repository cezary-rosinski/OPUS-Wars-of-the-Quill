from __future__ import annotations

import json
import re
import time
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional
from urllib.parse import urlparse

import pandas as pd
import requests

#%%
# ============================================================================
# CONFIG
# ============================================================================

DEFAULT_INPUT_FILENAME = "data/OPUS_data_model.xlsx"
DEFAULT_OUTPUT_FILENAME = "data/OPUS_data_model_enriched.xlsx"
DEFAULT_VALIDATION_FILENAME = "data/OPUS_data_model_validation_report.xlsx"
DEFAULT_CACHE_FILENAME = "data/wikidata_cache.json"

WIKIDATA_ENTITY_URL = "https://www.wikidata.org/wiki/Special:EntityData/{qid}.json"
USER_AGENT = "OPUS-Wars-of-the-Quill-Enrichment/1.0 (research data enrichment)"
REQUEST_TIMEOUT = 30
REQUEST_SLEEP = 0.15
MAX_RETRIES = 3

SHEET_ORDER = [
    "Letter manifestations",
    "People",
    "People Forms",
    "Places",
    "Places Forms",
    "Institutions",
    "Letter aggregations",
]

ID_COLUMNS = {
    "Letter manifestations": "letter_manifestation_ID",
    "People": "person_ID",
    "Places": "place_ID",
    "Institutions": "institution_ID",
    "Letter aggregations": "letter_aggregation_id",
}

LIST_LIKE_COLUMNS = {
    "author_ID": "|",
    "recipient_ID": "|",
    "people_mentioned_IDs": "|",
    "places_mentioned_IDs": "|",
    "keywords_manual": "|",
    "keywords_aggregated": "|",
}

#%%
# ============================================================================
# GENERIC HELPERS
# ============================================================================

def normalize_missing(value: Any) -> Any:
    if pd.isna(value):
        return pd.NA
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return pd.NA
        lowered = text.lower()
        if lowered in {"no data", "nan", "none", "null", "n/a", "na"}:
            return pd.NA
        return text
    return value



def df_strip_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df



def parse_pipe_list(value: Any) -> List[str]:
    if pd.isna(value):
        return []
    parts = [part.strip() for part in str(value).split("|")]
    return [part for part in parts if part]



def parse_semicolon_list(value: Any) -> List[str]:
    if pd.isna(value):
        return []
    parts = [part.strip() for part in str(value).split(";")]
    return [part for part in parts if part]



def unique_preserve_order(values: Iterable[Any]) -> List[Any]:
    seen = set()
    output = []
    for value in values:
        key = json.dumps(value, sort_keys=True, ensure_ascii=False) if isinstance(value, dict) else value
        if key not in seen:
            seen.add(key)
            output.append(value)
    return output



def safe_numeric(value: Any) -> Any:
    if pd.isna(value):
        return pd.NA
    try:
        return pd.to_numeric(value)
    except Exception:
        return value



def extract_qid_from_url(value: Any) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None

    # direct QID
    if re.fullmatch(r"Q\d+", text):
        return text

    # full wikidata url
    match = re.search(r"/wiki/(Q\d+)(?:$|[?#])", text)
    if match:
        return match.group(1)

    return None



def load_json(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)



def save_json(path: Path, payload: Dict[str, Any]) -> None:
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

#%%
# ============================================================================
# LOAD / CLEAN
# ============================================================================

def load_workbook_sheets(input_path: Path) -> Dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(input_path)
    data = {}
    for sheet_name in excel.sheet_names:
        data[sheet_name] = pd.read_excel(input_path, sheet_name=sheet_name)
        data[sheet_name] = df_strip_columns(data[sheet_name])
    return data



def clean_sheet(df: pd.DataFrame, id_col: Optional[str] = None) -> pd.DataFrame:
    df = df.copy()

    # pandas >=2.1 replacement for deprecated DataFrame.applymap
    df = df.map(normalize_missing)

    # drop fully empty rows/cols
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")

    # remove technical rows like "Completed letters:"
    if id_col and id_col in df.columns:
        id_series = df[id_col].astype("string")
        technical_mask = id_series.str.endswith(":", na=False)
        df = df.loc[~technical_mask].copy()

    # strip strings once again after filtering
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)

    df = df.reset_index(drop=True)
    return df



def clean_all_sheets(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    cleaned = {}
    for sheet_name, df in data.items():
        cleaned[sheet_name] = clean_sheet(df, ID_COLUMNS.get(sheet_name))
    return cleaned

#%%
# ============================================================================
# VALIDATION
# ============================================================================

def validate_references(data: Dict[str, pd.DataFrame]) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []

    people_ids = set(data["People"]["person_ID"].dropna().astype(str)) if "People" in data else set()
    place_ids = set(data["Places"]["place_ID"].dropna().astype(str)) if "Places" in data else set()
    institution_ids = set(data["Institutions"]["institution_ID"].dropna().astype(str)) if "Institutions" in data else set()
    lm_ids = set(data["Letter manifestations"]["letter_manifestation_ID"].dropna().astype(str)) if "Letter manifestations" in data else set()

    # Letter manifestations references
    if "Letter manifestations" in data:
        lm = data["Letter manifestations"].copy()
        for idx, row in lm.iterrows():
            row_id = row.get("letter_manifestation_ID")

            for col in ["author_ID", "recipient_ID", "people_mentioned_IDs"]:
                if col in row.index:
                    for ref in parse_pipe_list(row[col]):
                        if ref not in people_ids:
                            issues.append({
                                "sheet": "Letter manifestations",
                                "row_index": idx,
                                "record_id": row_id,
                                "column": col,
                                "invalid_reference": ref,
                                "expected_target_sheet": "People",
                                "issue_type": "missing_foreign_key",
                            })

            for col in ["origin_place_ID", "destination_place_ID"]:
                if col in row.index and pd.notna(row[col]):
                    ref = str(row[col])
                    if ref not in place_ids:
                        issues.append({
                            "sheet": "Letter manifestations",
                            "row_index": idx,
                            "record_id": row_id,
                            "column": col,
                            "invalid_reference": ref,
                            "expected_target_sheet": "Places",
                            "issue_type": "missing_foreign_key",
                        })

            if "places_mentioned_IDs" in row.index:
                for ref in parse_pipe_list(row["places_mentioned_IDs"]):
                    if ref not in place_ids:
                        issues.append({
                            "sheet": "Letter manifestations",
                            "row_index": idx,
                            "record_id": row_id,
                            "column": "places_mentioned_IDs",
                            "invalid_reference": ref,
                            "expected_target_sheet": "Places",
                            "issue_type": "missing_foreign_key",
                        })

            if "repository_ID" in row.index and pd.notna(row["repository_ID"]):
                ref = str(row["repository_ID"])
                if ref not in institution_ids:
                    issues.append({
                        "sheet": "Letter manifestations",
                        "row_index": idx,
                        "record_id": row_id,
                        "column": "repository_ID",
                        "invalid_reference": ref,
                        "expected_target_sheet": "Institutions",
                        "issue_type": "missing_foreign_key",
                    })

    # Forms references
    if "People Forms" in data:
        pf = data["People Forms"].copy()
        for idx, row in pf.iterrows():
            if pd.notna(row.get("letter_manifestation_ID")) and str(row["letter_manifestation_ID"]) not in lm_ids:
                issues.append({
                    "sheet": "People Forms",
                    "row_index": idx,
                    "record_id": row.get("person_ID"),
                    "column": "letter_manifestation_ID",
                    "invalid_reference": row.get("letter_manifestation_ID"),
                    "expected_target_sheet": "Letter manifestations",
                    "issue_type": "missing_foreign_key",
                })
            if pd.notna(row.get("person_ID")) and str(row["person_ID"]) not in people_ids:
                issues.append({
                    "sheet": "People Forms",
                    "row_index": idx,
                    "record_id": row.get("person_ID"),
                    "column": "person_ID",
                    "invalid_reference": row.get("person_ID"),
                    "expected_target_sheet": "People",
                    "issue_type": "missing_foreign_key",
                })

    if "Places Forms" in data:
        pf = data["Places Forms"].copy()
        for idx, row in pf.iterrows():
            if pd.notna(row.get("letter_manifestation_ID")) and str(row["letter_manifestation_ID"]) not in lm_ids:
                issues.append({
                    "sheet": "Places Forms",
                    "row_index": idx,
                    "record_id": row.get("place_ID"),
                    "column": "letter_manifestation_ID",
                    "invalid_reference": row.get("letter_manifestation_ID"),
                    "expected_target_sheet": "Letter manifestations",
                    "issue_type": "missing_foreign_key",
                })
            if pd.notna(row.get("place_ID")) and str(row["place_ID"]) not in place_ids:
                issues.append({
                    "sheet": "Places Forms",
                    "row_index": idx,
                    "record_id": row.get("place_ID"),
                    "column": "place_ID",
                    "invalid_reference": row.get("place_ID"),
                    "expected_target_sheet": "Places",
                    "issue_type": "missing_foreign_key",
                })

    if "Institutions" in data and "location_ID" in data["Institutions"].columns:
        inst = data["Institutions"].copy()
        for idx, row in inst.iterrows():
            if pd.notna(row.get("location_ID")) and str(row["location_ID"]) not in place_ids:
                issues.append({
                    "sheet": "Institutions",
                    "row_index": idx,
                    "record_id": row.get("institution_ID"),
                    "column": "location_ID",
                    "invalid_reference": row.get("location_ID"),
                    "expected_target_sheet": "Places",
                    "issue_type": "missing_foreign_key",
                })

    if "Letter aggregations" in data:
        la = data["Letter aggregations"].copy()
        for idx, row in la.iterrows():
            ref = row.get("letter_manifestation_ID")
            if pd.notna(ref) and str(ref) not in lm_ids:
                issues.append({
                    "sheet": "Letter aggregations",
                    "row_index": idx,
                    "record_id": row.get("letter_aggregation_id"),
                    "column": "letter_manifestation_ID",
                    "invalid_reference": ref,
                    "expected_target_sheet": "Letter manifestations",
                    "issue_type": "missing_foreign_key",
                })

    return issues

#%%
# ============================================================================
# WIKIDATA FETCH / PARSE
# ============================================================================

def get_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})
    return session



def fetch_wikidata_entity(qid: str, session: requests.Session, cache: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if qid in cache:
        return cache[qid]

    url = WIKIDATA_ENTITY_URL.format(qid=qid)
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(url, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            data = response.json()
            entity = data.get("entities", {}).get(qid)
            cache[qid] = entity
            time.sleep(REQUEST_SLEEP)
            return entity
        except Exception as exc:
            last_error = str(exc)
            time.sleep(0.75 * attempt)

    cache[qid] = {"_error": last_error}
    return None



def get_claim_values(entity: Dict[str, Any], prop: str) -> List[Dict[str, Any]]:
    if not entity:
        return []
    claims = entity.get("claims", {}).get(prop, [])
    out = []
    for claim in claims:
        mainsnak = claim.get("mainsnak", {})
        datavalue = mainsnak.get("datavalue")
        if datavalue is not None:
            out.append(datavalue.get("value"))
    return out



def get_first_time_year(entity: Dict[str, Any], prop: str) -> Optional[str]:
    values = get_claim_values(entity, prop)
    for value in values:
        if isinstance(value, dict) and "time" in value:
            time_value = value["time"]
            # '+1592-07-27T00:00:00Z' -> '1592'
            match = re.match(r"^[+-]?(\d{1,16})-", time_value)
            if match:
                return match.group(1)
    return None



def get_first_entity_id(entity: Dict[str, Any], prop: str) -> Optional[str]:
    values = get_claim_values(entity, prop)
    for value in values:
        if isinstance(value, dict):
            if "id" in value:
                return value["id"]
            if "numeric-id" in value:
                return f"Q{value['numeric-id']}"
    return None



def get_first_coordinate(entity: Dict[str, Any]) -> Dict[str, Optional[float]]:
    values = get_claim_values(entity, "P625")
    for value in values:
        if isinstance(value, dict):
            lat = value.get("latitude")
            lon = value.get("longitude")
            return {"latitude": lat, "longitude": lon}
    return {"latitude": None, "longitude": None}



def get_labels(entity: Dict[str, Any]) -> Dict[str, Optional[str]]:
    labels = entity.get("labels", {}) if entity else {}
    return {
        "label_en": labels.get("en", {}).get("value"),
        "label_pl": labels.get("pl", {}).get("value"),
    }



def get_aliases(entity: Dict[str, Any], lang: str = "en") -> Optional[str]:
    aliases = entity.get("aliases", {}).get(lang, []) if entity else []
    values = [item.get("value") for item in aliases if item.get("value")]
    return " | ".join(values) if values else None



def extract_entity_snapshot(qid: str, entity: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if not entity:
        return {
            "wikidata_qid": qid,
            "wikidata_label_en": pd.NA,
            "wikidata_label_pl": pd.NA,
            "wikidata_aliases_en": pd.NA,
            "wikidata_description_en": pd.NA,
            "wikidata_instance_of_qid": pd.NA,
            "wikidata_country_qid": pd.NA,
            "wikidata_birthplace_qid": pd.NA,
            "wikidata_deathplace_qid": pd.NA,
            "wikidata_birth_year": pd.NA,
            "wikidata_death_year": pd.NA,
            "wikidata_latitude": pd.NA,
            "wikidata_longitude": pd.NA,
        }

    labels = get_labels(entity)
    descriptions = entity.get("descriptions", {})
    coords = get_first_coordinate(entity)

    return {
        "wikidata_qid": qid,
        "wikidata_label_en": labels["label_en"],
        "wikidata_label_pl": labels["label_pl"],
        "wikidata_aliases_en": get_aliases(entity, "en"),
        "wikidata_description_en": descriptions.get("en", {}).get("value"),
        "wikidata_instance_of_qid": get_first_entity_id(entity, "P31"),
        "wikidata_country_qid": get_first_entity_id(entity, "P17"),
        "wikidata_birthplace_qid": get_first_entity_id(entity, "P19"),
        "wikidata_deathplace_qid": get_first_entity_id(entity, "P20"),
        "wikidata_birth_year": get_first_time_year(entity, "P569"),
        "wikidata_death_year": get_first_time_year(entity, "P570"),
        "wikidata_latitude": coords["latitude"],
        "wikidata_longitude": coords["longitude"],
    }

#%%
# ============================================================================
# ENRICHMENT PER SHEET
# ============================================================================

def enrich_people(df_people: pd.DataFrame, session: requests.Session, cache: Dict[str, Any]) -> pd.DataFrame:
    df = df_people.copy()
    df["wikidata_qid"] = df["wikidata_url"].map(extract_qid_from_url) if "wikidata_url" in df.columns else pd.NA

    snapshots = []
    for qid in df["wikidata_qid"].dropna().astype(str).unique():
        entity = fetch_wikidata_entity(qid, session, cache)
        snapshots.append(extract_entity_snapshot(qid, entity))

    snapshots_df = pd.DataFrame(snapshots)
    if not snapshots_df.empty:
        df = df.merge(snapshots_df, on="wikidata_qid", how="left")
    else:
        for col in [
            "wikidata_label_en", "wikidata_label_pl", "wikidata_aliases_en", "wikidata_description_en",
            "wikidata_instance_of_qid", "wikidata_country_qid", "wikidata_birthplace_qid",
            "wikidata_deathplace_qid", "wikidata_birth_year", "wikidata_death_year",
            "wikidata_latitude", "wikidata_longitude",
        ]:
            df[col] = pd.NA

    # fill birth/death from Wikidata only if workbook lacks them
    if "birthdate" in df.columns:
        df["birthdate_enriched"] = df["birthdate"].fillna(df["wikidata_birth_year"])
    if "deathdate" in df.columns:
        df["deathdate_enriched"] = df["deathdate"].fillna(df["wikidata_death_year"])

    return df



def enrich_places(df_places: pd.DataFrame, session: requests.Session, cache: Dict[str, Any]) -> pd.DataFrame:
    df = df_places.copy()
    df["wikidata_qid"] = df["wikidata_url"].map(extract_qid_from_url) if "wikidata_url" in df.columns else pd.NA

    snapshots = []
    for qid in df["wikidata_qid"].dropna().astype(str).unique():
        entity = fetch_wikidata_entity(qid, session, cache)
        snapshots.append(extract_entity_snapshot(qid, entity))

    snapshots_df = pd.DataFrame(snapshots)
    if not snapshots_df.empty:
        df = df.merge(snapshots_df, on="wikidata_qid", how="left")
    else:
        for col in [
            "wikidata_label_en", "wikidata_label_pl", "wikidata_aliases_en", "wikidata_description_en",
            "wikidata_instance_of_qid", "wikidata_country_qid", "wikidata_latitude", "wikidata_longitude",
        ]:
            df[col] = pd.NA

    # preserve manual coordinates when present
    if "latitude" in df.columns:
        df["latitude_enriched"] = df["latitude"].fillna(df["wikidata_latitude"])
    if "longitude" in df.columns:
        df["longitude_enriched"] = df["longitude"].fillna(df["wikidata_longitude"])

    return df



def enrich_institutions(df_inst: pd.DataFrame, session: requests.Session, cache: Dict[str, Any]) -> pd.DataFrame:
    df = df_inst.copy()
    df["wikidata_qid"] = df["wikidata_url"].map(extract_qid_from_url) if "wikidata_url" in df.columns else pd.NA

    snapshots = []
    for qid in df["wikidata_qid"].dropna().astype(str).unique():
        entity = fetch_wikidata_entity(qid, session, cache)
        snapshots.append(extract_entity_snapshot(qid, entity))

    snapshots_df = pd.DataFrame(snapshots)
    if not snapshots_df.empty:
        df = df.merge(snapshots_df, on="wikidata_qid", how="left")
    else:
        for col in [
            "wikidata_label_en", "wikidata_label_pl", "wikidata_aliases_en", "wikidata_description_en",
            "wikidata_instance_of_qid", "wikidata_country_qid", "wikidata_latitude", "wikidata_longitude",
        ]:
            df[col] = pd.NA

    return df

#%%
# ============================================================================
# RELATION TABLES FROM LETTER MANIFESTATIONS
# ============================================================================

def explode_relation_sheet(
    df_letters: pd.DataFrame,
    source_column: str,
    output_columns: List[str],
    delimiter: str = "|",
    relation_type: Optional[str] = None,
) -> pd.DataFrame:
    rows = []

    for _, row in df_letters.iterrows():
        lm_id = row.get("letter_manifestation_ID")
        values = parse_pipe_list(row.get(source_column)) if delimiter == "|" else parse_semicolon_list(row.get(source_column))
        for value in values:
            payload = {"letter_manifestation_ID": lm_id}
            payload[output_columns[1]] = value
            if relation_type is not None:
                payload["relation_type"] = relation_type
            rows.append(payload)

    out = pd.DataFrame(rows)
    if out.empty:
        out = pd.DataFrame(columns=output_columns + (["relation_type"] if relation_type is not None else []))
    else:
        # ensure requested column order
        ordered = ["letter_manifestation_ID", output_columns[1]]
        if relation_type is not None:
            ordered.append("relation_type")
        out = out[ordered]

    return out



def build_people_mentions_sheet(df_letters: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df_letters.iterrows():
        lm_id = row.get("letter_manifestation_ID")
        for pid in parse_pipe_list(row.get("people_mentioned_IDs")):
            rows.append({
                "letter_manifestation_ID": lm_id,
                "person_ID": pid,
                "relation_type": "mentioned_person",
            })
    return pd.DataFrame(rows)



def build_places_mentions_sheet(df_letters: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df_letters.iterrows():
        lm_id = row.get("letter_manifestation_ID")
        for place_id in parse_pipe_list(row.get("places_mentioned_IDs")):
            rows.append({
                "letter_manifestation_ID": lm_id,
                "place_ID": place_id,
                "relation_type": "mentioned_place",
            })
    return pd.DataFrame(rows)



def build_keywords_sheet(df_letters: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df_letters.iterrows():
        lm_id = row.get("letter_manifestation_ID")

        for keyword in parse_pipe_list(row.get("keywords_manual")):
            rows.append({
                "letter_manifestation_ID": lm_id,
                "keyword": keyword,
                "keyword_source": "manual",
            })

        for keyword in parse_pipe_list(row.get("keywords_aggregated")):
            rows.append({
                "letter_manifestation_ID": lm_id,
                "keyword": keyword,
                "keyword_source": "aggregated",
            })

    if not rows:
        return pd.DataFrame(columns=["letter_manifestation_ID", "keyword", "keyword_source"])
    return pd.DataFrame(rows)



def build_letter_role_sheet(df_letters: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df_letters.iterrows():
        lm_id = row.get("letter_manifestation_ID")

        for pid in parse_pipe_list(row.get("author_ID")):
            rows.append({
                "letter_manifestation_ID": lm_id,
                "person_ID": pid,
                "letter_role": "author",
            })

        for pid in parse_pipe_list(row.get("recipient_ID")):
            rows.append({
                "letter_manifestation_ID": lm_id,
                "person_ID": pid,
                "letter_role": "recipient",
            })

    if not rows:
        return pd.DataFrame(columns=["letter_manifestation_ID", "person_ID", "letter_role"])
    return pd.DataFrame(rows)



def build_enriched_tables(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    enriched = {name: df.copy() for name, df in data.items()}

    if "Letter manifestations" in data:
        letters = data["Letter manifestations"].copy()
        enriched["Letter roles"] = build_letter_role_sheet(letters)
        enriched["People mentions"] = build_people_mentions_sheet(letters)
        enriched["Place mentions"] = build_places_mentions_sheet(letters)
        enriched["Letter keywords"] = build_keywords_sheet(letters)

    return enriched

#%%
# ============================================================================
# SAVE
# ============================================================================

def save_workbook(data: Dict[str, pd.DataFrame], output_path: Path) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # keep core sheets first
        already_written = set()
        for sheet_name in SHEET_ORDER:
            if sheet_name in data:
                data[sheet_name].to_excel(writer, sheet_name=sheet_name[:31], index=False)
                already_written.add(sheet_name)

        for sheet_name, df in data.items():
            if sheet_name not in already_written:
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)



def save_validation_report(issues: List[Dict[str, Any]], output_path: Path) -> None:
    issues_df = pd.DataFrame(issues)
    if issues_df.empty:
        issues_df = pd.DataFrame(columns=[
            "sheet", "row_index", "record_id", "column",
            "invalid_reference", "expected_target_sheet", "issue_type",
        ])

    summary_df = (
        issues_df.groupby(["sheet", "issue_type"], dropna=False)
        .size()
        .reset_index(name="count")
        if not issues_df.empty
        else pd.DataFrame(columns=["sheet", "issue_type", "count"])
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        issues_df.to_excel(writer, sheet_name="issues", index=False)
        summary_df.to_excel(writer, sheet_name="summary", index=False)

#%%
# ============================================================================
# MAIN
# ============================================================================

def main() -> None:
    base_dir = Path(__file__).resolve().parent
    data_dir = base_dir / "data"
    input_path = data_dir / DEFAULT_INPUT_FILENAME
    output_path = data_dir / DEFAULT_OUTPUT_FILENAME
    validation_path = data_dir / DEFAULT_VALIDATION_FILENAME
    cache_path = data_dir / DEFAULT_CACHE_FILENAME

    if not input_path.exists():
        fallback = base_dir / DEFAULT_INPUT_FILENAME
        if fallback.exists():
            input_path = fallback
            output_path = base_dir / DEFAULT_OUTPUT_FILENAME
            validation_path = base_dir / DEFAULT_VALIDATION_FILENAME
            cache_path = base_dir / DEFAULT_CACHE_FILENAME
        else:
            raise FileNotFoundError(
                f"Nie znaleziono pliku wejściowego ani w {data_dir}, ani w {base_dir}: {DEFAULT_INPUT_FILENAME}"
            )

    print(f"[1/7] Wczytywanie: {input_path}")
    data_raw = load_workbook_sheets(input_path)

    print("[2/7] Czyszczenie arkuszy")
    data = clean_all_sheets(data_raw)

    print("[3/7] Walidacja relacji między arkuszami")
    validation_issues = validate_references(data)

    print("[4/7] Wczytywanie cache Wikidanych")
    cache = load_json(cache_path)
    session = get_session()

    print("[5/7] Wzbogacanie People / Places / Institutions")
    if "People" in data:
        data["People"] = enrich_people(data["People"], session, cache)
    if "Places" in data:
        data["Places"] = enrich_places(data["Places"], session, cache)
    if "Institutions" in data:
        data["Institutions"] = enrich_institutions(data["Institutions"], session, cache)

    print("[6/7] Budowanie dodatkowych tabel relacyjnych")
    enriched_data = build_enriched_tables(data)

    print(f"[7/7] Zapis: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    save_workbook(enriched_data, output_path)
    save_validation_report(validation_issues, validation_path)
    save_json(cache_path, cache)

    print("\nGotowe.")
    print(f"Nowy Excel: {output_path}")
    print(f"Raport walidacyjny: {validation_path}")
    print(f"Liczba problemów referencyjnych: {len(validation_issues)}")
    print(f"Liczba encji pobranych z cache/Wikidanych: {len(cache)}")


if __name__ == "__main__":
    main()


































