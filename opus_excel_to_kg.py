from pathlib import Path
import argparse
import re
from urllib.parse import quote

import pandas as pd
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL

#%%
# =========================================================
# KONFIGURACJA
# =========================================================

DEFAULT_INPUT = "data/OPUS_data_model_enriched.xlsx"
DEFAULT_OUTPUT = "data/opus_knowledge_graph.ttl"

BASE = "https://example.org/opus/"
OPUS = Namespace(BASE)
SCHEMA = Namespace("https://schema.org/")
DCTERMS = Namespace("http://purl.org/dc/terms/")
GEO = Namespace("http://www.w3.org/2003/01/geo/wgs84_pos#")

EXPECTED_SHEETS = [
    "Letter manifestations",
    "People",
    "People Forms",
    "Places",
    "Places Forms",
    "Institutions",
    "Letter aggregations",
]

ENTITY_ID_COLUMNS = {
    "Letter manifestations": "letter_manifestation_ID",
    "People": "person_ID",
    "Places": "place_ID",
    "Institutions": "institution_ID",
    "Letter aggregations": "letter_aggregation_id",
}

# =========================================================
# POMOCNICZE
# =========================================================

def get_script_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Ścieżka do wzbogaconego Excela")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Ścieżka do pliku RDF Turtle")
    return parser.parse_args()


def normalize_missing(value):
    if pd.isna(value):
        return pd.NA
    if isinstance(value, str):
        value = value.strip()
        if value == "":
            return pd.NA
        if value.lower() in {"nan", "none", "null", "no data", "n/a"}:
            return pd.NA
    return value


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.map(normalize_missing)
    return df


def safe_str(value):
    value = normalize_missing(value)
    if pd.isna(value):
        return None
    return str(value).strip()


def split_pipe_values(value):
    value = normalize_missing(value)
    if pd.isna(value):
        return []
    return [part.strip() for part in str(value).split("|") if part.strip()]


def looks_like_source_id(value):
    value = safe_str(value)
    if not value:
        return False

    lowered = value.lower()
    forbidden_starts = [
        "basic keywords:",
        "keywords:",
        "people mentioned:",
        "places mentioned:",
        "notes:",
        "abstract:",
        "completed ",
    ]
    if any(lowered.startswith(prefix) for prefix in forbidden_starts):
        return False

    if "\n" in value or "\r" in value:
        return False

    return bool(re.fullmatch(r"[A-Za-z0-9._\-]+", value))


def report_invalid_ids(df: pd.DataFrame, id_col: str, label: str):
    if id_col not in df.columns:
        return

    bad = df[df[id_col].notna() & ~df[id_col].map(looks_like_source_id)]
    if not bad.empty:
        print(f"[WARN] Niepoprawne wartości w {label}.{id_col}:")
        print(bad[[id_col]].drop_duplicates().head(20).to_string(index=False))


def filter_valid_entity_rows(df: pd.DataFrame, id_col: str) -> pd.DataFrame:
    if id_col not in df.columns:
        return df.copy()
    mask = df[id_col].map(looks_like_source_id)
    return df[mask].copy()


def parse_date_literal(value):
    value = safe_str(value)
    if not value:
        return None

    try:
        # pełna data ISO lub podobna
        ts = pd.to_datetime(value, errors="raise")
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
            return Literal(ts.date().isoformat(), datatype=XSD.date)
        # sam rok
        if re.fullmatch(r"-?\d{1,4}", value):
            return Literal(str(int(value)), datatype=XSD.gYear)
    except Exception:
        pass

    # liczba w Excelu zapisana jako float/int
    try:
        f = float(value)
        if f.is_integer() and 0 < abs(f) <= 9999:
            return Literal(str(int(f)), datatype=XSD.gYear)
    except Exception:
        pass

    return Literal(value)


def make_uri(entity_type: str, local_id: str) -> URIRef:
    local_id = safe_str(local_id)
    if not looks_like_source_id(local_id):
        raise ValueError(f"Niepoprawny identyfikator dla {entity_type}: {local_id}")
    return OPUS[f"{entity_type}/{quote(local_id, safe='')}"]


def add_literal(g: Graph, s, p, value, datatype=None, lang=None):
    value = normalize_missing(value)
    if pd.isna(value):
        return
    if datatype:
        g.add((s, p, Literal(value, datatype=datatype)))
    elif lang:
        g.add((s, p, Literal(value, lang=lang)))
    else:
        g.add((s, p, Literal(value)))


def add_multivalue_literals(g: Graph, s, p, value):
    for item in split_pipe_values(value):
        g.add((s, p, Literal(item)))


def add_uri_if_valid(g: Graph, s, p, entity_type: str, local_id):
    local_id = safe_str(local_id)
    if local_id and looks_like_source_id(local_id):
        g.add((s, p, make_uri(entity_type, local_id)))


def load_sheets(input_path: Path):
    if not input_path.exists():
        raise FileNotFoundError(f"Nie znaleziono pliku: {input_path}")

    xls = pd.ExcelFile(input_path)
    missing = [sheet for sheet in EXPECTED_SHEETS if sheet not in xls.sheet_names]
    if missing:
        raise ValueError(f"Brakuje arkuszy w Excelu: {', '.join(missing)}")

    data = {}
    for sheet in EXPECTED_SHEETS:
        df = pd.read_excel(input_path, sheet_name=sheet)
        df = clean_dataframe(df)
        data[sheet] = df

    for sheet_name, id_col in ENTITY_ID_COLUMNS.items():
        report_invalid_ids(data[sheet_name], id_col, sheet_name)
        data[sheet_name] = filter_valid_entity_rows(data[sheet_name], id_col)

    return data


def coerce_numeric_literal(value, datatype):
    value = normalize_missing(value)
    if pd.isna(value):
        return None
    try:
        return Literal(float(value), datatype=datatype)
    except Exception:
        return None


# =========================================================
# RDF
# =========================================================

def init_graph() -> Graph:
    g = Graph()
    g.bind("opus", OPUS)
    g.bind("schema", SCHEMA)
    g.bind("dcterms", DCTERMS)
    g.bind("geo", GEO)
    g.bind("owl", OWL)
    g.bind("rdfs", RDFS)
    return g


def add_places(g: Graph, df_places: pd.DataFrame):
    for _, row in df_places.iterrows():
        place_id = safe_str(row.get("place_ID"))
        if not place_id:
            continue

        place = make_uri("place", place_id)
        g.add((place, RDF.type, SCHEMA.Place))
        add_literal(g, place, SCHEMA.identifier, place_id)
        add_literal(g, place, SCHEMA.name, row.get("place_name"))
        add_literal(g, place, RDFS.label, row.get("place_name"))
        add_literal(g, place, OPUS.notes, row.get("notes"))

        wd_url = safe_str(row.get("wikidata_url"))
        if wd_url:
            g.add((place, OWL.sameAs, URIRef(wd_url)))

        lat = row.get("latitude_final") if "latitude_final" in df_places.columns else row.get("latitude")
        lon = row.get("longitude_final") if "longitude_final" in df_places.columns else row.get("longitude")

        lat_lit = coerce_numeric_literal(lat, XSD.float)
        lon_lit = coerce_numeric_literal(lon, XSD.float)

        if lat_lit:
            g.add((place, GEO.lat, lat_lit))
        if lon_lit:
            g.add((place, GEO.long, lon_lit))

        add_literal(g, place, OPUS.wikidataLabel, row.get("wikidata_label"))
        add_literal(g, place, OPUS.wikidataDescription, row.get("wikidata_description"))
        add_literal(g, place, OPUS.placeType, row.get("type"))


def add_people(g: Graph, df_people: pd.DataFrame):
    for _, row in df_people.iterrows():
        person_id = safe_str(row.get("person_ID"))
        if not person_id:
            continue

        person = make_uri("person", person_id)
        g.add((person, RDF.type, SCHEMA.Person))
        add_literal(g, person, SCHEMA.identifier, person_id)
        add_literal(g, person, SCHEMA.name, row.get("person_name"))
        add_literal(g, person, RDFS.label, row.get("person_name"))
        add_literal(g, person, OPUS.roleTitle, row.get("role/title"))
        add_literal(g, person, OPUS.notes, row.get("notes"))

        wd_url = safe_str(row.get("wikidata_url"))
        if wd_url:
            g.add((person, OWL.sameAs, URIRef(wd_url)))

        birth_final = row.get("birthdate_final") if "birthdate_final" in df_people.columns else row.get("birthdate")
        death_final = row.get("deathdate_final") if "deathdate_final" in df_people.columns else row.get("deathdate")

        birth_lit = parse_date_literal(birth_final)
        death_lit = parse_date_literal(death_final)
        if birth_lit:
            g.add((person, SCHEMA.birthDate, birth_lit))
        if death_lit:
            g.add((person, SCHEMA.deathDate, death_lit))

        birthplace_id = row.get("birthplace_id_final") if "birthplace_id_final" in df_people.columns else row.get("birthplace_id")
        deathplace_id = row.get("deathplace_id_final") if "deathplace_id_final" in df_people.columns else row.get("deathplace_id")

        add_uri_if_valid(g, person, SCHEMA.birthPlace, "place", birthplace_id)
        add_uri_if_valid(g, person, SCHEMA.deathPlace, "place", deathplace_id)

        add_literal(g, person, OPUS.wikidataLabel, row.get("wikidata_label"))
        add_literal(g, person, OPUS.wikidataDescription, row.get("wikidata_description"))


def add_institutions(g: Graph, df_institutions: pd.DataFrame):
    for _, row in df_institutions.iterrows():
        institution_id = safe_str(row.get("institution_ID"))
        if not institution_id:
            continue

        institution = make_uri("institution", institution_id)
        g.add((institution, RDF.type, SCHEMA.Organization))
        add_literal(g, institution, SCHEMA.identifier, institution_id)
        add_literal(g, institution, SCHEMA.name, row.get("institution_name"))
        add_literal(g, institution, RDFS.label, row.get("institution_name"))
        add_literal(g, institution, SCHEMA.additionalType, row.get("type"))
        add_literal(g, institution, OPUS.notes, row.get("notes"))

        wd_url = safe_str(row.get("wikidata_url"))
        if wd_url:
            g.add((institution, OWL.sameAs, URIRef(wd_url)))

        add_uri_if_valid(g, institution, SCHEMA.location, "place", row.get("location_ID"))

        add_literal(g, institution, OPUS.wikidataLabel, row.get("wikidata_label"))
        add_literal(g, institution, OPUS.wikidataDescription, row.get("wikidata_description"))


def add_letter_manifestations(g: Graph, df_letters: pd.DataFrame):
    for _, row in df_letters.iterrows():
        letter_id = safe_str(row.get("letter_manifestation_ID"))
        if not letter_id or not looks_like_source_id(letter_id):
            continue

        letter = make_uri("letter_manifestation", letter_id)
        g.add((letter, RDF.type, OPUS.LetterManifestation))
        g.add((letter, RDF.type, SCHEMA.Message))
        add_literal(g, letter, SCHEMA.identifier, letter_id)
        add_literal(g, letter, SCHEMA.name, letter_id)

        add_literal(g, letter, OPUS.letterInCollection, row.get("letter_in_collection"))
        add_literal(g, letter, OPUS.manifestationType, row.get("print"))
        add_literal(g, letter, OPUS.dateOfLetterOld, row.get("date_of_letter_OLD"))
        add_literal(g, letter, OPUS.dateCertainty, row.get("date_of_letter_certainty"))
        add_literal(g, letter, OPUS.dateAsMarked, row.get("date_as_marked"))

        date_literal = parse_date_literal(row.get("date_of_letter"))
        if date_literal:
            g.add((letter, SCHEMA.dateCreated, date_literal))

        add_literal(g, letter, OPUS.authorAsMarked, row.get("author_as_marked"))
        add_literal(g, letter, OPUS.authorCertainty, row.get("author_certainty"))
        add_literal(g, letter, OPUS.recipientAsMarked, row.get("recipient_as_marked"))
        add_literal(g, letter, OPUS.recipientCertainty, row.get("recipient_certainty"))

        add_uri_if_valid(g, letter, SCHEMA.author, "person", row.get("author_ID"))
        add_uri_if_valid(g, letter, OPUS.recipient, "person", row.get("recipient_ID"))

        add_uri_if_valid(g, letter, SCHEMA.locationCreated, "place", row.get("origin_place_ID"))
        add_literal(g, letter, OPUS.originPlaceNameAsMarked, row.get("origin_place_name"))
        add_literal(g, letter, OPUS.originPlaceCertainty, row.get("origin_place_certainty"))

        add_uri_if_valid(g, letter, SCHEMA.contentLocation, "place", row.get("destination_place_ID"))
        add_literal(g, letter, OPUS.destinationPlaceNameAsMarked, row.get("destination_place_name"))
        add_literal(g, letter, OPUS.destinationPlaceCertainty, row.get("destination_place_certainty"))

        add_multivalue_literals(g, letter, SCHEMA.inLanguage, row.get("languages"))
        add_literal(g, letter, OPUS.incipit, row.get("incipit"))
        add_literal(g, letter, DCTERMS.abstract, row.get("abstract"))
        add_literal(g, letter, OPUS.peopleMentionedNamesRaw, row.get("people_mentioned_names"))
        add_literal(g, letter, OPUS.placesMentionedNamesRaw, row.get("places_mentioned_names"))

        # keywordy jako literały, nie URI
        add_multivalue_literals(g, letter, OPUS.keywordManual, row.get("keywords_manual"))
        add_multivalue_literals(g, letter, OPUS.keywordAggregated, row.get("keywords_aggregated"))

        add_literal(g, letter, OPUS.shelfmark, row.get("shelfmark"))
        add_literal(g, letter, OPUS.shelfmarkForOriginal, row.get("shelfmark_for_original"))
        add_literal(g, letter, OPUS.printInformation, row.get("print_information"))
        add_literal(g, letter, OPUS.letterURL, row.get("letter_url"))
        add_literal(g, letter, OPUS.transcription, row.get("transcription"))
        add_literal(g, letter, OPUS.notes, row.get("notes"))

        add_uri_if_valid(g, letter, DCTERMS.holder, "institution", row.get("repository_ID"))
        add_literal(g, letter, OPUS.repositoryNameAsMarked, row.get("repository_name"))

        # wzmianki po ID zostają relacjami URI, ale tylko jeśli ID są poprawne
        for person_id in split_pipe_values(row.get("people_mentioned_IDs")):
            add_uri_if_valid(g, letter, OPUS.mentionsPerson, "person", person_id)

        for place_id in split_pipe_values(row.get("places_mentioned_IDs")):
            add_uri_if_valid(g, letter, OPUS.mentionsPlace, "place", place_id)


def add_people_forms_as_literals(g: Graph, df_people_forms: pd.DataFrame):
    # formy nazw jako literały przypisane do listu, bez osobnych URI
    for _, row in df_people_forms.iterrows():
        letter_id = safe_str(row.get("letter_manifestation_ID"))
        if not letter_id or not looks_like_source_id(letter_id):
            continue

        letter = make_uri("letter_manifestation", letter_id)

        person_id = safe_str(row.get("person_ID"))
        if person_id and looks_like_source_id(person_id):
            g.add((letter, OPUS.personMentionedByForm, make_uri("person", person_id)))

        add_literal(g, letter, OPUS.personNameCanonical, row.get("person_name"))
        add_literal(g, letter, OPUS.personNameForm, row.get("person_name_form"))


def add_places_forms_as_literals(g: Graph, df_places_forms: pd.DataFrame):
    # formy nazw jako literały przypisane do listu, bez osobnych URI
    for _, row in df_places_forms.iterrows():
        letter_id = safe_str(row.get("letter_manifestation_ID"))
        if not letter_id or not looks_like_source_id(letter_id):
            continue

        letter = make_uri("letter_manifestation", letter_id)

        place_id = safe_str(row.get("place_ID"))
        if place_id and looks_like_source_id(place_id):
            g.add((letter, OPUS.placeMentionedByForm, make_uri("place", place_id)))

        add_literal(g, letter, OPUS.placeNameCanonical, row.get("place_name"))
        add_literal(g, letter, OPUS.placeNameForm, row.get("place_name_form"))


def add_letter_aggregations(g: Graph, df_agg: pd.DataFrame):
    required_cols = {"letter_aggregation_id", "letter_manifestation_ID"}
    if not required_cols.issubset(df_agg.columns):
        print("[WARN] Arkusz 'Letter aggregations' nie ma oczekiwanych kolumn.")
        return

    df_work = df_agg.copy()
    df_work["letter_aggregation_id"] = df_work["letter_aggregation_id"].map(safe_str)
    df_work["letter_manifestation_ID"] = df_work["letter_manifestation_ID"].map(safe_str)

    df_work = df_work[
        df_work["letter_aggregation_id"].map(looks_like_source_id)
    ].copy()

    for agg_id, sub in df_work.groupby("letter_aggregation_id", dropna=True):
        aggregation = make_uri("letter_aggregation", agg_id)
        g.add((aggregation, RDF.type, OPUS.LetterAggregation))
        add_literal(g, aggregation, SCHEMA.identifier, agg_id)
        add_literal(g, aggregation, SCHEMA.name, agg_id)

        for _, row in sub.iterrows():
            letter_id = safe_str(row.get("letter_manifestation_ID"))
            if letter_id and looks_like_source_id(letter_id):
                g.add((aggregation, OPUS.hasManifestation, make_uri("letter_manifestation", letter_id)))
            elif letter_id:
                add_literal(g, aggregation, OPUS.hasManifestationAsText, letter_id)


# =========================================================
# GŁÓWNA LOGIKA
# =========================================================

def main():
    args = parse_args()

    input_path = Path(args.input)
    if not input_path.is_absolute():
        input_path = get_script_dir() / input_path

    output_path = Path(args.output)
    if not output_path.is_absolute():
        output_path = get_script_dir() / output_path
    output_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"[1/4] Wczytywanie danych: {input_path}")
    data = load_sheets(input_path)

    print("[2/4] Inicjalizacja grafu RDF")
    g = init_graph()

    print("[3/4] Dodawanie encji i relacji")
    add_places(g, data["Places"])
    add_people(g, data["People"])
    add_institutions(g, data["Institutions"])
    add_letter_manifestations(g, data["Letter manifestations"])
    add_people_forms_as_literals(g, data["People Forms"])
    add_places_forms_as_literals(g, data["Places Forms"])
    add_letter_aggregations(g, data["Letter aggregations"])

    print(f"[4/4] Serializacja: {output_path}")
    g.serialize(destination=output_path, format="turtle")

    print("Gotowe.")
    print(f"Liczba trójek RDF: {len(g)}")
    print(f"Plik wyjściowy: {output_path}")


if __name__ == "__main__":
    main()
