from pathlib import Path
import argparse
import hashlib
import json
import re
from urllib.parse import quote

import pandas as pd
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, DCTERMS, PROV


# =========================================================
# KONFIGURACJA
# =========================================================

DEFAULT_BASE_GRAPH = "data/opus_knowledge_graph.ttl"
DEFAULT_EXCEL = "data/OPUS_data_model_enriched.xlsx"
DEFAULT_GCMA_DIR = "data/GCMA"
DEFAULT_OUTPUT = "data/opus_knowledge_graph_gcma.ttl"

BASE = "https://example.org/opus/"
OPUS = Namespace(BASE)
SCHEMA = Namespace("https://schema.org/")
GCMA = Namespace(BASE + "gcma/")
SKOS = Namespace("http://www.w3.org/2004/02/skos/core#")

LETTER_ID_COL = "letter_manifestation_ID"

GCMA_CLASSIFICATIONS = {"observed", "rumored", "inferred"}

MARKER_FIELDS = {
    "modal_verbs": OPUS.gcmaModalVerb,
    "evaluative_language": OPUS.gcmaEvaluativeLanguage,
    "syntax_and_conditional_constructions": OPUS.gcmaConditionalConstruction,
    "discourse_markers": OPUS.gcmaDiscourseMarker,
    "punctuation": OPUS.gcmaPunctuationMarker,
}


# =========================================================
# POMOCNICZE
# =========================================================

def get_script_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def resolve_path(path_value: str | Path) -> Path:
    path = Path(path_value)
    if not path.is_absolute():
        path = get_script_dir() / path
    return path


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Dodaje wyniki Grammar of Conspiracy Matrix Analysis (GCMA) "
            "do wcześniej zbudowanego grafu wiedzy OPUS."
        )
    )
    parser.add_argument(
        "--base-graph",
        default=DEFAULT_BASE_GRAPH,
        help="Ścieżka do istniejącego grafu RDF Turtle."
    )
    parser.add_argument(
        "--excel",
        default=DEFAULT_EXCEL,
        help="Ścieżka do Excela OPUS. Używany do walidacji listów i opcjonalnego mapowania osób."
    )
    parser.add_argument(
        "--gcma-dir",
        default=DEFAULT_GCMA_DIR,
        help="Folder z plikami *_GCMA.json."
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT,
        help="Ścieżka do zapisu wzbogaconego grafu RDF Turtle."
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Przerwij działanie, jeśli plik GCMA nie pasuje do żadnego letter_manifestation_ID."
    )
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


def safe_str(value):
    value = normalize_missing(value)
    if pd.isna(value):
        return None
    return str(value).strip()


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


def make_uri(entity_type: str, local_id: str) -> URIRef:
    local_id = safe_str(local_id)
    if not looks_like_source_id(local_id):
        raise ValueError(f"Niepoprawny identyfikator dla {entity_type}: {local_id}")
    return OPUS[f"{entity_type}/{quote(local_id, safe='')}"]


def make_gcma_uri(kind: str, local_id: str) -> URIRef:
    return GCMA[f"{kind}/{quote(local_id, safe='')}"]


def slugify(value: str, max_len: int = 80) -> str:
    value = safe_str(value) or "unknown"
    value = re.sub(r"[^A-Za-z0-9._-]+", "-", value).strip("-").lower()
    return value[:max_len] or "unknown"


def stable_hash(value: str, length: int = 12) -> str:
    return hashlib.sha1(value.encode("utf-8")).hexdigest()[:length]


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


def add_bool(g: Graph, s, p, value):
    if isinstance(value, bool):
        g.add((s, p, Literal(value, datatype=XSD.boolean)))


def as_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [v for v in value if safe_str(v)]
    if isinstance(value, str) and safe_str(value):
        return [value]
    return []


def read_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def infer_letter_id_from_gcma_filename(path: Path) -> str:
    name = path.name
    if name.endswith("_GCMA.json"):
        return name[:-len("_GCMA.json")]
    return path.stem


# =========================================================
# WALIDACJA I MAPOWANIE ENCJI
# =========================================================

def load_letter_ids_from_excel(excel_path: Path) -> set[str]:
    if not excel_path.exists():
        print(f"[WARN] Nie znaleziono Excela do walidacji: {excel_path}")
        return set()

    df = pd.read_excel(excel_path, sheet_name="Letter manifestations")
    if LETTER_ID_COL not in df.columns:
        print(f"[WARN] Brak kolumny {LETTER_ID_COL} w arkuszu Letter manifestations.")
        return set()

    ids = {
        safe_str(v)
        for v in df[LETTER_ID_COL].tolist()
        if safe_str(v) and looks_like_source_id(v)
    }
    return ids


def build_person_name_index(excel_path: Path) -> dict[str, set[str]]:
    """
    Zwraca indeks: znormalizowana forma nazwy -> zbiór person_ID.

    Używa zarówno arkusza People, jak i People Forms.
    Dopasowanie jest celowo ostrożne: exact match po lowercase.
    """
    if not excel_path.exists():
        return {}

    index: dict[str, set[str]] = {}

    def add_name(name, person_id):
        name = safe_str(name)
        person_id = safe_str(person_id)
        if name and person_id and looks_like_source_id(person_id):
            index.setdefault(name.lower(), set()).add(person_id)

    xls = pd.ExcelFile(excel_path)

    if "People" in xls.sheet_names:
        df_people = pd.read_excel(excel_path, sheet_name="People")
        for _, row in df_people.iterrows():
            add_name(row.get("person_name"), row.get("person_ID"))

    if "People Forms" in xls.sheet_names:
        df_forms = pd.read_excel(excel_path, sheet_name="People Forms")
        for _, row in df_forms.iterrows():
            add_name(row.get("person_name"), row.get("person_ID"))
            add_name(row.get("person_name_form"), row.get("person_ID"))

    return index


# =========================================================
# RDF: SCHEMA GCMA
# =========================================================

def bind_namespaces(g: Graph):
    g.bind("opus", OPUS)
    g.bind("gcma", GCMA)
    g.bind("schema", SCHEMA)
    g.bind("dcterms", DCTERMS)
    g.bind("prov", PROV)
    g.bind("skos", SKOS)
    g.bind("rdfs", RDFS)


def add_gcma_schema(g: Graph):
    """
    Minimalna warstwa modelu, aby wyniki GCMA były czytelne także bez dokumentacji zewnętrznej.
    """
    classes = {
        OPUS.GCMAAnalysis: "GCMA analysis",
        OPUS.GCMAExcerpt: "GCMA excerpt",
        OPUS.GCMASegment: "GCMA segment",
        OPUS.EvidentiaryMode: "GCMA evidentiary mode",
    }
    for cls, label in classes.items():
        g.add((cls, RDF.type, RDFS.Class))
        g.add((cls, RDFS.label, Literal(label, lang="en")))

    modes = {
        "observed": "Observed conspiracy",
        "rumored": "Rumored conspiracy",
        "inferred": "Inferred conspiracy",
    }
    for mode, label in modes.items():
        uri = make_gcma_uri("mode", mode)
        g.add((uri, RDF.type, OPUS.EvidentiaryMode))
        g.add((uri, SKOS.prefLabel, Literal(label, lang="en")))

    properties = {
        OPUS.hasGCMAAnalysis: "has GCMA analysis",
        OPUS.hasGCMAExcerpt: "has GCMA excerpt",
        OPUS.gcmaEvidentiaryMode: "GCMA evidentiary mode",
        OPUS.gcmaExcerptText: "GCMA excerpt text",
        OPUS.gcmaSegment: "GCMA segment",
        OPUS.gcmaSegmentLabel: "GCMA segment label",
        OPUS.gcmaClassification: "GCMA classification",
        OPUS.gcmaParanoiaDetected: "GCMA paranoia detected",
        OPUS.gcmaClassificationRationale: "GCMA classification rationale",
        OPUS.gcmaParanoiaRationale: "GCMA paranoia rationale",
        OPUS.gcmaObservedCount: "GCMA observed count",
        OPUS.gcmaRumoredCount: "GCMA rumored count",
        OPUS.gcmaInferredCount: "GCMA inferred count",
        OPUS.gcmaExcerptCount: "GCMA excerpt count",
        OPUS.gcmaSegmentCount: "GCMA segment count",
        OPUS.gcmaMarkerCount: "GCMA marker count",
        OPUS.gcmaPersonMentionedName: "GCMA person mentioned name",
        OPUS.gcmaSourceFile: "GCMA source file",
        OPUS.gcmaTextSource: "GCMA text source",
        OPUS.gcmaAuthorFromOutput: "GCMA author from output",
        OPUS.gcmaRecipientFromOutput: "GCMA recipient from output",
        OPUS.gcmaDatePlaceFromOutput: "GCMA date/place from output",
        OPUS.gcmaKeywordsFromOutput: "GCMA keywords from output",
        OPUS.gcmaAbstractFromOutput: "GCMA abstract from output",
    }
    for prop, label in properties.items():
        g.add((prop, RDF.type, RDF.Property))
        g.add((prop, RDFS.label, Literal(label, lang="en")))

    for _, prop in MARKER_FIELDS.items():
        g.add((prop, RDF.type, RDF.Property))


# =========================================================
# RDF: DODAWANIE WYNIKÓW GCMA
# =========================================================

def add_gcma_file(
    g: Graph,
    json_path: Path,
    valid_letter_ids: set[str],
    person_name_index: dict[str, set[str]],
    strict: bool = False,
):
    data = read_json(json_path)
    letter_id = infer_letter_id_from_gcma_filename(json_path)

    if valid_letter_ids and letter_id not in valid_letter_ids:
        msg = f"Plik {json_path.name} nie pasuje do żadnego {LETTER_ID_COL}: {letter_id}"
        if strict:
            raise ValueError(msg)
        print(f"[WARN] {msg}. Pomijam.")
        return False

    letter = make_uri("letter_manifestation", letter_id)

    analysis_local_id = f"{letter_id}-{stable_hash(json_path.name)}"
    analysis = make_gcma_uri("analysis", analysis_local_id)

    g.add((analysis, RDF.type, OPUS.GCMAAnalysis))
    g.add((letter, OPUS.hasGCMAAnalysis, analysis))
    g.add((analysis, PROV.wasDerivedFrom, letter))
    add_literal(g, analysis, OPUS.gcmaSourceFile, json_path.name)
    add_literal(g, analysis, OPUS.gcmaTextSource, data.get("text_source"))

    metadata = data.get("metadata", {}) or {}
    add_literal(g, analysis, SCHEMA.additionalType, metadata.get("document_type"))
    add_literal(g, analysis, OPUS.gcmaAuthorFromOutput, metadata.get("author"))
    add_literal(g, analysis, OPUS.gcmaRecipientFromOutput, metadata.get("recipient"))
    add_literal(g, analysis, OPUS.gcmaDatePlaceFromOutput, metadata.get("date_place"))
    add_literal(g, analysis, OPUS.gcmaKeywordsFromOutput, metadata.get("keywords"))
    add_literal(g, analysis, OPUS.gcmaAbstractFromOutput, metadata.get("abstract"))

    overview = data.get("analysis_overview", {}) or {}
    segments = as_list(overview.get("segments"))
    counts = overview.get("no_of_excerpts", {}) or {}
    excerpts = as_list(data.get("excerpts"))

    for mode in GCMA_CLASSIFICATIONS:
        value = counts.get(mode)
        if value is not None:
            prop = {
                "observed": OPUS.gcmaObservedCount,
                "rumored": OPUS.gcmaRumoredCount,
                "inferred": OPUS.gcmaInferredCount,
            }[mode]
            add_literal(g, analysis, prop, int(value), datatype=XSD.integer)

    add_literal(g, analysis, OPUS.gcmaExcerptCount, len(excerpts), datatype=XSD.integer)
    add_literal(g, analysis, OPUS.gcmaSegmentCount, len(segments), datatype=XSD.integer)

    # Segmenty jako osobne węzły: przydatne do późniejszych zapytań i wizualizacji.
    segment_uri_by_label = {}
    for segment in segments:
        segment_local_id = f"{letter_id}-{slugify(segment)}"
        segment_uri = make_gcma_uri("segment", segment_local_id)
        segment_uri_by_label[segment] = segment_uri
        g.add((segment_uri, RDF.type, OPUS.GCMASegment))
        g.add((analysis, OPUS.gcmaSegment, segment_uri))
        add_literal(g, segment_uri, OPUS.gcmaSegmentLabel, segment)
        g.add((segment_uri, PROV.wasDerivedFrom, analysis))

    markers = data.get("markers_of_conspiratorial_language", {}) or {}
    marker_total = 0
    for field, prop in MARKER_FIELDS.items():
        for marker in as_list(markers.get(field)):
            add_literal(g, analysis, prop, marker)
            marker_total += 1
    add_literal(g, analysis, OPUS.gcmaMarkerCount, marker_total, datatype=XSD.integer)

    # Fragmenty jako osobne węzły zależne od analizy.
    for i, excerpt in enumerate(excerpts, start=1):
        if not isinstance(excerpt, dict):
            continue

        excerpt_text = safe_str(excerpt.get("excerpt")) or ""
        classification = safe_str(excerpt.get("gcma_classification"))
        segment_label = safe_str(excerpt.get("segment"))

        local_seed = f"{letter_id}|{i}|{classification}|{excerpt_text}"
        excerpt_uri = make_gcma_uri("excerpt", f"{letter_id}-{i:03d}-{stable_hash(local_seed)}")

        g.add((excerpt_uri, RDF.type, OPUS.GCMAExcerpt))
        g.add((analysis, OPUS.hasGCMAExcerpt, excerpt_uri))
        g.add((letter, OPUS.hasGCMAExcerpt, excerpt_uri))
        g.add((excerpt_uri, PROV.wasDerivedFrom, analysis))

        add_literal(g, excerpt_uri, OPUS.gcmaExcerptText, excerpt_text)
        add_literal(g, excerpt_uri, OPUS.gcmaClassification, classification)
        add_literal(
            g,
            excerpt_uri,
            OPUS.gcmaClassificationRationale,
            excerpt.get("explanation_of_chosen_gcma_type"),
        )
        add_bool(g, excerpt_uri, OPUS.gcmaParanoiaDetected, excerpt.get("paranoia_detected"))
        add_literal(
            g,
            excerpt_uri,
            OPUS.gcmaParanoiaRationale,
            excerpt.get("explanation_of_paranoia_detection"),
        )

        if classification in GCMA_CLASSIFICATIONS:
            g.add((excerpt_uri, OPUS.gcmaEvidentiaryMode, make_gcma_uri("mode", classification)))

        if segment_label:
            segment_uri = segment_uri_by_label.get(segment_label)
            if segment_uri is None:
                segment_uri = make_gcma_uri("segment", f"{letter_id}-{slugify(segment_label)}")
                g.add((segment_uri, RDF.type, OPUS.GCMASegment))
                add_literal(g, segment_uri, OPUS.gcmaSegmentLabel, segment_label)
                segment_uri_by_label[segment_label] = segment_uri
            g.add((excerpt_uri, OPUS.gcmaSegment, segment_uri))

        for person_name in as_list(excerpt.get("people_mentioned_in_excerpt")):
            add_literal(g, excerpt_uri, OPUS.gcmaPersonMentionedName, person_name)
            for person_id in person_name_index.get(person_name.lower(), set()):
                g.add((excerpt_uri, OPUS.mentionsPerson, make_uri("person", person_id)))

    return True


def enrich_graph_with_gcma(
    base_graph_path: Path,
    excel_path: Path,
    gcma_dir: Path,
    output_path: Path,
    strict: bool = False,
):
    if not base_graph_path.exists():
        raise FileNotFoundError(f"Nie znaleziono bazowego grafu RDF: {base_graph_path}")

    if not gcma_dir.exists():
        raise FileNotFoundError(f"Nie znaleziono folderu z plikami GCMA: {gcma_dir}")

    gcma_files = sorted(gcma_dir.glob("*_GCMA.json"))
    if not gcma_files:
        raise FileNotFoundError(f"Brak plików *_GCMA.json w folderze: {gcma_dir}")

    print(f"[1/5] Wczytywanie grafu bazowego: {base_graph_path}")
    g = Graph()
    bind_namespaces(g)
    g.parse(base_graph_path, format="turtle")

    print(f"[2/5] Wczytywanie identyfikatorów i indeksu osób z Excela: {excel_path}")
    valid_letter_ids = load_letter_ids_from_excel(excel_path)
    person_name_index = build_person_name_index(excel_path)

    print("[3/5] Dodawanie minimalnego modelu RDF dla GCMA")
    add_gcma_schema(g)

    print(f"[4/5] Dodawanie wyników GCMA z folderu: {gcma_dir}")
    added = 0
    skipped = 0
    for json_path in gcma_files:
        ok = add_gcma_file(
            g=g,
            json_path=json_path,
            valid_letter_ids=valid_letter_ids,
            person_name_index=person_name_index,
            strict=strict,
        )
        if ok:
            added += 1
        else:
            skipped += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"[5/5] Serializacja wzbogaconego grafu: {output_path}")
    g.serialize(destination=output_path, format="turtle")

    print("Gotowe.")
    print(f"Pliki GCMA dodane: {added}")
    print(f"Pliki GCMA pominięte: {skipped}")
    print(f"Liczba trójek RDF po wzbogaceniu: {len(g)}")
    print(f"Plik wyjściowy: {output_path}")


def main():
    args = parse_args()

    enrich_graph_with_gcma(
        base_graph_path=resolve_path(args.base_graph),
        excel_path=resolve_path(args.excel),
        gcma_dir=resolve_path(args.gcma_dir),
        output_path=resolve_path(args.output),
        strict=args.strict,
    )


if __name__ == "__main__":
    main()
























