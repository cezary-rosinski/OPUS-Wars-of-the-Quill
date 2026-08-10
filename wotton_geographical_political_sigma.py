#!/usr/bin/env python3
"""Buduje sieć geograficzno-polityczną z lokalnego skoroszytu OPUS.

Eksport Sigma zapisuje stan wyłącznie bieżącego widżetu. Zapobiega to
powstawaniu wielusetmegabajtowych plików HTML w długotrwale działającym
kernelu Spydera lub Jupytera.

Oczekiwany układ katalogów:

    katalog_skryptu/
    ├── wotton_geographical_political_sigma_compact_html.py
    └── data/
        └── OPUS_data_model.xlsx

Wyniki są zapisywane w katalogu ``output`` obok skryptu.

Instalacja zależności:
    py -m pip install pandas openpyxl networkx ipysigma ipywidgets
"""

import gc
import json
import os
from collections import defaultdict
from itertools import product
from pathlib import Path

import networkx as nx
import pandas as pd

try:
    from ipysigma import Sigma
    from ipywidgets.embed import dependency_state, embed_minimal_html
except ImportError:
    Sigma = None
    dependency_state = None
    embed_minimal_html = None

import sys
sys.path.insert(1, r'C:\Users\pracownik\Documents\IBL-PAN-Python')
sys.path.insert(1, r'C:\Users\Cezary\Documents\IBL-PAN-Python')
from my_functions import gsheet_to_df

# ============================================================
# 1. ŚCIEŻKI I USTAWIENIA
# ============================================================

# Foldery ``data`` i ``output`` są wyznaczane względem położenia skryptu,
# a nie względem katalogu roboczego Spydera ani katalogu głównego dysku.
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("WOTTON_DATA_DIR", SCRIPT_DIR / "data")).resolve()
OUTPUT_DIR = Path(os.environ.get("WOTTON_OUTPUT_DIR", SCRIPT_DIR / "data/wotton_geographical_political_results")).resolve()

MIN_SUPPORT = 2
WEIGHT_ATTRIBUTE = "fractional_weight"  # albo: "raw_support"
LOUVAIN_RESOLUTION = 1.0
RANDOM_SEED = 42

# Dla tej sieci HTML powinien mieć najwyżej kilka MB. Limit chroni przed
# przypadkowym zapisaniem setek MB, gdyby mechanizm widżetów znów się zmienił.
MAX_HTML_SIZE_MB = 25

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

print(f"Katalog skryptu: {SCRIPT_DIR}")
print(f"Katalog danych:  {DATA_DIR}")
print(f"Katalog wyników: {OUTPUT_DIR}")


def write_sigma_html(graph, path, **options):
    """Zapisuje jeden widżet Sigma bez stanów z wcześniejszych uruchomień."""
    if Sigma is None or dependency_state is None or embed_minimal_html is None:
        raise RuntimeError("Do eksportu HTML potrzebne są ipysigma i ipywidgets.")

    path = Path(path)
    temporary_path = path.with_suffix(path.suffix + ".tmp")
    widget = None

    try:
        widget = Sigma(
            graph,
            height=None,
            raw_height="calc(100vh - 16px)",
            **options,
        )

        # dependency_state ogranicza zapis do tego widżetu i jego zależności.
        # Bez tego ipywidgets może dołączyć stany wszystkich widżetów z kernela.
        state = dependency_state([widget])
        embed_minimal_html(
            str(temporary_path),
            views=[widget],
            state=state,
            title=path.stem,
        )

        size_mb = temporary_path.stat().st_size / (1024 * 1024)
        if size_mb > MAX_HTML_SIZE_MB:
            raise RuntimeError(
                f"Eksport ma {size_mb:.1f} MB, więc przekracza limit "
                f"{MAX_HTML_SIZE_MB} MB. Plik nie zastąpił poprzedniej wersji."
            )

        # Zamiana następuje dopiero po pełnym i poprawnym zapisaniu HTML.
        temporary_path.replace(path)
        return size_mb

    finally:
        if widget is not None:
            widget.close()
        gc.collect()

        # Usuwamy wyłącznie niedokończony plik tymczasowy.
        if temporary_path.exists():
            temporary_path.unlink()


# ============================================================
# 2. ODCZYT SKOROSZYTU Z FOLDERU data
# ============================================================

letters = gsheet_to_df('1Yi6m0vEJndcmApy-y9vlsMiLayUaYJ48aGlHJN-69RM', 'Letter manifestations')
places = gsheet_to_df('1Yi6m0vEJndcmApy-y9vlsMiLayUaYJ48aGlHJN-69RM', 'Places')
political_entities = gsheet_to_df('1Yi6m0vEJndcmApy-y9vlsMiLayUaYJ48aGlHJN-69RM', 'Political entities')

required_letter_columns = {
    "letter_manifestation_ID",
    "for Barcelona",
    "places_mentioned_IDs",
    "political_entities_mentioned_IDs",
}
required_place_columns = {"place_ID", "place_name"}
required_political_columns = {"political_entity_ID", "entity_name"}

for table_name, table, required_columns in [
    ("Letter manifestations", letters, required_letter_columns),
    ("Places", places, required_place_columns),
    ("Political entities", political_entities, required_political_columns),
]:
    missing_columns = required_columns - set(table.columns)
    if missing_columns:
        raise ValueError(
            f"W arkuszu {table_name!r} brakuje kolumn: "
            + ", ".join(sorted(missing_columns))
        )

barcelona_mask = (
    letters["for Barcelona"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.casefold()
    .isin({"yes", "true", "1", "y"})
)
letters = letters.loc[barcelona_mask].copy()

if letters.empty:
    raise ValueError("Nie znaleziono listów z wartością 'for Barcelona = Yes'.")


def split_ids(value):
    """Rozdziela identyfikatory zapisane znakiem | i usuwa duplikaty."""
    if pd.isna(value):
        return []
    return sorted({part.strip() for part in str(value).split("|") if part.strip()})


def make_label_lookup(table, id_column, label_column):
    """Tworzy słownik ID -> nazwa wyświetlana."""
    lookup = {}
    for _, row in table.iterrows():
        if pd.isna(row[id_column]):
            continue
        source_id = str(row[id_column]).strip()
        label = source_id if pd.isna(row[label_column]) else str(row[label_column]).strip()
        lookup[source_id] = label
    return lookup


place_labels = make_label_lookup(places, "place_ID", "place_name")
political_labels = make_label_lookup(
    political_entities,
    "political_entity_ID",
    "entity_name",
)


# ============================================================
# 3. TWORZENIE KRAWĘDZI Z WSPÓŁWYSTĘPOWANIA W LISTACH
# ============================================================

node_letters = defaultdict(set)
edge_letters = defaultdict(set)
edge_fractional_weights = defaultdict(float)

for _, row in letters.iterrows():
    if pd.isna(row["letter_manifestation_ID"]):
        continue

    letter_id = str(row["letter_manifestation_ID"]).strip()
    place_ids = split_ids(row["places_mentioned_IDs"])
    political_ids = split_ids(row["political_entities_mentioned_IDs"])

    for place_id in place_ids:
        node_letters[f"geo:{place_id}"].add(letter_id)
    for political_id in political_ids:
        node_letters[f"pol:{political_id}"].add(letter_id)

    if not place_ids or not political_ids:
        continue

    fractional_contribution = 1 / (len(place_ids) * len(political_ids))

    for place_id, political_id in product(place_ids, political_ids):
        edge_id = (f"geo:{place_id}", f"pol:{political_id}")

        # Zabezpieczenie na wypadek powtórzonego rekordu tego samego listu.
        if letter_id not in edge_letters[edge_id]:
            edge_letters[edge_id].add(letter_id)
            edge_fractional_weights[edge_id] += fractional_contribution

edge_rows = []
for (source, target), letter_ids in edge_letters.items():
    raw_support = len(letter_ids)
    if raw_support < MIN_SUPPORT:
        continue

    fractional_weight = edge_fractional_weights[(source, target)]
    weight = fractional_weight if WEIGHT_ATTRIBUTE == "fractional_weight" else raw_support

    edge_rows.append(
        {
            "source": source,
            "target": target,
            "raw_support": raw_support,
            "fractional_weight": float(fractional_weight),
            "weight": float(weight),
            "letter_ids": " | ".join(sorted(letter_ids)),
            "label": f"{raw_support} shared letters",
        }
    )

edges = pd.DataFrame(
    edge_rows,
    columns=[
        "source",
        "target",
        "raw_support",
        "fractional_weight",
        "weight",
        "letter_ids",
        "label",
    ],
)

if edges.empty:
    raise ValueError(
        "Nie utworzono żadnej krawędzi. Sprawdź dane albo ustaw MIN_SUPPORT = 1."
    )


# ============================================================
# 4. TWORZENIE WĘZŁÓW
# ============================================================

visible_node_ids = set(edges["source"]) | set(edges["target"])
node_rows = []

for node_id in sorted(visible_node_ids):
    layer, source_id = node_id.split(":", maxsplit=1)

    if layer == "geo":
        label = place_labels.get(source_id, source_id)
        layer_name = "geographical"
        bipartite = 0
    else:
        label = political_labels.get(source_id, source_id)
        layer_name = "political"
        bipartite = 1

    node_rows.append(
        {
            "id": node_id,
            "label": label,
            "layer": layer_name,
            "source_id": source_id,
            "document_frequency": len(node_letters[node_id]),
            "bipartite": bipartite,
        }
    )

nodes = pd.DataFrame(node_rows)


# ============================================================
# 5. GRAF NETWORKX I MIARY SIECIOWE
# ============================================================

G = nx.Graph()

for node in nodes.sort_values("id").to_dict(orient="records"):
    node_id = node.pop("id")
    G.add_node(node_id, **node)

for edge in edges.sort_values(["source", "target"]).to_dict(orient="records"):
    source = edge.pop("source")
    target = edge.pop("target")
    G.add_edge(source, target, **edge)

nx.set_node_attributes(G, dict(G.degree()), "degree")
nx.set_node_attributes(
    G,
    {node: float(value) for node, value in G.degree(weight="weight")},
    "weighted_degree",
)
nx.set_node_attributes(G, nx.pagerank(G, weight="weight"), "pagerank")

communities = nx.community.louvain_communities(
    G,
    weight="weight",
    resolution=LOUVAIN_RESOLUTION,
    seed=RANDOM_SEED,
)
communities = sorted(communities, key=len, reverse=True)

louvain = {}
for community_number, community in enumerate(communities, start=1):
    for node_id in community:
        louvain[node_id] = community_number

nx.set_node_attributes(G, louvain, "louvain")

modularity = nx.community.modularity(
    G,
    communities,
    weight="weight",
    resolution=LOUVAIN_RESOLUTION,
)

for metric in ["degree", "weighted_degree", "pagerank", "louvain"]:
    values = nx.get_node_attributes(G, metric)
    nodes[metric] = nodes["id"].map(values)

node_label_by_id = nodes.set_index("id")["label"].to_dict()
edges["source_label"] = edges["source"].map(node_label_by_id)
edges["target_label"] = edges["target"].map(node_label_by_id)


# ============================================================
# 6. EKSPORT CSV, GRAPHML I PODSUMOWANIA
# ============================================================

nodes_path = OUTPUT_DIR / "nodes.csv"
edges_path = OUTPUT_DIR / "edges.csv"
graphml_path = OUTPUT_DIR / "network.graphml"
summary_path = OUTPUT_DIR / "summary.json"

nodes.sort_values(["pagerank", "label"], ascending=[False, True]).to_csv(
    nodes_path,
    index=False,
    encoding="utf-8-sig",
)

edges.sort_values(
    ["raw_support", "source_label", "target_label"],
    ascending=[False, True, True],
).to_csv(
    edges_path,
    index=False,
    encoding="utf-8-sig",
)

nx.write_graphml(G, graphml_path)

summary = {
    "filter": "for Barcelona = Yes",
    "letters": len(letters),
    "minimum_support": MIN_SUPPORT,
    "weight_attribute": WEIGHT_ATTRIBUTE,
    "nodes": G.number_of_nodes(),
    "geographical_nodes": int((nodes["layer"] == "geographical").sum()),
    "political_nodes": int((nodes["layer"] == "political").sum()),
    "edges": G.number_of_edges(),
    "communities": len(communities),
    "modularity": float(modularity),
}

summary_path.write_text(
    json.dumps(summary, ensure_ascii=False, indent=2),
    encoding="utf-8",
)

for path in [nodes_path, edges_path, graphml_path, summary_path]:
    print(f"Zapisano: {path}")


# ============================================================
# 7. EKSPORT SIGMA HTML
# ============================================================

if Sigma is None:
    print(
        "Pominięto eksport HTML, ponieważ brakuje ipysigma lub ipywidgets.\n"
        "Zainstaluj je poleceniem: py -m pip install ipysigma ipywidgets"
    )
else:
    sigma_options = {
        "node_size": "degree",
        "node_label_size": "degree",
        "node_size_range": (3, 20),
        "node_label_size_range": (9, 18),
        "node_shape": "layer",
        "edge_size": "weight",
        "edge_weight": "weight",
        "edge_size_range": (0.5, 5),
        "edge_size_scale": "sqrt",
        "default_edge_type": "curve",
        "default_node_label_size": 14,
        "hide_edges_on_move": True,
        "start_layout": 10,
    }

    html_exports = [
        (
            OUTPUT_DIR / "network_louvain.html",
            {
                "node_color": "louvain",
                "max_categorical_colors": 30,
            },
        ),
        (
            OUTPUT_DIR / "network_pagerank.html",
            {
                "node_color": "pagerank",
                "node_color_gradient": "Viridis",
                "node_color_scale": "sqrt",
            },
        ),
    ]

    for html_path, color_options in html_exports:
        try:
            size_mb = write_sigma_html(
                G,
                html_path,
                **color_options,
                **sigma_options,
            )
            print(f"Zapisano: {html_path} ({size_mb:.2f} MB)")
        except Exception as error:
            print(f"Nie udało się zapisać {html_path.name}: {error}")


print("\nPodsumowanie:")
print(json.dumps(summary, ensure_ascii=False, indent=2))
print(f"\nWszystkie wyniki znajdują się w:\n{OUTPUT_DIR}")
