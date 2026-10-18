
"""Point d'entrée : python main.py [chemin_config.json]."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from excel_reader import (
    ConfigurationError,
    read_configured_files,
)
from excel_writer import write_result


def load_config(path: Path) -> dict[str, Any]:
    """Charge et valide la structure générale du fichier JSON."""

    try:
        with path.open("r", encoding="utf-8") as file:
            config = json.load(file)

    except FileNotFoundError as exc:
        raise FileNotFoundError(
            f"Fichier de configuration introuvable : {path}"
        ) from exc

    except json.JSONDecodeError as exc:
        raise ConfigurationError(
            f"JSON invalide dans {path}, ligne {exc.lineno}, "
            f"colonne {exc.colno} : {exc.msg}"
        ) from exc

    if not isinstance(config, dict):
        raise ConfigurationError(
            "La racine de config.json doit être un objet JSON."
        )

    return config


def close_workbooks(data: dict[str, Any] | None) -> None:
    """Ferme les classeurs source, même en cas d'erreur."""

    if not data:
        return

    for file_data in data.get("files", []):
        wb_values = file_data.get("workbook_values")
        wb_styles = file_data.get("workbook_styles")

        if wb_values is not None:
            wb_values.close()

        if wb_styles is not None and wb_styles is not wb_values:
            wb_styles.close()


def print_exclusion_details(data: dict[str, Any]) -> None:
    """Affiche les coordonnées et les causes des exclusions."""

    exclusions = data.get("exclusion_details", [])

    print("\n--- Détail des exclusions ---")

    if not exclusions:
        print("Aucune exclusion enregistrée.")
        return

    # Regroupement visuel par fichier source.
    for index, item in enumerate(exclusions, start=1):
        exclusion_type = str(item.get("type", "inconnu")).upper()
        source_file = item.get("source_file", "Fichier inconnu")
        source_sheet = item.get("source_sheet", "Feuille inconnue")
        source_row = item.get("source_row", "?")
        source_column = item.get("source_column", "?")
        source_cell = item.get(
            "source_cell",
            f"{source_column}{source_row}",
        )
        value = item.get("value")
        reason = item.get("reason", "Cause non précisée")

        print(f"\n[{index}] [{exclusion_type}]")
        print(f"    Fichier  : {source_file}")
        print(f"    Feuille  : {source_sheet}")
        print(f"    Ligne    : {source_row}")
        print(f"    Colonne  : {source_column}")
        print(f"    Cellule  : {source_sheet}!{source_cell}")
        print(f"    Valeur   : {value!r}")
        print(f"    Cause    : {reason}")


def print_summary(
    config: dict[str, Any],
    data: dict[str, Any],
    output_path: Path,
) -> None:
    """Affiche le bilan final de l'extraction."""

    file_01 = data["files"][0]
    file_02 = data["files"][1]

    print("\n")
    print("=" * 65)
    print("       EXCEL EXTRACTOR - BILAN DU TRAITEMENT")
    print("=" * 65)

    print(
        f"Fichier 1                 : "
        f"{config['file_01']['input_file']}"
    )
    print(
        f"Fichier 2                 : "
        f"{config['file_02']['input_file']}"
    )

    # ---------------------------------------------------------
    # Lignes exportées
    # ---------------------------------------------------------
    print("\n--- Lignes ---")

    print(
        "Lignes exportées depuis file_01 : "
        f"{len(file_01['records'])}"
    )
    print(
        "Lignes exportées depuis file_02 : "
        f"{len(file_02['records'])}"
    )
    print(
        "Total des lignes exportées      : "
        f"{len(data['records'])}"
    )

    # ---------------------------------------------------------
    # Résumé des exclusions
    # ---------------------------------------------------------
    print("\n--- Exclusions ---")

    print(
        "Lignes exclues pour motif 18 : "
        f"{data.get('excluded_motif_count', 0)}"
    )
    print(
        "Lignes exclues par filtre de dates : "
        f"{data.get('excluded_date_count', 0)}"
    )

    exclusions = data.get("exclusion_details", [])

    print(
        f"Total des exclusions détaillées : {len(exclusions)}"
    )

    # Afficher les coordonnées exactes de chaque exclusion.
    print_exclusion_details(data)

    # ---------------------------------------------------------
    # Cellules vides et remplacements
    # ---------------------------------------------------------
    print("\n--- Cellules vides et remplacements ---")

    print(
        f"Cellules remplacées par "
        f"{config['empty_value']!r} : "
        f"{data.get('empty_count', 0)}"
    )

    motif_empty_count = sum(
        1
        for warning in data.get("warnings", [])
        if warning.get("type") == "motif_vide"
    )

    print(
        "Cellules Motif Depart vides signalées : "
        f"{motif_empty_count}"
    )

    # ---------------------------------------------------------
    # Durées
    # ---------------------------------------------------------
    print("\n--- Durées ---")

    file_01_seconds = file_01.get("read_seconds", 0.0)
    file_02_seconds = file_02.get("read_seconds", 0.0)
    write_seconds = data.get("write_seconds", 0.0)

    print(
        f"Lecture fichier 1 : "
        f"{file_01_seconds:.3f} seconde(s)"
    )
    print(
        f"Lecture fichier 2 : "
        f"{file_02_seconds:.3f} seconde(s)"
    )
    print(
        f"Écriture résultat : "
        f"{write_seconds:.3f} seconde(s)"
    )

    total_read_seconds = file_01_seconds + file_02_seconds

    print(
        f"Total lecture des deux fichiers : "
        f"{total_read_seconds:.3f} seconde(s)"
    )

    print(
        f"Total lecture + écriture : "
        f"{total_read_seconds + write_seconds:.3f} seconde(s)"
    )

    # ---------------------------------------------------------
    # Fichiers générés
    # ---------------------------------------------------------
    journal_path = output_path.with_name(
        output_path.stem + "_journal.txt"
    )

    print("\n--- Fichiers générés ---")
    print(f"Résultat : {output_path}")
    print(f"Journal  : {journal_path}")

    print("=" * 65)


def main() -> int:
    """Exécute le traitement complet."""

    config_path = (
        Path(sys.argv[1])
        if len(sys.argv) > 1
        else Path("config.json")
    )

    data = None

    try:
        config = load_config(config_path)

        print("Début du traitement Excel...\n", flush=True)

        # Lecture et filtrage des fichiers sources.
        data = read_configured_files(config)

        # Création du fichier résultat et du journal.
        output_path = write_result(data, config)

        # Affichage du bilan.
        print_summary(config, data, output_path)

        return 0

    except (
        ConfigurationError,
        FileNotFoundError,
        FileExistsError,
        PermissionError,
        OSError,
        ValueError,
    ) as exc:
        print(f"\nERREUR : {exc}", file=sys.stderr)
        return 1

    except Exception as exc:
        print(
            f"\nERREUR INATTENDUE : "
            f"{type(exc).__name__}: {exc}",
            file=sys.stderr,
        )
        return 2

    finally:
        close_workbooks(data)


if __name__ == "__main__":
    raise SystemExit(main())
