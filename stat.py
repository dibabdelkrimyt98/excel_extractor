
"""Point d'entrée principal des statistiques Excel Extractor."""

import sys
import time
from pathlib import Path
from typing import Any, Callable, TypeVar

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from stat_utils import (
    add_age_column_to_input,
    load_config,
    load_employees,
)
from stat_sexe import calculate_gender_statistics
from stat_age import calculate_age_statistics
from stat_writer import write_statistics


T = TypeVar("T")


def timed_operation(
    label: str,
    operation: Callable[[], T],
) -> T:
    """Exécute une opération et affiche sa durée."""

    start = time.perf_counter()

    try:
        return operation()
    finally:
        elapsed = time.perf_counter() - start
        print(
            f"[TEMPS] {label} : {elapsed:.3f} seconde(s)",
            flush=True,
        )


def append_average_ages(
    output_file: str | Path,
    age_stats: dict[str, Any],
) -> float:
    """
    Ajoute les moyennes d'âge à la fin de la feuille
    'Tranches d'âge' dans le classeur de statistiques.

    Retourne le temps nécessaire à cette opération.
    """

    start = time.perf_counter()

    output_path = Path(output_file).expanduser().resolve()

    if not output_path.is_file():
        raise FileNotFoundError(
            f"Fichier de statistiques introuvable : {output_path}"
        )

    wb = None

    try:
        wb = load_workbook(output_path)

        sheet_name = "Tranches d'âge"

        if sheet_name not in wb.sheetnames:
            raise ValueError(
                f"La feuille {sheet_name!r} est absente du fichier "
                f"{output_path.name}. "
                f"Feuilles disponibles : {', '.join(wb.sheetnames)}"
            )

        ws = wb[sheet_name]

        # Laisser une ligne vide après les statistiques existantes.
        start_row = ws.max_row + 2

        # Styles du tableau récapitulatif.
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78",
        )
        label_font = Font(bold=True)
        thin_gray = Side(style="thin", color="D9E2F3")
        cell_border = Border(
            left=thin_gray,
            right=thin_gray,
            top=thin_gray,
            bottom=thin_gray,
        )

        # Titre du récapitulatif.
        ws.cell(
            row=start_row,
            column=1,
            value="RÉCAPITULATIF DES MOYENNES D'ÂGE",
        )

        ws.cell(
            row=start_row,
            column=1,
        ).font = Font(bold=True, size=12)

        ws.merge_cells(
            start_row=start_row,
            start_column=1,
            end_row=start_row,
            end_column=3,
        )

        # En-têtes.
        header_row = start_row + 1

        headers = [
            "Population",
            "Moyenne d'âge (ans)",
            "Nombre d'âges valides",
        ]

        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(
                row=header_row,
                column=col_idx,
                value=header,
            )
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )
            cell.border = cell_border

        averages = age_stats.get("average_ages", {})
        counts = age_stats.get("average_age_counts", {})

        populations = [
            ("Femme", "Femme"),
            ("Homme", "Homme"),
            ("Total", "Total"),
        ]

        for offset, (label, key) in enumerate(
            populations,
            start=1,
        ):
            row = header_row + offset
            average = averages.get(key)
            valid_count = counts.get(key, 0)

            ws.cell(
                row=row,
                column=1,
                value=label,
            )

            average_cell = ws.cell(
                row=row,
                column=2,
                value=round(average, 2) if average is not None else "N/A",
            )

            ws.cell(
                row=row,
                column=3,
                value=valid_count,
            )

            ws.cell(row=row, column=1).font = label_font

            for col_idx in range(1, 4):
                cell = ws.cell(row=row, column=col_idx)
                cell.border = cell_border
                cell.alignment = Alignment(
                    vertical="center",
                    horizontal=(
                        "left" if col_idx == 1 else "center"
                    ),
                )

            if average is not None:
                average_cell.number_format = "0.00"

        # Ajustement des largeurs des colonnes du récapitulatif.
        ws.column_dimensions["A"].width = max(
            ws.column_dimensions["A"].width or 0,
            22,
        )
        ws.column_dimensions["B"].width = max(
            ws.column_dimensions["B"].width or 0,
            24,
        )
        ws.column_dimensions["C"].width = max(
            ws.column_dimensions["C"].width or 0,
            25,
        )

        ws.freeze_panes = ws.freeze_panes or "A2"

        wb.save(output_path)

        print(
            f"[MOYENNES D'ÂGE] Résultats ajoutés dans "
            f"'{sheet_name}', lignes {start_row} à {header_row + 3}.",
            flush=True,
        )

    finally:
        if wb is not None:
            wb.close()

    return time.perf_counter() - start


def main() -> int:
    """Exécute le calcul des statistiques et mesure les durées."""

    total_start = time.perf_counter()

    try:
        # 1. Chargement de la configuration.
        config = timed_operation(
            "Chargement de stat_config.json",
            lambda: load_config("stat_config.json"),
        )

        # 2. Calcul et sauvegarde de la colonne AGE dans le fichier source.
        age_column_result = timed_operation(
            "Calcul et mise à jour de la colonne AGE",
            lambda: add_age_column_to_input(config),
        )

        print("\n=== Calcul de la colonne AGE ===")
        print(
            f"Colonne ajoutée/mise à jour : "
            f"{age_column_result['column']}"
        )
        print(
            f"Lignes traitées             : "
            f"{age_column_result['rows']}"
        )
        print(
            f"Dates absentes/invalides    : "
            f"{age_column_result['invalid_dates']}"
        )
        print(
            f"Sauvegarde initiale         : "
            f"{age_column_result['backup_path']}"
        )

        # 3. Chargement et préparation des employés.
        employees_data = timed_operation(
            "Chargement des données employés",
            lambda: load_employees(config),
        )

        # 4. Statistiques par sexe.
        gender_stats = timed_operation(
            "Calcul des statistiques par sexe",
            lambda: calculate_gender_statistics(employees_data),
        )

        # 5. Statistiques d'âge, tranches et moyennes.
        age_stats = timed_operation(
            "Calcul des statistiques d'âge et des moyennes",
            lambda: calculate_age_statistics(employees_data),
        )

        # 6. Écriture du classeur principal.
        output_path = timed_operation(
            "Génération du fichier de statistiques",
            lambda: write_statistics(
                gender_stats=gender_stats,
                age_stats=age_stats,
                output_file=config["output_file"],
                employees_data=employees_data,
            ),
        )

        # 7. Ajout des moyennes à la fin de la feuille ciblée.
        average_write_seconds = timed_operation(
            "Ajout des moyennes dans la feuille Tranches d'âge",
            lambda: append_average_ages(
                output_file=output_path,
                age_stats=age_stats,
            ),
        )

        # -----------------------------------------------------
        # Bilan final
        # -----------------------------------------------------
        print("\n" + "=" * 65)
        print("       EXCEL EXTRACTOR - STATISTIQUES")
        print("=" * 65)

        print(f"Employés analysés : {gender_stats['total']}")
        print(f"Femmes            : {gender_stats['counts']['Femme']}")
        print(f"Hommes            : {gender_stats['counts']['Homme']}")
        print(
            f"Civilités inconnues : "
            f"{gender_stats['counts']['VIDE']}"
        )

        print(
            f"Lignes exclues (Code Depart = 18) : "
            f"{age_stats['code_depart_filtered_count']}"
        )
        print(
            f"Dates de naissance absentes/invalides : "
            f"{age_stats['invalid_birth_dates']}"
        )
        print(
            f"Dates de naissance futures : "
            f"{age_stats['future_birth_dates']}"
        )
        print(
            f"Lignes entièrement vides ignorées : "
            f"{age_stats['empty_rows_skipped']}"
        )

        print(
            f"Date de référence : {config['reference_date']}"
        )

        # Affichage des moyennes calculées.
        averages = age_stats["average_ages"]
        counts = age_stats["average_age_counts"]

        print("\n--- Moyennes d'âge ---")

        for label in ("Femme", "Homme", "Total"):
            average = averages[label]
            count = counts[label]

            if average is None:
                print(
                    f"{label:<10} : N/A "
                    f"(aucun âge valide)"
                )
            else:
                print(
                    f"{label:<10} : {average:.2f} ans "
                    f"({count} âge(s) valide(s))"
                )

        print("\n--- Fichier généré ---")
        print(f"Fichier : {output_path}")

        print("\n--- Temps de traitement ---")
        total_elapsed = time.perf_counter() - total_start
        print(f"Temps total : {total_elapsed:.3f} seconde(s)")
        print(
            f"Temps d'ajout des moyennes : "
            f"{average_write_seconds:.3f} seconde(s)"
        )

        print("=" * 65)

        return 0

    except (
        FileNotFoundError,
        PermissionError,
        ValueError,
        OSError,
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


if __name__ == "__main__":
    raise SystemExit(main())
