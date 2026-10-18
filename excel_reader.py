
"""Lecture, validation et filtrage des fichiers Excel configurés."""
from __future__ import annotations

import time
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string


class ConfigurationError(ValueError):
    """Configuration JSON invalide ou incohérente."""


def validate_config(config: dict[str, Any]) -> None:
    """Vérifie la cohérence de la configuration JSON."""

    required = (
        "output_file",
        "overwrite_output",
        "empty_value",
        "formula_mode",
        "file_01",
        "file_02",
    )

    missing = [key for key in required if key not in config]

    if missing:
        raise ConfigurationError(
            "Paramètre(s) manquant(s) dans config.json : "
            + ", ".join(missing)
        )

    if not isinstance(config["overwrite_output"], bool):
        raise ConfigurationError(
            "'overwrite_output' doit être true ou false."
        )

    if config["formula_mode"] not in {
        "cached_value",
        "formula_text",
    }:
        raise ConfigurationError(
            "'formula_mode' doit être 'cached_value' "
            "ou 'formula_text'."
        )

    # Validation des deux fichiers sources.
    for file_key in ("file_01", "file_02"):
        section = config[file_key]

        if not isinstance(section, dict):
            raise ConfigurationError(
                f"'{file_key}' doit être un objet JSON."
            )

        for key in (
            "input_file",
            "sheet_name",
            "start_row",
            "columns",
        ):
            if key not in section:
                raise ConfigurationError(
                    f"Paramètre '{key}' manquant dans '{file_key}'."
                )

        if not isinstance(section["input_file"], str) or not section[
            "input_file"
        ].strip():
            raise ConfigurationError(
                f"'{file_key}.input_file' doit être un chemin valide."
            )

        if not isinstance(section["sheet_name"], str) or not section[
            "sheet_name"
        ].strip():
            raise ConfigurationError(
                f"'{file_key}.sheet_name' doit être renseigné."
            )

        start_row = section["start_row"]

        if (
            not isinstance(start_row, int)
            or isinstance(start_row, bool)
            or start_row < 1
        ):
            raise ConfigurationError(
                f"'{file_key}.start_row' doit être un entier >= 1."
            )

        columns = section["columns"]

        if not isinstance(columns, dict) or not columns:
            raise ConfigurationError(
                f"'{file_key}.columns' doit associer les champs "
                "aux colonnes Excel."
            )

        for field_name, col in columns.items():
            if (
                not isinstance(field_name, str)
                or not field_name.strip()
            ):
                raise ConfigurationError(
                    f"Nom de champ invalide dans '{file_key}.columns'."
                )

            if not isinstance(col, str) or not col.strip().isalpha():
                raise ConfigurationError(
                    f"Colonne invalide pour "
                    f"'{file_key}.{field_name}' : {col!r}."
                )

            try:
                column_index_from_string(col.strip().upper())
            except ValueError as exc:
                raise ConfigurationError(
                    f"Colonne Excel invalide pour "
                    f"'{file_key}.{field_name}' : {col!r}."
                ) from exc

    # Validation du filtre des dates de départ.
    date_filter = config["file_02"].get("date_filter")

    if date_filter is not None:
        if not isinstance(date_filter, dict):
            raise ConfigurationError(
                "'file_02.date_filter' doit être un objet JSON."
            )

        for key in (
            "column",
            "start_exclusive",
            "end_exclusive",
            "format",
        ):
            if key not in date_filter:
                raise ConfigurationError(
                    f"Paramètre '{key}' manquant dans "
                    "'file_02.date_filter'."
                )

        date_column = date_filter["column"]

        if (
            not isinstance(date_column, str)
            or not date_column.strip().isalpha()
        ):
            raise ConfigurationError(
                "'file_02.date_filter.column' "
                "doit être une colonne Excel."
            )

        try:
            column_index_from_string(date_column.strip().upper())
        except ValueError as exc:
            raise ConfigurationError(
                f"Colonne Excel invalide pour le filtre de dates : "
                f"{date_column!r}."
            ) from exc

        for key in ("start_exclusive", "end_exclusive"):
            try:
                datetime.strptime(date_filter[key], "%Y-%m-%d")
            except (TypeError, ValueError) as exc:
                raise ConfigurationError(
                    f"'file_02.date_filter.{key}' "
                    "doit respecter le format YYYY-MM-DD."
                ) from exc

        try:
            start_date = datetime.strptime(
                date_filter["start_exclusive"], "%Y-%m-%d"
            ).date()

            end_date = datetime.strptime(
                date_filter["end_exclusive"], "%Y-%m-%d"
            ).date()
        except (TypeError, ValueError) as exc:
            raise ConfigurationError(
                "Les bornes du filtre de dates sont invalides."
            ) from exc

        if start_date >= end_date:
            raise ConfigurationError(
                "'start_exclusive' doit être antérieure "
                "à 'end_exclusive'."
            )

        if (
            not isinstance(date_filter["format"], str)
            or not date_filter["format"].strip()
        ):
            raise ConfigurationError(
                "'file_02.date_filter.format' doit être renseigné."
            )

    # Validation du filtre du motif de départ.
    motif_filter = config["file_02"].get("motif_filter")

    if motif_filter is not None:
        if not isinstance(motif_filter, dict):
            raise ConfigurationError(
                "'file_02.motif_filter' doit être un objet JSON."
            )

        for key in ("field", "column", "exclude_value"):
            if key not in motif_filter:
                raise ConfigurationError(
                    f"Paramètre '{key}' manquant dans "
                    "'file_02.motif_filter'."
                )

        motif_field = motif_filter["field"]

        if (
            not isinstance(motif_field, str)
            or not motif_field.strip()
        ):
            raise ConfigurationError(
                "'file_02.motif_filter.field' doit être renseigné."
            )

        if motif_field not in config["file_02"]["columns"]:
            raise ConfigurationError(
                f"Le champ motif '{motif_field}' est absent "
                "de 'file_02.columns'."
            )

        motif_column = motif_filter["column"]

        if (
            not isinstance(motif_column, str)
            or not motif_column.strip().isalpha()
        ):
            raise ConfigurationError(
                "'file_02.motif_filter.column' "
                "doit être une colonne Excel."
            )

        try:
            column_index_from_string(
                motif_column.strip().upper()
            )
        except ValueError as exc:
            raise ConfigurationError(
                f"Colonne Excel invalide pour le motif de départ : "
                f"{motif_column!r}."
            ) from exc

        # Empêche le filtre et l'export de lire deux colonnes différentes.
        export_column = config["file_02"]["columns"][motif_field]

        if motif_column.strip().upper() != export_column.strip().upper():
            raise ConfigurationError(
                "Incohérence de configuration : "
                f"'motif_filter.column' vaut {motif_column!r}, "
                f"mais le champ '{motif_field}' est associé à "
                f"la colonne {export_column!r} dans "
                "'file_02.columns'. Les deux colonnes doivent "
                "être identiques."
            )

        if is_empty(motif_filter["exclude_value"]):
            raise ConfigurationError(
                "'file_02.motif_filter.exclude_value' "
                "ne peut pas être vide."
            )


def is_empty(value: Any) -> bool:
    """Retourne True si la valeur est vide ou ne contient que des espaces."""
    return value is None or (
        isinstance(value, str) and value.strip() == ""
    )


def _parse_date(value: Any, date_format: str) -> date | None:
    """Convertit une valeur Excel ou texte en date."""

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if is_empty(value):
        return None

    if isinstance(value, str):
        value = value.strip()

        formats = [
            date_format,
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%m/%d/%Y",
            "%d-%m-%Y",
            "%m-%d-%Y",
        ]

        for fmt in dict.fromkeys(formats):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue

    return None


def _last_data_row(
    ws: Any,
    start_row: int,
    columns: list[str],
) -> int:
    """Dernière ligne ayant au moins une valeur dans les colonnes utiles."""

    last_row = start_row - 1

    for row_idx in range(start_row, ws.max_row + 1):
        if any(
            not is_empty(ws[f"{col}{row_idx}"].value)
            for col in columns
        ):
            last_row = row_idx

    return last_row


def _matches_excluded_value(
    value: Any,
    excluded_value: Any,
) -> bool:
    """
    Compare une valeur avec le motif à exclure.

    Accepte notamment 18, '18', ' 18 ' et 18.0.
    """

    if is_empty(value):
        return False

    value_text = str(value).strip()
    excluded_text = str(excluded_value).strip()

    if value_text == excluded_text:
        return True

    try:
        return (
            float(value_text) == float(excluded_text)
            and value_text != ""
            and excluded_text != ""
        )
    except (TypeError, ValueError):
        return False


def read_one_file(
    file_key: str,
    section: dict[str, Any],
    config: dict[str, Any],
) -> dict[str, Any]:
    """Lit un fichier et retourne les lignes retenues et métadonnées."""

    read_start = time.perf_counter()

    input_path = Path(section["input_file"]).expanduser().resolve()

    if not input_path.is_file():
        raise FileNotFoundError(
            f"Fichier source introuvable pour {file_key} : "
            f"{input_path}"
        )

    if input_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ConfigurationError(
            f"Le fichier de {file_key} doit être au format "
            f".xlsx ou .xlsm : {input_path}"
        )

    formula_mode = config.get("formula_mode", "cached_value")

    wb_values = None
    wb_styles = None

    try:
        wb_values = load_workbook(
            input_path,
            data_only=(formula_mode == "cached_value"),
            read_only=False,
        )

        if formula_mode == "cached_value":
            wb_styles = load_workbook(
                input_path,
                data_only=False,
                read_only=False,
            )
        else:
            wb_styles = wb_values

        sheet_name = section["sheet_name"]

        if sheet_name not in wb_values.sheetnames:
            raise ConfigurationError(
                f"Feuille {sheet_name!r} introuvable dans "
                f"{input_path.name}. "
                f"Feuilles disponibles : "
                f"{', '.join(wb_values.sheetnames)}"
            )

        ws_values = wb_values[sheet_name]
        ws_styles = wb_styles[sheet_name]

        fields = {
            name: col.strip().upper()
            for name, col in section["columns"].items()
        }

        source_columns = list(dict.fromkeys(fields.values()))
        start_row = section["start_row"]

        # Recherche de la dernière ligne utile.
        last_row = _last_data_row(
            ws_values,
            start_row,
            source_columns,
        )

        records: list[dict[str, Any]] = []
        change_log: list[dict[str, Any]] = []
        warnings: list[dict[str, Any]] = []
        exclusion_details: list[dict[str, Any]] = []

        empty_count = 0
        excluded_motif_count = 0
        excluded_date_count = 0

        # ---------------------------------------------------------
        # Configuration du filtre des dates
        # ---------------------------------------------------------
        date_filter = (
            section.get("date_filter")
            if file_key == "file_02"
            else None
        )

        date_column = (
            date_filter["column"].strip().upper()
            if date_filter
            else None
        )

        start_date = (
            datetime.strptime(
                date_filter["start_exclusive"], "%Y-%m-%d"
            ).date()
            if date_filter
            else None
        )

        end_date = (
            datetime.strptime(
                date_filter["end_exclusive"], "%Y-%m-%d"
            ).date()
            if date_filter
            else None
        )

        # ---------------------------------------------------------
        # Configuration du filtre des motifs
        # ---------------------------------------------------------
        motif_filter = (
            section.get("motif_filter")
            if file_key == "file_02"
            else None
        )

        motif_field = (
            motif_filter["field"]
            if motif_filter
            else None
        )

        # IMPORTANT : la colonne utilisée pour filtrer est celle
        # explicitement déclarée dans motif_filter.column.
        motif_column = (
            motif_filter["column"].strip().upper()
            if motif_filter
            else None
        )

        excluded_motif_value = (
            motif_filter["exclude_value"]
            if motif_filter
            else None
        )

        # ---------------------------------------------------------
        # Lecture des lignes
        # ---------------------------------------------------------
        for row_idx in range(start_row, last_row + 1):

            # 1. Filtre des dates de départ.
            if date_filter:
                raw_date = ws_values[
                    f"{date_column}{row_idx}"
                ].value

                parsed_date = _parse_date(
                    raw_date,
                    date_filter["format"],
                )

                if (
                    parsed_date is None
                    or not (
                        start_date < parsed_date <= end_date
                    )
                ):
                    excluded_date_count += 1

                    message = (
                        f"[EXCLUSION][FILTRE DATE] "
                        f"{input_path.name} | "
                        f"{sheet_name}!{date_column}{row_idx} | "
                        f"valeur={raw_date!r} | "
                        f"date attendue strictement entre "
                        f"{start_date} et {end_date}"
                    )

                    print(message, flush=True)

                    warnings.append({
                        "type": "exclusion_date",
                        "message": message,
                        "source_row": row_idx,
                        "source_column": date_column,
                    })

                    exclusion_details.append({
                        "type": "date",
                        "source_file": input_path.name,
                        "source_sheet": sheet_name,
                        "source_row": row_idx,
                        "source_column": date_column,
                        "source_cell": f"{date_column}{row_idx}",
                        "value": raw_date,
                        "reason": (
                            "Date vide ou invalide"
                            if parsed_date is None
                            else (
                                f"Date hors période stricte : "
                                f"{start_date} < date < {end_date}"
                            )
                        ),
                    })
                    continue

            # 2. Filtre du motif de départ.
            if motif_filter:
                motif_cell = ws_values[
                    f"{motif_column}{row_idx}"
                ]
                motif_value = motif_cell.value

                if is_empty(motif_value):
                    # Le motif vide n'entraîne pas l'exclusion.
                    # La ligne est conservée et l'emplacement signalé.
                    message = (
                        f"[AVERTISSEMENT][MOTIF VIDE] "
                        f"{input_path.name} | "
                        f"{sheet_name}!{motif_column}{row_idx} | "
                        f"ligne={row_idx}, "
                        f"colonne={motif_column} | "
                        "Motif de départ vide ; ligne conservée."
                    )

                    print(message, flush=True)

                    warnings.append({
                        "type": "motif_vide",
                        "source_file": input_path.name,
                        "source_sheet": sheet_name,
                        "source_row": row_idx,
                        "source_column": motif_column,
                        "message": message,
                    })

                elif _matches_excluded_value(
                    motif_value,
                    excluded_motif_value,
                ):
                    excluded_motif_count += 1

                    message = (
                        f"[EXCLUSION][MOTIF DEPART] "
                        f"{input_path.name} | "
                        f"{sheet_name}!{motif_column}{row_idx} | "
                        f"valeur={motif_value!r} | "
                        f"motif exclu={excluded_motif_value!r}"
                    )

                    print(message, flush=True)

                    warnings.append({
                        "type": "exclusion_motif",
                        "source_file": input_path.name,
                        "source_sheet": sheet_name,
                        "source_row": row_idx,
                        "source_column": motif_column,
                        "value": motif_value,
                        "message": message,
                    })
                    exclusion_details.append({
                        "type": "motif",
                        "source_file": input_path.name,
                        "source_sheet": sheet_name,
                        "source_row": row_idx,
                        "source_column": motif_column,
                        "source_cell": f"{motif_column}{row_idx}",
                        "value": motif_value,
                        "reason": (
                            f"Motif de départ égal à "
                            f"{excluded_motif_value!r}"
                        ),
                    })
                    continue

            # 3. Lecture des champs des lignes conservées.
            values_by_field: dict[str, Any] = {}
            styles_by_field: dict[str, Any] = {}

            for field_name, source_col in fields.items():
                value_cell = ws_values[
                    f"{source_col}{row_idx}"
                ]
                style_cell = ws_styles[
                    f"{source_col}{row_idx}"
                ]

                original_value = value_cell.value
                value = original_value

                if is_empty(value):
                    value = config["empty_value"]
                    empty_count += 1

                    # Le motif vide a déjà son propre avertissement.
                    is_empty_motif = (
                        file_key == "file_02"
                        and field_name == motif_field
                    )

                    if not is_empty_motif:
                        entry = {
                            "file_key": file_key,
                            "source_file": input_path.name,
                            "source_sheet": sheet_name,
                            "source_coordinate": value_cell.coordinate,
                            "source_row": row_idx,
                            "source_column": source_col,
                            "field_name": field_name,
                            "original_value": original_value,
                            "replacement_value": value,
                            "reason": (
                                "cellule vide remplacée "
                                "par empty_value"
                            ),
                        }

                        change_log.append(entry)

                        print(
                            f"[REMPLACEMENT][LECTURE] "
                            f"{input_path.name} | "
                            f"{sheet_name}!{value_cell.coordinate} "
                            f"(ligne={row_idx}, "
                            f"colonne={source_col}) : "
                            f"cellule vide -> {value!r}",
                            flush=True,
                        )

                values_by_field[field_name] = value
                styles_by_field[field_name] = style_cell

            records.append({
                "file_key": file_key,
                "source_file": input_path.name,
                "source_sheet": sheet_name,
                "source_row": row_idx,
                "values_by_field": values_by_field,
                "styles_by_field": styles_by_field,
            })

        read_seconds = time.perf_counter() - read_start

        print(
            f"[TEMPS DE LECTURE] {input_path.name} : "
            f"{read_seconds:.3f} seconde(s)",
            flush=True,
        )

        print(
            f"[BILAN LECTURE] {input_path.name} | "
            f"lignes retenues={len(records)} | "
            f"motifs exclus={excluded_motif_count} | "
            f"dates exclues={excluded_date_count} | "
            f"cellules vides remplacées={empty_count}",
            flush=True,
        )

        return {
            "file_key": file_key,
            "input_path": input_path,
            "sheet_name": sheet_name,
            "fields": fields,
            "records": records,
            "change_log": change_log,
            "warnings": warnings,
            "empty_count": empty_count,
            "excluded_motif_count": excluded_motif_count,
            "excluded_date_count": excluded_date_count,

            "read_seconds": read_seconds,
            "workbook_values": wb_values,
            "workbook_styles": wb_styles,
            "exclusion_details": exclusion_details,

        }

    except Exception:
        if wb_values is not None:
            wb_values.close()

        if (
            wb_styles is not None
            and wb_styles is not wb_values
        ):
            wb_styles.close()

        raise


def read_configured_files(
    config: dict[str, Any],
) -> dict[str, Any]:
    """Valide la configuration puis lit les deux fichiers."""

    validate_config(config)

    output_path = Path(
        config["output_file"]
    ).expanduser().resolve()

    file_01_path = Path(
        config["file_01"]["input_file"]
    ).expanduser().resolve()

    file_02_path = Path(
        config["file_02"]["input_file"]
    ).expanduser().resolve()

    if output_path in {file_01_path, file_02_path}:
        raise ConfigurationError(
            "Le fichier de sortie doit être différent "
            "des deux fichiers source."
        )

    data_01 = read_one_file(
        "file_01",
        config["file_01"],
        config,
    )

    try:
        data_02 = read_one_file(
            "file_02",
            config["file_02"],
            config,
        )

    except Exception:
        data_01["workbook_values"].close()

        if (
            data_01["workbook_styles"]
            is not data_01["workbook_values"]
        ):
            data_01["workbook_styles"].close()

        raise

    output_fields = list(data_01["fields"].keys())

    for field_name in data_02["fields"]:
        if field_name not in output_fields:
            output_fields.append(field_name)

    return {
        "output_path": output_path,
        "output_fields": output_fields,
        "files": [data_01, data_02],
        "records": (
            data_01["records"] + data_02["records"]
        ),
        "change_log": (
            data_01["change_log"] + data_02["change_log"]
        ),
        "warnings": (
            data_01["warnings"] + data_02["warnings"]
        ),
        "empty_count": (
            data_01["empty_count"] + data_02["empty_count"]
        ),
        "excluded_motif_count": (
            data_01["excluded_motif_count"]
            + data_02["excluded_motif_count"]
        ),
        "excluded_date_count": (
            data_01["excluded_date_count"]
            + data_02["excluded_date_count"]
        ),
        "exclusion_details": (
            data_01["exclusion_details"]
            + data_02["exclusion_details"]
        ),
    }
