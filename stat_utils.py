"""Fonctions communes aux modules de statistiques Excel."""

import json
import math
import shutil
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import from_excel
from copy import copy


def load_config(config_path="stat_config.json"):
    path = Path(config_path)
    if not path.is_file():
        raise FileNotFoundError(f"Configuration introuvable : {path.resolve()}")

    try:
        with path.open("r", encoding="utf-8") as file:
            config = json.load(file)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"JSON invalide à la ligne {exc.lineno}, colonne {exc.colno}."
        ) from exc

    required = (
        "input_file", "output_file", "sheet_name", "start_row",
        "reference_date", "columns", "age_bands",
    )
    missing = [key for key in required if key not in config]
    if missing:
        raise ValueError(f"Paramètres manquants : {', '.join(missing)}")

    if not isinstance(config["columns"], dict):
        raise ValueError("'columns' doit être un objet JSON.")
    if not isinstance(config["age_bands"], list):
        raise ValueError("'age_bands' doit être une liste.")
    if not isinstance(config["start_row"], int) or config["start_row"] < 1:
        raise ValueError("'start_row' doit être un entier positif.")

    for name, letter in config["columns"].items():
        if not isinstance(letter, str):
            raise ValueError(f"Colonne invalide pour {name!r} : {letter!r}")
        try:
            column_index_from_string(letter.strip().upper())
        except ValueError as exc:
            raise ValueError(
                f"Colonne Excel invalide pour {name!r} : {letter!r}"
            ) from exc

    code_column = config.get("code_depart_column")
    if code_column is not None:
        if not isinstance(code_column, str):
            raise ValueError("'code_depart_column' doit être une lettre Excel ou null.")
        try:
            column_index_from_string(code_column.strip().upper())
        except ValueError as exc:
            raise ValueError(
                f"Colonne Excel invalide pour 'code_depart_column' : {code_column!r}"
            ) from exc

    try:
        config["_reference_date"] = date.fromisoformat(config["reference_date"])
    except (TypeError, ValueError) as exc:
        raise ValueError("'reference_date' doit respecter le format AAAA-MM-JJ.") from exc

    date_formats = config.get("date_formats", ["%m/%d/%Y"])
    if not isinstance(date_formats, list) or date_formats != ["%m/%d/%Y"]:
        raise ValueError(
            "'date_formats' doit être exactement ['%m/%d/%Y'] "
            "car les dates textuelles source sont au format MM/DD/YYYY."
        )
    config["date_formats"] = date_formats

    validated_bands = []
    labels = set()
    for band in config["age_bands"]:
        if not isinstance(band, dict) or not {"label", "min", "max"} <= set(band):
            raise ValueError("Chaque tranche doit contenir label, min et max.")
        label, minimum, maximum = band["label"], band["min"], band["max"]
        if not isinstance(label, str) or not label.strip():
            raise ValueError("Chaque tranche doit avoir un label non vide.")
        if label in labels:
            raise ValueError(f"Label de tranche en double : {label!r}")
        if not isinstance(minimum, int) or minimum < 0:
            raise ValueError(f"Borne minimale invalide pour {label!r}.")
        if maximum is not None and (not isinstance(maximum, int) or maximum < minimum):
            raise ValueError(f"Borne maximale invalide pour {label!r}.")
        labels.add(label)
        validated_bands.append((minimum, maximum, label))

    validated_bands.sort(key=lambda item: item[0])
    previous_max = None
    for minimum, maximum, label in validated_bands:
        if previous_max is not None and minimum <= previous_max:
            raise ValueError(f"Les tranches d'âge se chevauchent près de {label!r}.")
        previous_max = float("inf") if maximum is None else maximum

    # Colonne de sortie de l'âge dans resultat.xlsx (I par défaut, après A:H).
    age_column = config.get("age_output_column", "I")
    if not isinstance(age_column, str):
        raise ValueError("'age_output_column' doit être une lettre Excel.")
    try:
        column_index_from_string(age_column.strip().upper())
    except ValueError as exc:
        raise ValueError(f"Colonne invalide pour 'age_output_column' : {age_column!r}") from exc
    config["age_output_column"] = age_column.strip().upper()
    config["age_header"] = config.get("age_header", "AGE")

    return config


def is_empty(value):
    return value is None or (
        isinstance(value, str) and value.strip().upper() in {"", "VIDE"}
    )


def parse_date(value, formats=None, excel_epoch=None):
    """Convertit une date Excel ou une chaîne MM/DD/YYYY en datetime.date."""
    if is_empty(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if not math.isfinite(value) or not 1000 <= value <= 100000:
            return None
        try:
            converted = from_excel(value, excel_epoch) if excel_epoch else from_excel(value)
            if isinstance(converted, datetime):
                return converted.date()
            if isinstance(converted, date):
                return converted
        except (OverflowError, TypeError, ValueError):
            return None
        return None
    if isinstance(value, str):
        text = value.strip()
        for fmt in (formats if formats is not None else ["%m/%d/%Y"]):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def calculate_age(birth_date, reference_date):
    """Âge révolu en années complètes, équivalent à DATEDIF(date_naissance; référence; "Y")."""
    if birth_date is None or birth_date > reference_date:
        return None
    years = reference_date.year - birth_date.year
    if (reference_date.month, reference_date.day) < (birth_date.month, birth_date.day):
        years -= 1
    return years


def add_age_column_to_input(config):
    """Ajoute/met à jour la colonne AGE dans resultat.xlsx sans exposer les données.

    Crée une copie de sauvegarde .bak lors de la première exécution.
    Les dates textuelles sont acceptées uniquement au format MM/DD/YYYY.
    """
    input_path = Path(config["input_file"]).expanduser().resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"Fichier source introuvable : {input_path}")
    if input_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Le fichier source doit être au format .xlsx ou .xlsm.")

    backup_path = input_path.with_suffix(input_path.suffix + ".bak")
    if not backup_path.exists():
        shutil.copy2(input_path, backup_path)

    workbook = load_workbook(input_path, data_only=False)
    try:
        sheet_name = config["sheet_name"]
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Feuille {sheet_name!r} introuvable dans {input_path.name}.")
        worksheet = workbook[sheet_name]

        birth_letter = config["columns"]["DATE NAISSANCE"].strip().upper()
        age_letter = config["age_output_column"]
        birth_index = column_index_from_string(birth_letter)
        age_index = column_index_from_string(age_letter)
        header_value = worksheet.cell(row=config.get("header_row", 1), column=age_index).value
        if not is_empty(header_value) and str(header_value).strip().casefold() != str(config["age_header"]).strip().casefold():
            raise ValueError(
                f"La colonne {age_letter} contient déjà l'en-tête {header_value!r}. "
                f"Vérifie 'age_output_column' dans stat_config.json avant de continuer."
            )

        header_row = config.get("header_row", 1)
        header_cell = worksheet.cell(row=header_row, column=age_index)
        source_header = worksheet.cell(row=header_row, column=max(1, age_index - 1))
        if source_header.has_style:
            header_cell._style = copy(source_header._style)
        header_cell.value = config["age_header"]
        header_cell.alignment = copy(source_header.alignment) if source_header.alignment else Alignment(horizontal="center")

        # Dernière ligne réellement renseignée dans les colonnes métier configurées.
        source_indexes = [column_index_from_string(letter.strip().upper()) for letter in config["columns"].values()]
        last_row = header_row
        for row_num in range(config["start_row"], worksheet.max_row + 1):
            if any(not is_empty(worksheet.cell(row=row_num, column=idx).value) for idx in source_indexes):
                last_row = row_num

        reference_date = config["_reference_date"]
        invalid_dates = 0
        for row_num in range(config["start_row"], last_row + 1):
            birth_cell = worksheet.cell(row=row_num, column=birth_index)
            birth_date = parse_date(birth_cell.value, config["date_formats"], workbook.epoch)
            age = calculate_age(birth_date, reference_date)
            age_cell = worksheet.cell(row=row_num, column=age_index)
            style_source = worksheet.cell(row=row_num, column=max(1, age_index - 1))
            if style_source.has_style:
                age_cell._style = copy(style_source._style)
            if age is None:
                age_cell.value = config.get("empty_value", "VIDE")
                if not is_empty(birth_cell.value):
                    invalid_dates += 1
            else:
                age_cell.value = age
                age_cell.number_format = "0"

        worksheet.column_dimensions[age_letter].width = 12
        workbook.save(input_path)
        return {"path": input_path, "backup_path": backup_path, "rows": max(0, last_row - config["start_row"] + 1), "invalid_dates": invalid_dates, "column": age_letter}
    finally:
        workbook.close()


def normalize_code(value):
    """Normalise un code numérique, y compris ' 18 ' et '18.0'."""
    if is_empty(value):
        return None
    text = str(value).strip()
    try:
        number = Decimal(text)
        if number.is_finite() and number == number.to_integral_value():
            return str(int(number))
    except (InvalidOperation, ValueError):
        pass
    return text.casefold()


def classify_gender(civilite):
    """Mme/Mlle = Femme; M/Mr/Monsieur = Homme; autres valeurs = VIDE."""
    if is_empty(civilite):
        return "VIDE"
    value = str(civilite).strip().casefold().replace(".", "")
    if value in {"mme", "mlle"}:
        return "Femme"
    if value in {"m", "mr", "monsieur"}:
        return "Homme"
    return "VIDE"


def classify_age(age, age_bands):
    if age is None:
        return "Âge inconnu"
    if age < 16:
        return "Moins de 16 ans"
    for band in age_bands:
        minimum = band["min"]
        maximum = band["max"]
        if age >= minimum and (maximum is None or age <= maximum):
            return band["label"]
    return "Âge inconnu"


def load_employees(config):
    """Lit les employés, exclut le code départ 18 si une colonne est configurée."""
    input_path = Path(config["input_file"]).expanduser().resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"Fichier source introuvable : {input_path}")
    if input_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Le fichier source doit être au format .xlsx ou .xlsm.")
    output_path = Path(config["output_file"]).expanduser().resolve()
    if input_path == output_path:
        raise ValueError("Le fichier source et le fichier de sortie doivent être différents.")

    columns = config["columns"]
    required_columns = {"CIVILITE", "DATE NAISSANCE"}
    missing_columns = required_columns - set(columns)
    if missing_columns:
        raise ValueError("Colonnes nécessaires manquantes : " + ", ".join(sorted(missing_columns)))
    column_indexes = {name: column_index_from_string(letter.strip().upper()) for name, letter in columns.items()}
    code_depart_index = column_index_from_string(config["code_depart_column"].strip().upper()) if config.get("code_depart_column") else None
    max_column = max([*column_indexes.values(), code_depart_index or 0])

    workbook = load_workbook(input_path, read_only=True, data_only=True)
    try:
        sheet_name = config["sheet_name"]
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Feuille {sheet_name!r} introuvable. Feuilles disponibles : {', '.join(workbook.sheetnames)}")
        worksheet = workbook[sheet_name]
        reference_date = config["_reference_date"]
        age_bands = config["age_bands"]
        employees = []
        invalid_birth_dates = future_birth_dates = code_depart_filtered_count = empty_rows_skipped = 0
        for row_number, row in enumerate(worksheet.iter_rows(min_row=config["start_row"], max_col=max_column, values_only=True), start=config["start_row"]):
            values = {name: row[index - 1] if index - 1 < len(row) else None for name, index in column_indexes.items()}
            code_value = row[code_depart_index - 1] if code_depart_index is not None and code_depart_index - 1 < len(row) else None
            if all(is_empty(value) for value in values.values()) and is_empty(code_value):
                empty_rows_skipped += 1
                continue
            if code_depart_index is not None and normalize_code(code_value) == "18":
                code_depart_filtered_count += 1
                continue
            birth_date = parse_date(values.get("DATE NAISSANCE"), formats=config["date_formats"], excel_epoch=workbook.epoch)
            gender = classify_gender(values.get("CIVILITE"))
            if birth_date is None:
                invalid_birth_dates += 1
                age = None
            elif birth_date > reference_date:
                future_birth_dates += 1
                age = None
            else:
                age = calculate_age(birth_date, reference_date)
            employees.append({"source_row": row_number, "values": values, "gender": gender, "birth_date": birth_date, "age": age, "age_band": classify_age(age, age_bands)})
        return {"employees": employees, "invalid_birth_dates": invalid_birth_dates, "future_birth_dates": future_birth_dates, "code_depart_filtered_count": code_depart_filtered_count, "empty_rows_skipped": empty_rows_skipped, "config": config}
    finally:
        workbook.close()
