
"""Calcul des statistiques d'effectifs : contrats, départs et grades."""

import unicodedata
from collections import Counter
from typing import Any


CONTRACT_TYPES = ("CDD", "CDI")
GRADES = ("Encadrement", "Cadre", "Maitrise", "Execution")
GENDERS = ("Femme", "Homme")


def normalize_text(value: Any) -> str:
    """
    Normalise un texte pour les comparaisons :
    - ignore la casse ;
    - supprime les accents ;
    - supprime les espaces inutiles ;
    - réduit les espaces successifs.
    """
    if value is None:
        return ""

    text = str(value).strip()
    text = " ".join(text.split())

    normalized = unicodedata.normalize("NFKD", text)
    normalized = "".join(
        char for char in normalized
        if not unicodedata.combining(char)
    )

    return normalized.upper()


def get_value(employee: dict[str, Any], *keys: str) -> Any:
    """Récupère une valeur en acceptant plusieurs noms de champs."""

    # Correspondance exacte en premier.
    for key in keys:
        if key in employee:
            return employee[key]

    # Correspondance insensible à la casse ensuite.
    normalized_keys = {
        normalize_text(key): value
        for key, value in employee.items()
    }

    for key in keys:
        normalized_key = normalize_text(key)
        if normalized_key in normalized_keys:
            return normalized_keys[normalized_key]

    return None


def classify_gender(civilite: Any) -> str | None:
    """
    Classe les employés selon leur civilité.

    Mme / Mlle : Femme
    M / Monsieur : Homme
    Autre valeur : sexe non déterminé
    """
    value = normalize_text(civilite)

    if value in {"MME", "MLLE", "MADEMOISELLE"}:
        return "Femme"

    if value in {"M", "MONSIEUR", "MR"}:
        return "Homme"

    return None


def calculate_contract_statistics(
    employees: list[dict[str, Any]],
) -> dict[str, Any]:
    """Compte les contrats CDD, CDI et CCTA."""

    counts = Counter({contract: 0 for contract in CONTRACT_TYPES})
    unknown_count = 0
    empty_count = 0

    for employee in employees:
        value = get_value(
            employee,
            "NATURE CONTRAT",
            "NATURE_CONTRAT",
            "CONTRAT",
        )
        normalized = normalize_text(value)

        if not normalized or normalized == "VIDE":
            empty_count += 1
        elif normalized == "CCTA":
            counts["CDI"] += 1

        elif normalized in counts:
            counts[normalized] += 1
        else:
            unknown_count += 1

    return {
        "counts": dict(counts),
        "total_known": sum(counts.values()),
        "unknown_count": unknown_count,
        "empty_count": empty_count,
    }



def calculate_departure_reason_statistics(
    employees: list[dict[str, Any]],
) -> dict[str, Any]:
    """
    Compte les motifs de départ et signale les motifs réellement vides.

    - Ignore la casse, les accents et les espaces superflus.
    - Ignore les valeurs VIDE dans les statistiques.
    - Conserve les lignes sans motif pour un contrôle dans Excel.
    """

    counts: Counter[str] = Counter()
    display_names: dict[str, str] = {}
    ignored_count = 0
    unread_rows: list[dict[str, Any]] = []

    for employee in employees:
        value = get_value(
            employee,
            "Motif DEPART",
            "MOTIF DEPART",
            "MOTIF_DEPART",
            "MOTIF DE DÉPART",
            "MOTIF",
        )

        display_value = (
            " ".join(str(value).split())
            if value is not None
            else ""
        )

        normalized = normalize_text(display_value)

        # Les cellules réellement vides sont conservées
        # dans un rapport de contrôle.
        if not normalized:
            ignored_count += 1

            unread_rows.append({
                "ligne_source": employee.get("source_row", "Inconnue"),
                "CCP": get_value(employee, "CCP") or "",
                "motif_brut": "",
                "anomalie": "Motif vide ou non récupéré",
            })
            continue

        # VIDE est exclu des statistiques, mais ne signifie
        # pas automatiquement que la lecture a échoué.
        if normalized == "VIDE":
            ignored_count += 1
            continue

        # Toute valeur non vide est comptabilisée.
        # Les variantes de casse et d'accents sont regroupées.
        counts[normalized] += 1

        if normalized not in display_names:
            display_names[normalized] = display_value

    reasons = [
        {
            "motif": display_names[key],
            "count": count,
        }
        for key, count in counts.most_common()
    ]

    return {
        "reasons": reasons,
        "total_counted": sum(counts.values()),
        "ignored_count": ignored_count,
        "unread_rows": unread_rows,
    }


def calculate_grade_statistics(
    employees: list[dict[str, Any]],
) -> dict[str, Any]:
    """
    Compte les grades Encadrement, Cadre, Maitrise et Execution
    séparément pour les femmes et les hommes.
    """
    counts = {
        gender: {grade: 0 for grade in GRADES}
        for gender in GENDERS
    }

    unknown_gender_count = 0
    unknown_grade_count = 0
    empty_grade_count = 0

    for employee in employees:
        civilite = get_value(
            employee,
            "CIVILITE",
            "CIVILITÉ",
        )
        gender = classify_gender(civilite)

        grade_value = get_value(
            employee,
            "GRADE",
            "Grade",
            "grade",
        )
        normalized_grade = normalize_text(grade_value)

        if not normalized_grade or normalized_grade == "VIDE":
            empty_grade_count += 1
            continue

        grade_mapping = {
            normalize_text(grade): grade
            for grade in GRADES
        }
        grade = grade_mapping.get(normalized_grade)

        if grade is None:
            unknown_grade_count += 1
            continue

        if gender is None:
            unknown_gender_count += 1
            continue

        counts[gender][grade] += 1

    return {
        "counts": counts,
        "unknown_gender_count": unknown_gender_count,
        "unknown_grade_count": unknown_grade_count,
        "empty_grade_count": empty_grade_count,
    }


def calculate_effectifs_statistics(
    employees_data: dict[str, Any],
) -> dict[str, Any]:
    """Calcule toutes les statistiques d'effectifs."""

    employees = employees_data.get("employees")

    if not isinstance(employees, list):
        raise ValueError(
            "La structure des données est incorrecte : "
            "la clé 'employees' doit contenir une liste."
        )

    if not employees:
        raise ValueError(
            "Aucun employé disponible pour calculer les effectifs."
        )

    # Vérification explicite : le grade doit être présent dans les données.
    grade_exists = any(
        get_value(employee, "GRADE", "Grade", "grade") is not None
        for employee in employees
    )

    if not grade_exists:
        raise ValueError(
            "La colonne GRADE est absente des données chargées. "
            "Ajoute 'GRADE' dans la section 'columns' de "
            "stat_config.json et indique la lettre Excel correspondante."
        )

    return {
        "total_employees": len(employees),
        "contracts": calculate_contract_statistics(employees),
        "departure_reasons": calculate_departure_reason_statistics(
            employees
        ),
        "grades": calculate_grade_statistics(employees),
    }
