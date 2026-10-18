
"""Calcul des statistiques par âge et tranche d'âge."""

from collections import Counter
from typing import Any


def calculate_age_statistics(data: dict[str, Any]) -> dict[str, Any]:
    """
    Calcule les statistiques d'âge :
    - effectifs par tranche et par sexe ;
    - répartition des âges exacts ;
    - moyenne des âges par sexe ;
    - moyenne totale des âges valides.
    """

    employees = data["employees"]
    config = data["config"]

    labels = ["Moins de 16 ans"]
    labels.extend(band["label"] for band in config["age_bands"])
    labels.append("Âge inconnu")

    genders = ["Femme", "Homme", "VIDE"]

    bands = {
        label: {gender: 0 for gender in genders}
        for label in labels
    }

    exact_ages = Counter()

    # Les âges valides sont séparés par sexe.
    valid_ages_by_gender = {
        "Femme": [],
        "Homme": [],
        "VIDE": [],
    }

    for employee in employees:
        band = employee.get("age_band", "Âge inconnu")
        gender = employee.get("gender", "VIDE")

        if band not in bands:
            band = "Âge inconnu"

        if gender not in bands[band]:
            gender = "VIDE"

        bands[band][gender] += 1

        age = employee.get("age")

        # Un âge est exploitable uniquement s'il est numérique,
        # entier, non négatif et non nul.
        if (
            isinstance(age, int)
            and not isinstance(age, bool)
            and age >= 0
        ):
            exact_ages[age] += 1
            valid_ages_by_gender[gender].append(age)

    # Vérification de la cohérence des effectifs.
    band_total = sum(
        sum(counts.values())
        for counts in bands.values()
    )

    if band_total != len(employees):
        raise ValueError(
            f"Incohérence des statistiques d'âge : "
            f"{band_total} répartis pour "
            f"{len(employees)} employés retenus."
        )

    def calculate_mean(ages: list[int]) -> float | None:
        """Calcule la moyenne ou renvoie None si aucun âge valide."""
        if not ages:
            return None
        return sum(ages) / len(ages)

    female_ages = valid_ages_by_gender["Femme"]
    male_ages = valid_ages_by_gender["Homme"]

    # La moyenne totale exclut les civilités inconnues (VIDE),
    # afin de calculer la moyenne générale sur les sexes identifiés.
    # Si tu veux inclure les âges de sexe inconnu dans le total,
    # remplace total_ages par la concaténation des trois listes.
    total_ages = female_ages + male_ages

    average_ages = {
        "Femme": calculate_mean(female_ages),
        "Homme": calculate_mean(male_ages),
        "Total": calculate_mean(total_ages),
    }

    average_age_counts = {
        "Femme": len(female_ages),
        "Homme": len(male_ages),
        "Total": len(total_ages),
    }

    return {
        "reference_date": config["reference_date"],
        "bands": bands,
        "exact_ages": dict(sorted(exact_ages.items())),
        "average_ages": average_ages,
        "average_age_counts": average_age_counts,
        "invalid_birth_dates": data["invalid_birth_dates"],
        "future_birth_dates": data["future_birth_dates"],
        "code_depart_filtered_count": data[
            "code_depart_filtered_count"
        ],
        "empty_rows_skipped": data["empty_rows_skipped"],
    }
