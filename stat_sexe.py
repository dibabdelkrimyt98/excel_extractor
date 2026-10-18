"""Calcul des statistiques par civilité."""

from collections import Counter


def calculate_gender_statistics(data):
    counts = Counter(employee.get("gender", "VIDE") for employee in data["employees"])
    return {
        "total": len(data["employees"]),
        "counts": {
            "Femme": counts.get("Femme", 0),
            "Homme": counts.get("Homme", 0),
            "VIDE": counts.get("VIDE", 0),
        },
    }
