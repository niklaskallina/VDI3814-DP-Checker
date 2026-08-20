"""Kostenschaetzung ueber Einheitspreise je Funktionsspalte."""

from __future__ import annotations

import pandas as pd

from .config import SETTINGS
from .profiles_loader import get_profile


def price_template(profile_id: str, prices: dict[str, float] | None = None) -> pd.DataFrame:
    """Preisliste fuer ein Profil - auch ohne bereits importierte Daten nutzbar."""
    profile = get_profile(profile_id)
    if profile is None:
        return pd.DataFrame(columns=["spalte_key", "spalte_adresse", "gruppe", "untergruppe",
                                     "spalte", "einheitspreis"])
    prices = prices or {}
    return pd.DataFrame([
        {
            "spalte_key": column.key,
            "spalte_adresse": column.address,
            "gruppe": column.group,
            "untergruppe": column.subgroup,
            "spalte": column.label,
            "einheitspreis": float(prices.get(column.key, 0.0)),
        }
        for column in profile.columns
    ])


def apply_prices(summary: pd.DataFrame, prices: dict[str, float],
                 currency: str | None = None) -> pd.DataFrame:
    """Ergaenzt eine Spaltenuebersicht um Einheitspreis und Kosten."""
    currency = currency or SETTINGS.currency
    result = summary.copy()
    if result.empty:
        result["einheitspreis"] = []
        result["kosten"] = []
        return result
    result["einheitspreis"] = [float(prices.get(key, 0.0)) for key in result["spalte_key"]]
    result["kosten"] = result["menge"].astype(float) * result["einheitspreis"]
    result["waehrung"] = currency
    return result


def total_cost(priced: pd.DataFrame) -> float:
    if priced.empty or "kosten" not in priced:
        return 0.0
    return float(priced["kosten"].sum())
