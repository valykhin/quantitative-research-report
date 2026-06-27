"""Этап 3. Операционализация переменных и подготовка аналитического датасета."""

from pathlib import Path

import pandas as pd
import numpy as np


def load_raw() -> pd.DataFrame:
    return pd.read_excel(Path(__file__).parent / "data" / "companies_financials.xlsx")


def add_derived(df: pd.DataFrame) -> pd.DataFrame:
    """Рассчитать производные показатели."""
    d = df.copy()

    # --- Рентабельность ---
    d["Валовая рентабельность, %"] = (
        (d["Выручка"] - d["Себестоимость продаж"]) / d["Выручка"] * 100
    )
    d["Рентабельность по ЧП, %"] = (
        d["Чистая прибыль (убыток)"] / d["Выручка"] * 100
    )
    d["Доля себестоимости, %"] = d["Себестоимость продаж"] / d["Выручка"] * 100

    # --- Структура активов ---
    d["Доля ОС в активах, %"] = d["Основные средства"] / d["Активы всего"] * 100
    d["Доля ДС в активах, %"] = d["Денежные средства"] / d["Активы всего"] * 100
    d["Доля ДЗ в активах, %"] = d["Дебиторская задолженность"] / d["Активы всего"] * 100

    # --- Структура пассивов ---
    d["Доля КЗ в активах, %"] = d["Кредиторская задолженность"] / d["Активы всего"] * 100
    d["Коэффициент автономии"] = d["Капитал и резервы"] / d["Активы всего"]

    # --- Оборачиваемость ---
    d["Оборачиваемость активов"] = d["Выручка"] / d["Активы всего"]

    # --- Темпы роста (год к году) ---
    d = d.sort_values(["Компания", "Период"])
    for col in ["Выручка", "Активы всего", "Кредиторская задолженность"]:
        d[f"Темп роста {col}, %"] = (
            d.groupby("Компания")[col].pct_change() * 100
        )

    # --- ОПФ ---
    d["ОПФ"] = d["Компания"].apply(
        lambda x: "ООО" if x.startswith("ООО") else "НКО"
    )

    return d


def compute_mean_changes(df: pd.DataFrame) -> pd.DataFrame:
    """Средние абсолютные изменения показателей за период (по образцу примера)."""
    key_cols = [
        "Основные средства", "Запасы", "Дебиторская задолженность",
        "Денежные средства", "Кредиторская задолженность",
        "Выручка", "Себестоимость продаж", "Чистая прибыль (убыток)",
    ]

    rows = []
    for name, grp in df.sort_values("Период").groupby("Компания"):
        row = {"Компания": name}
        for col in key_cols:
            vals = grp[col].dropna()
            if len(vals) >= 2:
                diffs = vals.diff().dropna()
                row[f"Ср. изм. {col}"] = diffs.mean()
            else:
                row[f"Ср. изм. {col}"] = np.nan
        rows.append(row)

    return pd.DataFrame(rows)


def main():
    out_dir = Path(__file__).parent / "data"

    df = load_raw()
    df = add_derived(df)

    # --- Полный датасет ---
    full_path = out_dir / "analysis_dataset.xlsx"
    df.to_excel(full_path, index=False, engine="xlsxwriter")
    print(f"Полный датасет: {full_path} ({len(df)} строк, {len(df.columns)} столбцов)")

    # --- Выборка: только компании с данными по выручке ---
    df_rev = df[df["Выручка"].notna()].copy()
    rev_path = out_dir / "analysis_revenue_subset.xlsx"
    df_rev.to_excel(rev_path, index=False, engine="xlsxwriter")
    print(f"Выборка с выручкой: {rev_path} ({len(df_rev)} строк, {df_rev['Компания'].nunique()} компаний)")

    # --- Средние изменения ---
    changes = compute_mean_changes(df)
    changes_path = out_dir / "mean_changes.xlsx"
    changes.to_excel(changes_path, index=False, engine="xlsxwriter")
    print(f"Средние изменения: {changes_path}")

    # --- Сводка ---
    print(f"\n{'='*60}")
    print("Переменные в датасете:")
    for i, col in enumerate(df.columns, 1):
        dtype = df[col].dtype
        notna = df[col].notna().sum()
        print(f"  {i:>2}. {col:45s} {str(dtype):>10s}  {notna:>3}/{len(df)}")

    print(f"\n{'='*60}")
    print("Производные показатели (пример, последний период):")
    latest = df[df["Период"] == 2025]
    show_cols = [
        "Компания", "Выручка", "Валовая рентабельность, %",
        "Рентабельность по ЧП, %", "Оборачиваемость активов",
        "Доля КЗ в активах, %", "ОПФ",
    ]
    existing = [c for c in show_cols if c in latest.columns]
    print(latest[existing].to_string(index=False))


if __name__ == "__main__":
    main()
