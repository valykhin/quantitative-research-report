"""Этап 4. Однофакторный (дескриптивный) анализ."""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
import seaborn as sns
from scipy import stats

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "figure.dpi": 150,
    "savefig.bbox": "tight",
    "savefig.pad_inches": 0.2,
})

OUT = Path(__file__).parent / "output"
PLOTS = OUT / "plots"
PLOTS.mkdir(parents=True, exist_ok=True)

KEY_COLS = [
    "Выручка", "Себестоимость продаж", "Чистая прибыль (убыток)",
    "Активы всего", "Денежные средства", "Кредиторская задолженность",
    "Валовая рентабельность, %", "Рентабельность по ЧП, %",
    "Доля себестоимости, %", "Оборачиваемость активов",
    "Темп роста Выручка, %",
]


def load() -> pd.DataFrame:
    return pd.read_excel(Path(__file__).parent / "data" / "analysis_dataset.xlsx")


# ── 4.1 Таблицы частот ─────────────────────────────────────────────

def frequency_tables(df: pd.DataFrame) -> pd.DataFrame:
    print("=" * 60)
    print("4.1.1  Таблица частот: ОПФ")
    print(df["ОПФ"].value_counts().to_string())

    print("\n4.1.2  Таблица частот: Компании × Периоды")
    ct = df.pivot_table(index="Компания", columns="Период", values="Активы всего", aggfunc="count")
    print(ct.fillna(0).astype(int).to_string())

    rev = df[df["Выручка"].notna()]
    print(f"\n4.1.3  Наблюдений с выручкой: {len(rev)} из {len(df)}")
    print(f"       Компаний с выручкой: {rev['Компания'].nunique()} из {df['Компания'].nunique()}")
    return rev


# ── 4.2 Описательные статистики ────────────────────────────────────

def descriptive_stats(df: pd.DataFrame) -> pd.DataFrame:
    print("\n" + "=" * 60)
    print("4.2  Описательные статистики (выборка с выручкой)")

    rows = []
    for col in KEY_COLS:
        vals = df[col].dropna()
        if len(vals) < 3:
            continue

        skew = stats.skew(vals, bias=False)
        kurt = stats.kurtosis(vals, bias=False)
        q1, q3 = np.percentile(vals, [25, 75])
        iqr = q3 - q1

        typical = "медиана" if abs(skew) > 1 else "среднее"

        row = {
            "Переменная": col,
            "N": len(vals),
            "Среднее": vals.mean(),
            "Медиана": vals.median(),
            "Станд. откл.": vals.std(),
            "Min": vals.min(),
            "Max": vals.max(),
            "Размах": vals.max() - vals.min(),
            "Q1": q1,
            "Q3": q3,
            "IQR": iqr,
            "Асимметрия": skew,
            "Эксцесс": kurt,
            "Типичное среднее": typical,
        }
        rows.append(row)
        print(f"\n  {col}:")
        print(f"    N={len(vals)}, среднее={vals.mean():,.1f}, медиана={vals.median():,.1f}")
        print(f"    σ={vals.std():,.1f}, min={vals.min():,.1f}, max={vals.max():,.1f}")
        print(f"    асимметрия={skew:.2f}, эксцесс={kurt:.2f} → типичное: {typical}")

    result = pd.DataFrame(rows)
    path = OUT / "descriptive_stats.xlsx"
    result.to_excel(path, index=False, engine="xlsxwriter")
    print(f"\n  Сохранено: {path}")
    return result


# ── 4.3 Проверка нормальности ──────────────────────────────────────

def normality_tests(df: pd.DataFrame) -> pd.DataFrame:
    print("\n" + "=" * 60)
    print("4.3  Проверка нормальности (Шапиро-Уилк)")

    rows = []
    for col in KEY_COLS:
        vals = df[col].dropna()
        if len(vals) < 3:
            continue
        stat, p = stats.shapiro(vals)
        normal = "Да" if p > 0.05 else "Нет"
        rows.append({"Переменная": col, "N": len(vals), "W": stat, "p-value": p, "Нормальное (α=0.05)": normal})
        print(f"  {col:40s}  W={stat:.4f}  p={p:.4f}  {'✓ нормальное' if p > 0.05 else '✗ ненормальное'}")

    result = pd.DataFrame(rows)
    path = OUT / "normality_tests.xlsx"
    result.to_excel(path, index=False, engine="xlsxwriter")
    print(f"\n  Сохранено: {path}")
    return result


# ── 4.4 Графики ────────────────────────────────────────────────────

def plot_histograms(df: pd.DataFrame):
    """Гистограммы распределения ключевых переменных."""
    plot_cols = [c for c in KEY_COLS if df[c].notna().sum() >= 5]
    n = len(plot_cols)
    cols_grid = 3
    rows_grid = (n + cols_grid - 1) // cols_grid

    fig, axes = plt.subplots(rows_grid, cols_grid, figsize=(14, 4 * rows_grid))
    axes = axes.flatten()

    for i, col in enumerate(plot_cols):
        vals = df[col].dropna()
        axes[i].hist(vals, bins=min(15, max(5, len(vals) // 3)), edgecolor="white", alpha=0.8, color="#4C72B0")
        axes[i].axvline(vals.mean(), color="red", linestyle="--", linewidth=1, label=f"среднее={vals.mean():.1f}")
        axes[i].axvline(vals.median(), color="green", linestyle="-", linewidth=1, label=f"медиана={vals.median():.1f}")
        axes[i].set_title(col, fontsize=10)
        axes[i].legend(fontsize=7)

    for j in range(i + 1, len(axes)):
        axes[j].set_visible(False)

    fig.suptitle("Гистограммы распределения ключевых переменных", fontsize=13, y=1.01)
    fig.tight_layout()
    fig.savefig(PLOTS / "histograms.png")
    plt.close(fig)
    print(f"  Гистограммы: {PLOTS / 'histograms.png'}")


def plot_boxplots(df: pd.DataFrame):
    """Boxplot для сравнения компаний."""
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))

    for ax, col in zip(axes.flatten(), ["Выручка", "Чистая прибыль (убыток)", "Валовая рентабельность, %", "Оборачиваемость активов"]):
        data = df[["Компания", col]].dropna()
        if data.empty:
            continue
        order = data.groupby("Компания")[col].median().sort_values(ascending=False).index
        sns.boxplot(data=data, y="Компания", x=col, order=order, ax=ax, palette="Set2")
        ax.set_title(col, fontsize=11)
        ax.set_ylabel("")

    fig.suptitle("Распределение показателей по компаниям", fontsize=13, y=1.01)
    fig.tight_layout()
    fig.savefig(PLOTS / "boxplots.png")
    plt.close(fig)
    print(f"  Boxplot: {PLOTS / 'boxplots.png'}")


def plot_dynamics(df: pd.DataFrame):
    """Динамика ключевых показателей по годам."""
    rev_companies = df[df["Выручка"].notna()].groupby("Компания").filter(lambda x: len(x) >= 3)
    companies = sorted(rev_companies["Компания"].unique())

    fig, axes = plt.subplots(2, 2, figsize=(16, 11))
    metrics = [
        ("Выручка", "Выручка, тыс. руб."),
        ("Чистая прибыль (убыток)", "Чистая прибыль, тыс. руб."),
        ("Валовая рентабельность, %", "Валовая рентабельность, %"),
        ("Активы всего", "Активы, тыс. руб."),
    ]

    for ax, (col, label) in zip(axes.flatten(), metrics):
        for comp in companies:
            grp = rev_companies[rev_companies["Компания"] == comp].sort_values("Период")
            vals = grp[["Период", col]].dropna()
            if len(vals) >= 2:
                short = comp.replace("ООО ", "").replace("АНО ", "").replace("АССОЦИАЦИЯ ", "").strip('"')
                ax.plot(vals["Период"], vals[col], marker="o", markersize=4, label=short, linewidth=1.5)
        ax.set_title(label, fontsize=11)
        ax.legend(fontsize=7, loc="upper left")
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))

    fig.suptitle("Динамика ключевых показателей (2021–2025)", fontsize=13, y=1.01)
    fig.tight_layout()
    fig.savefig(PLOTS / "dynamics.png")
    plt.close(fig)
    print(f"  Динамика: {PLOTS / 'dynamics.png'}")


def plot_revenue_growth(df: pd.DataFrame):
    """Темпы роста выручки по годам."""
    data = df[["Компания", "Период", "Темп роста Выручка, %"]].dropna()
    companies = data.groupby("Компания").filter(lambda x: len(x) >= 2)["Компания"].unique()

    fig, ax = plt.subplots(figsize=(12, 6))
    for comp in sorted(companies):
        grp = data[data["Компания"] == comp].sort_values("Период")
        short = comp.replace("ООО ", "").replace("АНО ", "").replace("АССОЦИАЦИЯ ", "").strip('"')
        ax.plot(grp["Период"], grp["Темп роста Выручка, %"], marker="o", markersize=5, label=short, linewidth=1.5)

    ax.axhline(0, color="black", linewidth=0.5)
    ax.set_title("Темпы роста выручки, % г/г", fontsize=12)
    ax.set_ylabel("%")
    ax.legend(fontsize=8)
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    fig.tight_layout()
    fig.savefig(PLOTS / "revenue_growth.png")
    plt.close(fig)
    print(f"  Темпы роста: {PLOTS / 'revenue_growth.png'}")


def plot_market_summary(df: pd.DataFrame):
    """Сводные показатели рынка по годам (медиана)."""
    rev = df[df["Выручка"].notna()]
    agg = rev.groupby("Период").agg(
        Медиана_выручки=("Выручка", "median"),
        Сумма_выручки=("Выручка", "sum"),
        Медиана_рентабельности=("Валовая рентабельность, %", "median"),
        Кол_во_компаний=("Компания", "nunique"),
    ).reset_index()

    fig, axes = plt.subplots(1, 3, figsize=(15, 5))

    axes[0].bar(agg["Период"], agg["Сумма_выручки"] / 1000, color="#4C72B0", alpha=0.8)
    axes[0].set_title("Суммарная выручка рынка, млн руб.")
    axes[0].set_ylabel("млн руб.")

    axes[1].plot(agg["Период"], agg["Медиана_выручки"], marker="o", color="#DD8452", linewidth=2)
    axes[1].set_title("Медианная выручка, тыс. руб.")

    axes[2].plot(agg["Период"], agg["Медиана_рентабельности"], marker="s", color="#55A868", linewidth=2)
    axes[2].set_title("Медианная валовая рентабельность, %")
    axes[2].set_ylabel("%")

    for ax in axes:
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))

    fig.suptitle("Сводные показатели рынка беговых мероприятий", fontsize=13, y=1.02)
    fig.tight_layout()
    fig.savefig(PLOTS / "market_summary.png")
    plt.close(fig)
    print(f"  Сводка рынка: {PLOTS / 'market_summary.png'}")


def main():
    df = load()
    rev_df = frequency_tables(df)
    desc = descriptive_stats(rev_df)
    norm = normality_tests(rev_df)

    print("\n" + "=" * 60)
    print("4.4  Графики")
    plot_histograms(rev_df)
    plot_boxplots(rev_df)
    plot_dynamics(rev_df)
    plot_revenue_growth(rev_df)
    plot_market_summary(rev_df)

    print("\n" + "=" * 60)
    print("ИТОГ")
    print(f"  Описательные статистики: {OUT / 'descriptive_stats.xlsx'}")
    print(f"  Тесты нормальности: {OUT / 'normality_tests.xlsx'}")
    print(f"  Графики: {PLOTS}/")

    normal_count = (norm["Нормальное (α=0.05)"] == "Да").sum()
    total = len(norm)
    print(f"\n  Нормальность: {normal_count}/{total} переменных прошли тест Шапиро-Уилка")
    if normal_count < total:
        print("  → Рекомендуется использовать непараметрические методы (Спирмен, Вилкоксон)")


if __name__ == "__main__":
    main()
