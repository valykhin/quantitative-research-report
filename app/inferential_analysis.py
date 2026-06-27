"""Этап 5. Инференциальный (заключающий) анализ."""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from scipy import stats

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "figure.dpi": 150,
    "savefig.bbox": "tight",
})

OUT = Path(__file__).parent / "output"
PLOTS = OUT / "plots"
PLOTS.mkdir(parents=True, exist_ok=True)


def load() -> pd.DataFrame:
    return pd.read_excel(Path(__file__).parent / "data" / "analysis_dataset.xlsx")


# ── H1: Рынок демонстрирует устойчивый рост ────────────────────────

def test_h1(df: pd.DataFrame):
    print("=" * 60)
    print("H1: Медианный темп роста выручки положителен")
    print()

    growth = df["Темп роста Выручка, %"].dropna()
    print(f"  N = {len(growth)}")
    print(f"  Среднее = {growth.mean():.1f}%")
    print(f"  Медиана = {growth.median():.1f}%")
    print(f"  Положительных = {(growth > 0).sum()} из {len(growth)}")

    # Вилкоксон (ненормальное распределение — подтверждено на этапе 4)
    stat_w, p_w = stats.wilcoxon(growth, alternative="greater")
    print(f"\n  Тест Вилкоксона (H₀: медиана ≤ 0, H₁: медиана > 0):")
    print(f"    W = {stat_w:.1f}, p = {p_w:.6f}")
    print(f"    → {'ОТВЕРГАЕМ H₀' if p_w < 0.05 else 'НЕ ОТВЕРГАЕМ H₀'} (α = 0.05)")

    # Для полноты — t-тест
    stat_t, p_t = stats.ttest_1samp(growth, 0, alternative="greater")
    print(f"\n  t-тест (для сравнения):")
    print(f"    t = {stat_t:.3f}, p = {p_t:.6f}")

    # По годам
    print("\n  Медианный темп роста по годам:")
    by_year = df.groupby("Период")["Темп роста Выручка, %"].agg(["median", "count"])
    for year, row in by_year.iterrows():
        if not np.isnan(row["median"]):
            print(f"    {year}: медиана = {row['median']:+.1f}%, N = {int(row['count'])}")

    return {
        "hypothesis": "H1",
        "test": "Вилкоксон (односторонний)",
        "statistic": stat_w,
        "p_value": p_w,
        "result": "Подтверждена" if p_w < 0.05 else "Не подтверждена",
        "detail": f"Медиана темпа роста = {growth.median():.1f}%",
    }


# ── H2: Масштаб бизнеса vs рентабельность ──────────────────────────

def test_h2(df: pd.DataFrame):
    print("\n" + "=" * 60)
    print("H2: Зависимость между выручкой и рентабельностью по ЧП")
    print()

    data = df[["Выручка", "Рентабельность по ЧП, %"]].dropna()
    print(f"  N = {len(data)}")

    # Спирмен (выручка ненормальна, рентабельность нормальна → Спирмен надёжнее)
    rho, p_s = stats.spearmanr(data["Выручка"], data["Рентабельность по ЧП, %"])
    print(f"\n  Корреляция Спирмена:")
    print(f"    ρ = {rho:.4f}, p = {p_s:.6f}")
    print(f"    → {'ЗНАЧИМА' if p_s < 0.05 else 'НЕ ЗНАЧИМА'} (α = 0.05)")

    # Пирсон для сравнения
    r, p_p = stats.pearsonr(data["Выручка"], data["Рентабельность по ЧП, %"])
    print(f"\n  Корреляция Пирсона (для сравнения):")
    print(f"    r = {r:.4f}, p = {p_p:.6f}")

    return {
        "hypothesis": "H2",
        "test": "Корреляция Спирмена",
        "statistic": rho,
        "p_value": p_s,
        "result": "Подтверждена" if p_s < 0.05 else "Не подтверждена",
        "detail": f"ρ = {rho:.3f}",
    }


# ── H3: Доля себестоимости снижается с ростом выручки ───────────────

def test_h3(df: pd.DataFrame):
    print("\n" + "=" * 60)
    print("H3: Доля себестоимости снижается с ростом выручки")
    print()

    data = df[["Выручка", "Доля себестоимости, %"]].dropna()
    print(f"  N = {len(data)}")

    rho, p = stats.spearmanr(data["Выручка"], data["Доля себестоимости, %"])
    print(f"\n  Корреляция Спирмена:")
    print(f"    ρ = {rho:.4f}, p = {p:.6f}")
    print(f"    Направление: {'отрицательная (ожидаемо)' if rho < 0 else 'положительная (неожиданно)'}")
    print(f"    → {'ЗНАЧИМА' if p < 0.05 else 'НЕ ЗНАЧИМА'} (α = 0.05)")

    return {
        "hypothesis": "H3",
        "test": "Корреляция Спирмена",
        "statistic": rho,
        "p_value": p,
        "result": "Подтверждена" if (p < 0.05 and rho < 0) else "Не подтверждена",
        "detail": f"ρ = {rho:.3f}, направление {'−' if rho < 0 else '+'}",
    }


# ── H5: Рост выручки сопровождается ростом КЗ ─────────────────────

def test_h5(df: pd.DataFrame):
    print("\n" + "=" * 60)
    print("H5: Рост выручки сопровождается ростом кредиторской задолженности")
    print()

    data = df[["Темп роста Выручка, %", "Темп роста Кредиторская задолженность, %"]].dropna()
    print(f"  N = {len(data)}")

    rho, p = stats.spearmanr(
        data["Темп роста Выручка, %"],
        data["Темп роста Кредиторская задолженность, %"],
    )
    print(f"\n  Корреляция Спирмена:")
    print(f"    ρ = {rho:.4f}, p = {p:.6f}")
    print(f"    → {'ЗНАЧИМА' if p < 0.05 else 'НЕ ЗНАЧИМА'} (α = 0.05)")

    return {
        "hypothesis": "H5",
        "test": "Корреляция Спирмена",
        "statistic": rho,
        "p_value": p,
        "result": "Подтверждена" if (p < 0.05 and rho > 0) else "Не подтверждена",
        "detail": f"ρ = {rho:.3f}",
    }


# ── Корреляционная матрица ─────────────────────────────────────────

def correlation_matrix(df: pd.DataFrame):
    print("\n" + "=" * 60)
    print("Корреляционная матрица (Спирмен)")

    corr_cols = [
        "Выручка", "Себестоимость продаж", "Чистая прибыль (убыток)",
        "Активы всего", "Денежные средства", "Кредиторская задолженность",
        "Валовая рентабельность, %", "Рентабельность по ЧП, %",
        "Оборачиваемость активов",
    ]

    data = df[corr_cols].dropna()
    print(f"  N = {len(data)} (полных наблюдений)")

    # Спирмен
    corr_matrix = data.corr(method="spearman")

    # p-values
    n = len(corr_cols)
    pval_matrix = pd.DataFrame(np.ones((n, n)), index=corr_cols, columns=corr_cols)
    for i in range(n):
        for j in range(i + 1, n):
            vals = df[[corr_cols[i], corr_cols[j]]].dropna()
            if len(vals) >= 3:
                _, p = stats.spearmanr(vals.iloc[:, 0], vals.iloc[:, 1])
                pval_matrix.iloc[i, j] = p
                pval_matrix.iloc[j, i] = p

    # Save
    corr_path = OUT / "correlation_matrix.xlsx"
    with pd.ExcelWriter(corr_path, engine="xlsxwriter") as writer:
        corr_matrix.to_excel(writer, sheet_name="Корреляция Спирмена")
        pval_matrix.to_excel(writer, sheet_name="p-values")
    print(f"  Сохранено: {corr_path}")

    # Significant correlations
    print("\n  Значимые корреляции (|ρ| > 0.3, p < 0.05):")
    for i in range(n):
        for j in range(i + 1, n):
            rho = corr_matrix.iloc[i, j]
            p = pval_matrix.iloc[i, j]
            if abs(rho) > 0.3 and p < 0.05:
                print(f"    {corr_cols[i]:35s} ↔ {corr_cols[j]:35s}  ρ={rho:+.3f}  p={p:.4f}")

    # Heatmap
    short_labels = [c.replace("Чистая прибыль (убыток)", "Чист. прибыль")
                     .replace("Кредиторская задолженность", "Кредит. задолж.")
                     .replace("Себестоимость продаж", "Себестоимость")
                     .replace("Валовая рентабельность, %", "Вал. рент., %")
                     .replace("Рентабельность по ЧП, %", "Рент. по ЧП, %")
                     .replace("Оборачиваемость активов", "Оборач. активов")
                     .replace("Денежные средства", "Ден. средства")
                     .replace("Активы всего", "Активы")
                    for c in corr_cols]

    fig, ax = plt.subplots(figsize=(10, 8))
    mask = np.triu(np.ones_like(corr_matrix, dtype=bool), k=1)
    sns.heatmap(
        corr_matrix, mask=mask, annot=True, fmt=".2f", cmap="RdBu_r",
        center=0, vmin=-1, vmax=1, ax=ax, square=True,
        xticklabels=short_labels, yticklabels=short_labels,
        annot_kws={"size": 9},
    )
    ax.set_title("Корреляционная матрица (Спирмен)", fontsize=12, pad=15)
    fig.tight_layout()
    fig.savefig(PLOTS / "correlation_heatmap.png")
    plt.close(fig)
    print(f"  Heatmap: {PLOTS / 'correlation_heatmap.png'}")

    return corr_matrix, pval_matrix


# ── Scatter-графики для гипотез ────────────────────────────────────

def plot_hypothesis_scatters(df: pd.DataFrame):
    fig, axes = plt.subplots(1, 3, figsize=(18, 5))

    # H2: Выручка vs Рентабельность по ЧП
    data = df[["Выручка", "Рентабельность по ЧП, %"]].dropna()
    axes[0].scatter(data["Выручка"] / 1000, data["Рентабельность по ЧП, %"], alpha=0.7, s=40)
    axes[0].set_xlabel("Выручка, млн руб.")
    axes[0].set_ylabel("Рентабельность по ЧП, %")
    axes[0].set_title("H2: Масштаб vs Рентабельность")
    axes[0].axhline(0, color="grey", linewidth=0.5)
    axes[0].grid(True, alpha=0.3)

    # H3: Выручка vs Доля себестоимости
    data = df[["Выручка", "Доля себестоимости, %"]].dropna()
    axes[1].scatter(data["Выручка"] / 1000, data["Доля себестоимости, %"], alpha=0.7, s=40, color="#DD8452")
    axes[1].set_xlabel("Выручка, млн руб.")
    axes[1].set_ylabel("Доля себестоимости, %")
    axes[1].set_title("H3: Масштаб vs Доля себестоимости")
    axes[1].axhline(100, color="red", linewidth=0.5, linestyle="--")
    axes[1].grid(True, alpha=0.3)

    # H5: Рост выручки vs Рост КЗ
    data = df[["Темп роста Выручка, %", "Темп роста Кредиторская задолженность, %"]].dropna()
    axes[2].scatter(data["Темп роста Выручка, %"], data["Темп роста Кредиторская задолженность, %"], alpha=0.7, s=40, color="#55A868")
    axes[2].set_xlabel("Темп роста выручки, %")
    axes[2].set_ylabel("Темп роста КЗ, %")
    axes[2].set_title("H5: Рост выручки vs Рост КЗ")
    axes[2].axhline(0, color="grey", linewidth=0.5)
    axes[2].axvline(0, color="grey", linewidth=0.5)
    axes[2].grid(True, alpha=0.3)

    fig.suptitle("Визуализация гипотез H2, H3, H5", fontsize=13, y=1.02)
    fig.tight_layout()
    fig.savefig(PLOTS / "hypothesis_scatters.png")
    plt.close(fig)
    print(f"\n  Scatter-графики: {PLOTS / 'hypothesis_scatters.png'}")


# ── Хи-квадрат: ОПФ vs прибыльность ───────────────────────────────

def test_chi2_opf(df: pd.DataFrame):
    print("\n" + "=" * 60)
    print("Доп. тест: Хи-квадрат (ОПФ vs прибыльность)")

    data = df[["ОПФ", "Чистая прибыль (убыток)"]].dropna()
    data["Прибыльная"] = (data["Чистая прибыль (убыток)"] > 0).map({True: "Да", False: "Нет"})

    ct = pd.crosstab(data["ОПФ"], data["Прибыльная"])
    print(f"\n  Таблица сопряжённости:")
    print(f"  {ct.to_string()}")

    if ct.shape == (2, 2) and (ct.values >= 5).all():
        chi2, p, dof, expected = stats.chi2_contingency(ct)
        print(f"\n  χ² = {chi2:.3f}, df = {dof}, p = {p:.4f}")
        print(f"  → {'ЗНАЧИМА' if p < 0.05 else 'НЕ ЗНАЧИМА'} (α = 0.05)")
    else:
        stat, p = stats.fisher_exact(ct)
        print(f"\n  Точный тест Фишера (малые частоты):")
        print(f"  OR = {stat:.3f}, p = {p:.4f}")
        print(f"  → {'ЗНАЧИМА' if p < 0.05 else 'НЕ ЗНАЧИМА'} (α = 0.05)")


def main():
    df = load()
    rev_df = df[df["Выручка"].notna()].copy()

    results = []
    results.append(test_h1(rev_df))
    results.append(test_h2(rev_df))
    results.append(test_h3(rev_df))
    results.append(test_h5(rev_df))

    correlation_matrix(rev_df)
    plot_hypothesis_scatters(rev_df)
    test_chi2_opf(rev_df)

    # Summary
    print("\n" + "=" * 60)
    print("СВОДКА ПО ГИПОТЕЗАМ")
    summary = pd.DataFrame(results)
    print(summary.to_string(index=False))

    summary_path = OUT / "hypothesis_tests.xlsx"
    summary.to_excel(summary_path, index=False, engine="xlsxwriter")
    print(f"\nСохранено: {summary_path}")


if __name__ == "__main__":
    main()
