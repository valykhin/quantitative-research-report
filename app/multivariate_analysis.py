"""Этап 6. Многофакторный анализ (регрессия, кластерный анализ)."""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
import seaborn as sns
from scipy import stats
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler

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


# ── 6.1 Регрессионный анализ ───────────────────────────────────────

def regression_analysis(df: pd.DataFrame):
    print("=" * 60)
    print("6.1  Регрессионный анализ")
    print("      Зависимая: Валовая рентабельность, %")
    print("      Предиктор: ln(Выручка)")

    data = df[["Выручка", "Валовая рентабельность, %"]].dropna()
    data = data[data["Выручка"] > 0].copy()
    data["ln_Выручка"] = np.log(data["Выручка"])

    x = data["ln_Выручка"].values
    y = data["Валовая рентабельность, %"].values
    n = len(x)

    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
    r_sq = r_value ** 2
    y_pred = intercept + slope * x
    residuals = y - y_pred

    print(f"\n  N = {n}")
    print(f"  Коэффициенты:")
    print(f"    Intercept (a) = {intercept:.4f}")
    print(f"    Slope (b)     = {slope:.4f}  (p = {p_value:.6f})")
    print(f"  R² = {r_sq:.4f}")
    print(f"  Стд. ошибка slope = {std_err:.4f}")

    # Проверка условий МНК
    print(f"\n  Проверка условий:")

    # 1. Нормальность остатков
    w_stat, w_p = stats.shapiro(residuals)
    print(f"    Нормальность остатков (Шапиро-Уилк): W={w_stat:.4f}, p={w_p:.4f}"
          f" → {'✓' if w_p > 0.05 else '✗'}")

    # 2. Среднее остатков ≈ 0
    print(f"    Среднее остатков: {residuals.mean():.6f} (≈ 0)")

    # 3. Корреляция остатков с предиктором
    r_res, p_res = stats.pearsonr(x, residuals)
    print(f"    Корреляция остатков с предиктором: r={r_res:.6f}, p={p_res:.4f}"
          f" → {'✓' if p_res > 0.05 else '✗'}")

    # Визуализация
    fig, axes = plt.subplots(1, 3, figsize=(18, 5))

    # Scatter + линия регрессии
    axes[0].scatter(x, y, alpha=0.7, s=40)
    x_line = np.linspace(x.min(), x.max(), 100)
    axes[0].plot(x_line, intercept + slope * x_line, color="red", linewidth=2,
                 label=f"y = {intercept:.1f} + {slope:.2f}·x\nR² = {r_sq:.3f}, p = {p_value:.4f}")
    axes[0].set_xlabel("ln(Выручка)")
    axes[0].set_ylabel("Валовая рентабельность, %")
    axes[0].set_title("Регрессия: рентабельность от ln(выручки)")
    axes[0].legend(fontsize=8)
    axes[0].grid(True, alpha=0.3)

    # Гистограмма остатков
    axes[1].hist(residuals, bins=10, edgecolor="white", alpha=0.8, color="#4C72B0")
    axes[1].set_title("Гистограмма остатков")
    axes[1].set_xlabel("Остатки")

    # Остатки vs предиктор
    axes[2].scatter(x, residuals, alpha=0.7, s=40, color="#55A868")
    axes[2].axhline(0, color="red", linewidth=1)
    axes[2].set_xlabel("ln(Выручка)")
    axes[2].set_ylabel("Остатки")
    axes[2].set_title("Остатки vs предиктор")
    axes[2].grid(True, alpha=0.3)

    fig.suptitle("Регрессионный анализ", fontsize=13, y=1.02)
    fig.tight_layout()
    fig.savefig(PLOTS / "regression.png")
    plt.close(fig)
    print(f"\n  Графики: {PLOTS / 'regression.png'}")

    results = {
        "intercept": intercept, "slope": slope, "r_squared": r_sq,
        "p_value": p_value, "std_err": std_err,
        "residuals_normality_p": w_p, "n": n,
    }

    res_df = pd.DataFrame([{
        "Предиктор": "ln(Выручка)",
        "Коэффициент": slope, "Стд. ошибка": std_err,
        "t-статистика": slope / std_err, "p-value": p_value,
        "R²": r_sq, "N": n,
        "Остатки: нормальность (p)": w_p,
    }])
    res_path = OUT / "regression_results.xlsx"
    res_df.to_excel(res_path, index=False, engine="xlsxwriter")
    print(f"  Результаты: {res_path}")

    return results


# ── 6.2 Кластерный анализ ──────────────────────────────────────────

def cluster_analysis(df: pd.DataFrame):
    print("\n" + "=" * 60)
    print("6.2  Кластерный анализ (H4)")

    # Берём последний доступный период по каждой компании (с выручкой)
    rev_df = df[df["Выручка"].notna()].copy()
    latest = rev_df.sort_values("Период").groupby("Компания").last().reset_index()

    cluster_vars = [
        "Выручка", "Валовая рентабельность, %",
        "Оборачиваемость активов", "Доля КЗ в активах, %",
        "Доля ДС в активах, %",
    ]

    data = latest[["Компания"] + cluster_vars].dropna()
    print(f"  Компаний для кластеризации: {len(data)}")
    print(f"  Переменные: {cluster_vars}")

    X = data[cluster_vars].values
    company_names = data["Компания"].values

    # Стандартизация
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    # Оптимальное число кластеров
    max_k = min(len(data) - 1, 6)
    inertias = []
    silhouettes = []

    for k in range(2, max_k + 1):
        km = KMeans(n_clusters=k, n_init=20, random_state=42)
        labels = km.fit_predict(X_scaled)
        inertias.append(km.inertia_)
        sil = silhouette_score(X_scaled, labels)
        silhouettes.append(sil)
        print(f"    k={k}: инерция={km.inertia_:.2f}, силуэт={sil:.3f}")

    best_k = 2 + np.argmax(silhouettes)
    print(f"\n  Оптимальное k = {best_k} (макс. силуэт = {max(silhouettes):.3f})")

    # Финальная кластеризация
    km_final = KMeans(n_clusters=best_k, n_init=20, random_state=42)
    labels = km_final.fit_predict(X_scaled)

    data = data.copy()
    data["Кластер"] = labels

    # Интерпретация кластеров
    print(f"\n  Состав кластеров:")
    for cl in sorted(data["Кластер"].unique()):
        members = data[data["Кластер"] == cl]
        names = ", ".join(members["Компания"].apply(
            lambda x: x.replace("ООО ", "").replace("АНО ", "").strip('"')
        ))
        print(f"    Кластер {cl}: {names}")

    print(f"\n  Средние значения по кластерам:")
    cluster_means = data.groupby("Кластер")[cluster_vars].mean()
    print(cluster_means.to_string())

    # PCA для визуализации
    pca = PCA(n_components=2)
    X_pca = pca.fit_transform(X_scaled)
    var_explained = pca.explained_variance_ratio_

    # Визуализация
    fig, axes = plt.subplots(1, 3, figsize=(18, 6))

    # Метод локтя + силуэт
    ks = list(range(2, max_k + 1))
    axes[0].plot(ks, inertias, "o-", color="#4C72B0", label="Инерция")
    ax_sil = axes[0].twinx()
    ax_sil.plot(ks, silhouettes, "s-", color="#DD8452", label="Силуэт")
    axes[0].set_xlabel("Число кластеров (k)")
    axes[0].set_ylabel("Инерция", color="#4C72B0")
    ax_sil.set_ylabel("Силуэт", color="#DD8452")
    axes[0].set_title("Выбор числа кластеров")
    axes[0].xaxis.set_major_locator(mticker.MaxNLocator(integer=True))

    # Scatter PCA
    colors = ["#4C72B0", "#DD8452", "#55A868", "#C44E52"]
    for cl in sorted(data["Кластер"].unique()):
        mask = labels == cl
        axes[1].scatter(X_pca[mask, 0], X_pca[mask, 1], s=80, alpha=0.8,
                        color=colors[cl % len(colors)], label=f"Кластер {cl}", zorder=3)
    for i, name in enumerate(company_names):
        short = name.replace("ООО ", "").replace("АНО ", "").replace("АССОЦИАЦИЯ ", "").strip('"')
        axes[1].annotate(short, (X_pca[i, 0], X_pca[i, 1]),
                         fontsize=7, ha="center", va="bottom", textcoords="offset points", xytext=(0, 6))
    axes[1].set_xlabel(f"PC1 ({var_explained[0]*100:.1f}%)")
    axes[1].set_ylabel(f"PC2 ({var_explained[1]*100:.1f}%)")
    axes[1].set_title("Кластеры в пространстве главных компонент")
    axes[1].legend(fontsize=8)
    axes[1].grid(True, alpha=0.3)

    # Профиль кластеров (radar-like bar chart)
    cluster_means_scaled = pd.DataFrame(
        scaler.transform(cluster_means.values),
        columns=cluster_vars, index=cluster_means.index,
    )
    short_vars = [v.replace("Валовая рентабельность, %", "Вал. рент.")
                   .replace("Оборачиваемость активов", "Оборач.")
                   .replace("Доля КЗ в активах, %", "Доля КЗ")
                   .replace("Доля ДС в активах, %", "Доля ДС")
                  for v in cluster_vars]

    x_pos = np.arange(len(cluster_vars))
    width = 0.35
    for i, cl in enumerate(sorted(cluster_means_scaled.index)):
        axes[2].bar(x_pos + i * width, cluster_means_scaled.loc[cl].values,
                    width, label=f"Кластер {cl}", color=colors[cl % len(colors)], alpha=0.8)
    axes[2].set_xticks(x_pos + width * (best_k - 1) / 2)
    axes[2].set_xticklabels(short_vars, fontsize=8, rotation=15)
    axes[2].set_ylabel("Стандартизованное значение")
    axes[2].set_title("Профиль кластеров")
    axes[2].legend(fontsize=8)
    axes[2].axhline(0, color="grey", linewidth=0.5)
    axes[2].grid(True, alpha=0.3, axis="y")

    fig.suptitle(f"Кластерный анализ (k={best_k})", fontsize=13, y=1.02)
    fig.tight_layout()
    fig.savefig(PLOTS / "clusters.png")
    plt.close(fig)
    print(f"\n  Графики: {PLOTS / 'clusters.png'}")

    # Сохранение
    cluster_path = OUT / "cluster_results.xlsx"
    with pd.ExcelWriter(cluster_path, engine="xlsxwriter") as writer:
        data[["Компания", "Кластер"] + cluster_vars].to_excel(writer, sheet_name="Принадлежность", index=False)
        cluster_means.to_excel(writer, sheet_name="Средние по кластерам")
    print(f"  Результаты: {cluster_path}")

    return {
        "best_k": best_k,
        "best_silhouette": max(silhouettes),
        "cluster_data": data,
        "pca_variance": var_explained,
    }


def main():
    df = load()
    rev_df = df[df["Выручка"].notna()].copy()

    reg = regression_analysis(rev_df)
    cl = cluster_analysis(rev_df)

    print("\n" + "=" * 60)
    print("СВОДКА")
    print(f"  Регрессия: R² = {reg['r_squared']:.3f}, p = {reg['p_value']:.4f}"
          f" → {'значима' if reg['p_value'] < 0.05 else 'незначима'}")
    print(f"  Кластеры: k = {cl['best_k']}, силуэт = {cl['best_silhouette']:.3f}")
    print(f"  H4: {'Подтверждена' if cl['best_silhouette'] > 0.3 else 'Не подтверждена'}"
          f" (силуэт {'>' if cl['best_silhouette'] > 0.3 else '≤'} 0.3)")


if __name__ == "__main__":
    main()
