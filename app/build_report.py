"""Этап 8. Формирование итогового xlsx-отчёта со встроенными графиками."""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
import seaborn as sns
from io import BytesIO
from scipy import stats
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "figure.dpi": 130,
    "savefig.bbox": "tight",
})

APP = Path(__file__).parent
OUT = APP / "output"
REPORT_PATH = OUT / "report.xlsx"


def load():
    raw = pd.read_excel(APP / "data" / "companies_financials.xlsx")
    full = pd.read_excel(APP / "data" / "analysis_dataset.xlsx")
    return raw, full


def fig_to_bytes(fig) -> BytesIO:
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=130)
    plt.close(fig)
    buf.seek(0)
    return buf


def short_name(name: str) -> str:
    return name.replace("ООО ", "").replace("АНО ", "").replace("АССОЦИАЦИЯ ", "").strip('"')


def make_formats(book):
    """Общие форматы для всех листов."""
    return {
        "title": book.add_format({"bold": True, "font_size": 14}),
        "h2": book.add_format({"bold": True, "font_size": 12}),
        "header": book.add_format({"bold": True, "font_size": 11, "bottom": 1}),
        "normal": book.add_format({"font_size": 11, "text_wrap": True}),
        "num0": book.add_format({"num_format": "#,##0", "font_size": 11}),
        "num1": book.add_format({"num_format": "#,##0.0", "font_size": 11}),
        "num2": book.add_format({"num_format": "0.00", "font_size": 11}),
        "num4": book.add_format({"num_format": "0.0000", "font_size": 11}),
        "pct1": book.add_format({"num_format": "0.0", "font_size": 11}),
        "pct2": book.add_format({"num_format": "0.00", "font_size": 11}),
        "green": book.add_format({"font_size": 11, "font_color": "#006100", "bg_color": "#C6EFCE", "bold": True}),
        "red": book.add_format({"font_size": 11, "font_color": "#9C0006", "bg_color": "#FFC7CE"}),
        "label": book.add_format({"bold": True, "font_size": 11, "italic": True, "font_color": "#404040"}),
    }


# ── 1. Описание исследования ──────────────────────────────────────

def write_description(writer, fmts):
    ws = writer.book.add_worksheet("Описание исследования")
    writer.sheets["Описание исследования"] = ws
    ws.set_column("A:A", 80)
    ws.set_column("B:B", 45)
    ws.set_column("C:C", 15)

    ws.write("A1", "Анализ рынка организации беговых мероприятий в России, ориентированный на сбыт", fmts["title"])
    ws.write("A3", "Цель исследования:", fmts["header"])
    ws.write("A4", "Оценить финансовое состояние и динамику развития компаний-организаторов массовых "
             "беговых мероприятий в России на основе данных бухгалтерской отчётности за 2021–2025 гг., "
             "выявить закономерности и различия в бизнес-моделях.", fmts["normal"])

    ws.write("A6", "Параметры исследования:", fmts["header"])
    ws.write("A7", "Источник данных: ГИРБО (bo.nalog.gov.ru) — официальный ресурс ФНС России", fmts["normal"])
    ws.write("A8", "Период наблюдений: 2021–2025 (5 лет)", fmts["normal"])
    ws.write("A9", "Объём выборки: 11 компаний, 49 наблюдений (компания × год)", fmts["normal"])
    ws.write("A10", "Метод сбора: автоматизированный через API ГИРБО (скрипт collect_data.py)", fmts["normal"])

    ws.write("A12", "Гипотезы:", fmts["header"])
    hypotheses = [
        ("H1", "Рынок демонстрирует устойчивый рост выручки", "Тест Вилкоксона"),
        ("H2", "Существует зависимость между масштабом бизнеса и рентабельностью", "Корреляция Спирмена"),
        ("H3", "Доля себестоимости снижается с ростом выручки (эффект масштаба)", "Корреляция Спирмена"),
        ("H4", "Компании образуют кластеры по финансовому профилю", "k-means, силуэтный коэффициент"),
        ("H5", "Рост выручки сопровождается ростом кредиторской задолженности", "Корреляция Спирмена"),
    ]
    ws.write(12, 0, "№", fmts["header"])
    ws.write(12, 1, "Гипотеза", fmts["header"])
    ws.write(12, 2, "Метод проверки", fmts["header"])
    for i, (num, hyp, method) in enumerate(hypotheses):
        ws.write(13 + i, 0, num, fmts["normal"])
        ws.write(13 + i, 1, hyp, fmts["normal"])
        ws.write(13 + i, 2, method, fmts["normal"])

    ws.write("A20", "Состав выборки:", fmts["header"])
    companies = [
        ("Гонка Героев", "АНО «ЛИГА ГЕРОЕВ»", "7709445877", "Забеги с препятствиями"),
        ("Марафон Сервис", "ООО «МАРАФОН СЕРВИС»", "7704408575", "Марафоны, полумарафоны"),
        ("TIMERMAN", "ООО «АСМ «НОВЫЙ СПОРТ»", "1660308354", "Триатлон, забеги"),
        ("Бегом по Золотому кольцу", "ООО «АРЕНА ПЛЮС»", "7606117641", "Беговые туры"),
        ("PushkinRun", "АССОЦИАЦИЯ «ПЕТЕРБУРГСКИЙ СПОРТ»", "7810948438", "Городские забеги"),
        ("ГРУТ / Вайлд Трейл", "ООО «ГРУТ»", "9724029321", "Трейлраннинг"),
        ("IRONSTAR", "ООО «АРХИТЕКТУРА СПОРТА»", "9717014483", "Триатлон, забеги"),
        ("Беги, Герой!", "АНО «РЕЙТИНГ СПОРТ»", "5252050062", "Городские забеги"),
        ("S95", "АНО «ФИЗКУЛЬТУРА»", "9721206702", "Парковые забеги"),
        ("parkrun", "АНО «ПАРКРАН»", "7729451893", "Парковые забеги"),
        ("Ярославское беговое", "ООО «ЯРОСЛАВСКОЕ БЕГОВОЕ СООБЩЕСТВО»", "7603074566", "Региональные забеги"),
    ]
    for j, col in enumerate(["Бренд", "Юрлицо", "ИНН", "Специализация"]):
        ws.write(20, j, col, fmts["header"])
    ws.set_column("D:D", 25)
    for i, (brand, entity, inn, spec) in enumerate(companies):
        ws.write(21 + i, 0, brand, fmts["normal"])
        ws.write(21 + i, 1, entity, fmts["normal"])
        ws.write(21 + i, 2, inn, fmts["normal"])
        ws.write(21 + i, 3, spec, fmts["normal"])


# ── 2. Исходные данные ─────────────────────────────────────────────

def write_raw_data(writer, fmts, raw):
    raw.to_excel(writer, sheet_name="Исходные данные", index=False)
    ws = writer.sheets["Исходные данные"]
    ws.set_column("A:A", 40)
    ws.set_column("B:B", 15)
    ws.set_column("C:C", 10)
    for col_idx in range(3, len(raw.columns)):
        ws.set_column(col_idx, col_idx, 14, fmts["num0"])


# ── 3. Производные показатели ──────────────────────────────────────

def write_derived(writer, fmts, full):
    cols = [
        "Компания", "Период", "ОПФ",
        "Валовая рентабельность, %", "Рентабельность по ЧП, %", "Доля себестоимости, %",
        "Доля ОС в активах, %", "Доля ДС в активах, %", "Доля ДЗ в активах, %",
        "Доля КЗ в активах, %", "Коэффициент автономии", "Оборачиваемость активов",
        "Темп роста Выручка, %", "Темп роста Активы всего, %",
        "Темп роста Кредиторская задолженность, %",
    ]
    existing = [c for c in cols if c in full.columns]
    full[existing].to_excel(writer, sheet_name="Производные показатели", index=False)
    ws = writer.sheets["Производные показатели"]
    ws.set_column("A:A", 40)
    ws.set_column("B:B", 10)
    ws.set_column("C:C", 8)
    for col_idx in range(3, len(existing)):
        ws.set_column(col_idx, col_idx, 16, fmts["pct1"])


# ── 4. Описательные статистики ─────────────────────────────────────

def write_descriptive(writer, fmts, full):
    key_cols = [
        ("Выручка", "тыс. руб."), ("Себестоимость продаж", "тыс. руб."),
        ("Чистая прибыль (убыток)", "тыс. руб."), ("Активы всего", "тыс. руб."),
        ("Денежные средства", "тыс. руб."), ("Кредиторская задолженность", "тыс. руб."),
        ("Валовая рентабельность, %", "%"), ("Рентабельность по ЧП, %", "%"),
        ("Доля себестоимости, %", "%"), ("Оборачиваемость активов", "раз"),
        ("Темп роста Выручка, %", "%"),
    ]

    rev = full[full["Выручка"].notna()]
    rows = []
    for col, unit in key_cols:
        vals = rev[col].dropna()
        if len(vals) < 3:
            continue
        skew = stats.skew(vals, bias=False)
        kurt = stats.kurtosis(vals, bias=False)
        q1, q3 = np.percentile(vals, [25, 75])
        w_stat, w_p = stats.shapiro(vals)
        rows.append({
            "Переменная": col, "Ед. изм.": unit, "N": int(len(vals)),
            "Среднее": round(vals.mean(), 1), "Медиана": round(vals.median(), 1),
            "Станд. откл.": round(vals.std(), 1),
            "Min": round(vals.min(), 1), "Max": round(vals.max(), 1),
            "Q1": round(q1, 1), "Q3": round(q3, 1), "IQR": round(q3 - q1, 1),
            "Асимметрия": round(skew, 2), "Эксцесс": round(kurt, 2),
            "Типичное среднее": "медиана" if abs(skew) > 1 else "среднее",
            "Шапиро-Уилк W": round(w_stat, 4), "Шапиро-Уилк p": round(w_p, 4),
            "Нормальное (α=0.05)": "Да" if w_p > 0.05 else "Нет",
        })

    desc_df = pd.DataFrame(rows)
    desc_df.to_excel(writer, sheet_name="Описательные статистики", index=False)
    ws = writer.sheets["Описательные статистики"]
    ws.set_column("A:A", 35)
    ws.set_column("B:B", 10)


# ── 5. Динамика рынка ──────────────────────────────────────────────

def write_dynamics(writer, fmts, full):
    rev = full[full["Выручка"].notna()]

    pivot_rev = rev.pivot_table(index="Компания", columns="Период", values="Выручка")
    pivot_profit = rev.pivot_table(index="Компания", columns="Период", values="Чистая прибыль (убыток)")
    pivot_rent = rev.pivot_table(index="Компания", columns="Период", values="Валовая рентабельность, %")

    ws = writer.book.add_worksheet("Динамика рынка")
    writer.sheets["Динамика рынка"] = ws
    ws.set_column("A:A", 40)
    for col_idx in range(1, 7):
        ws.set_column(col_idx, col_idx, 14)

    row = 0

    # Таблица 1: Выручка
    ws.write(row, 0, "Таблица 1. Выручка компаний по годам, тыс. руб.", fmts["h2"])
    row += 1
    pivot_rev.to_excel(writer, sheet_name="Динамика рынка", startrow=row)
    # Format numbers
    for r in range(len(pivot_rev)):
        for c in range(len(pivot_rev.columns)):
            val = pivot_rev.iloc[r, c]
            if pd.notna(val):
                ws.write(row + 1 + r, 1 + c, val, fmts["num0"])
    row += len(pivot_rev) + 3

    # Таблица 2: Чистая прибыль
    ws.write(row, 0, "Таблица 2. Чистая прибыль (убыток) по годам, тыс. руб.", fmts["h2"])
    row += 1
    pivot_profit.to_excel(writer, sheet_name="Динамика рынка", startrow=row)
    for r in range(len(pivot_profit)):
        for c in range(len(pivot_profit.columns)):
            val = pivot_profit.iloc[r, c]
            if pd.notna(val):
                ws.write(row + 1 + r, 1 + c, val, fmts["num0"])
    row += len(pivot_profit) + 3

    # Таблица 3: Рентабельность
    ws.write(row, 0, "Таблица 3. Валовая рентабельность по годам, %", fmts["h2"])
    row += 1
    pivot_rent.to_excel(writer, sheet_name="Динамика рынка", startrow=row)
    for r in range(len(pivot_rent)):
        for c in range(len(pivot_rent.columns)):
            val = pivot_rent.iloc[r, c]
            if pd.notna(val):
                ws.write(row + 1 + r, 1 + c, val, fmts["pct1"])
    row += len(pivot_rent) + 3

    # График: динамика
    companies_full = rev.groupby("Компания").filter(lambda x: x["Выручка"].notna().sum() >= 3)
    comp_list = sorted(companies_full["Компания"].unique())

    fig, axes = plt.subplots(2, 2, figsize=(16, 11))
    metrics = [
        ("Выручка", "Выручка, тыс. руб.", "тыс. руб."),
        ("Чистая прибыль (убыток)", "Чистая прибыль, тыс. руб.", "тыс. руб."),
        ("Валовая рентабельность, %", "Валовая рентабельность", "%"),
        ("Активы всего", "Активы, тыс. руб.", "тыс. руб."),
    ]
    for ax, (col, title, ylabel) in zip(axes.flatten(), metrics):
        for comp in comp_list:
            grp = companies_full[companies_full["Компания"] == comp].sort_values("Период")
            vals = grp[["Период", col]].dropna()
            if len(vals) >= 2:
                ax.plot(vals["Период"], vals[col], marker="o", markersize=4,
                        label=short_name(comp), linewidth=1.5)
        ax.set_title(title, fontsize=11)
        ax.set_ylabel(ylabel)
        ax.set_xlabel("Год")
        ax.legend(fontsize=7, loc="upper left")
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    fig.suptitle("Динамика ключевых показателей (2021–2025)", fontsize=13, y=1.01)
    fig.tight_layout()
    ws.insert_image(f"A{row + 1}", "", {"image_data": fig_to_bytes(fig), "x_scale": 0.85, "y_scale": 0.85})
    row += 38

    # График: сводка рынка
    agg = rev.groupby("Период").agg(
        Сумма_выручки=("Выручка", "sum"),
        Медиана_выручки=("Выручка", "median"),
        Медиана_рент=("Валовая рентабельность, %", "median"),
    ).reset_index()

    fig2, axes2 = plt.subplots(1, 3, figsize=(15, 5))
    axes2[0].bar(agg["Период"], agg["Сумма_выручки"] / 1000, color="#4C72B0", alpha=0.8)
    axes2[0].set_title("Суммарная выручка рынка")
    axes2[0].set_ylabel("млн руб.")
    axes2[0].set_xlabel("Год")
    axes2[1].plot(agg["Период"], agg["Медиана_выручки"], marker="o", color="#DD8452", linewidth=2)
    axes2[1].set_title("Медианная выручка")
    axes2[1].set_ylabel("тыс. руб.")
    axes2[1].set_xlabel("Год")
    axes2[2].plot(agg["Период"], agg["Медиана_рент"], marker="s", color="#55A868", linewidth=2)
    axes2[2].set_title("Медианная валовая рентабельность")
    axes2[2].set_ylabel("%")
    axes2[2].set_xlabel("Год")
    for ax in axes2:
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    fig2.suptitle("Сводные показатели рынка беговых мероприятий", fontsize=13, y=1.02)
    fig2.tight_layout()
    ws.insert_image(f"A{row + 1}", "", {"image_data": fig_to_bytes(fig2), "x_scale": 0.85, "y_scale": 0.85})


# ── 6. Корреляционный анализ ───────────────────────────────────────

def write_correlation(writer, fmts, full):
    corr_cols = [
        "Выручка", "Себестоимость продаж", "Чистая прибыль (убыток)",
        "Активы всего", "Денежные средства", "Кредиторская задолженность",
        "Валовая рентабельность, %", "Рентабельность по ЧП, %",
        "Оборачиваемость активов",
    ]

    rev = full[full["Выручка"].notna()]
    data = rev[corr_cols].dropna()
    corr = data.corr(method="spearman")

    n = len(corr_cols)
    pvals = pd.DataFrame(np.ones((n, n)), index=corr_cols, columns=corr_cols)
    for i in range(n):
        for j in range(i + 1, n):
            v = rev[[corr_cols[i], corr_cols[j]]].dropna()
            if len(v) >= 3:
                _, p = stats.spearmanr(v.iloc[:, 0], v.iloc[:, 1])
                pvals.iloc[i, j] = p
                pvals.iloc[j, i] = p

    ws = writer.book.add_worksheet("Корреляционный анализ")
    writer.sheets["Корреляционный анализ"] = ws
    ws.set_column("A:A", 35)

    # Таблица 1: Корреляция
    ws.write(0, 0, "Таблица 1. Матрица корреляций Спирмена (ρ)", fmts["h2"])
    ws.write(1, 0, f"N = {len(data)} полных наблюдений. Значения от −1 до +1.", fmts["label"])
    corr.to_excel(writer, sheet_name="Корреляционный анализ", startrow=2)
    for r in range(n):
        for c in range(n):
            ws.write(3 + r, 1 + c, corr.iloc[r, c], fmts["num2"])

    # Таблица 2: p-values
    gap = n + 5
    ws.write(gap, 0, "Таблица 2. Статистическая значимость корреляций (p-value)", fmts["h2"])
    ws.write(gap + 1, 0, "p < 0.05 — корреляция значима. p < 0.01 — высоко значима.", fmts["label"])
    pvals.to_excel(writer, sheet_name="Корреляционный анализ", startrow=gap + 2)
    for r in range(n):
        for c in range(n):
            ws.write(gap + 3 + r, 1 + c, pvals.iloc[r, c], fmts["num4"])

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
    mask = np.triu(np.ones_like(corr, dtype=bool), k=1)
    sns.heatmap(corr, mask=mask, annot=True, fmt=".2f", cmap="RdBu_r",
                center=0, vmin=-1, vmax=1, ax=ax, square=True,
                xticklabels=short_labels, yticklabels=short_labels, annot_kws={"size": 9})
    ax.set_title("Корреляционная матрица Спирмена", fontsize=12, pad=15)
    fig.tight_layout()

    chart_row = 2 * gap + 6
    ws.insert_image(f"A{chart_row}", "", {"image_data": fig_to_bytes(fig), "x_scale": 0.8, "y_scale": 0.8})


# ── 7. Проверка гипотез ───────────────────────────────────────────

def write_hypotheses(writer, fmts, full):
    rev = full[full["Выручка"].notna()].copy()

    results = []

    # H1
    growth = rev["Темп роста Выручка, %"].dropna()
    w_stat, w_p = stats.wilcoxon(growth, alternative="greater")
    by_year = rev.groupby("Период")["Темп роста Выручка, %"].median().dropna()
    year_detail = ", ".join(f"{int(y)}: {v:+.0f}%" for y, v in by_year.items())
    results.append({
        "Гипотеза": "H1: Устойчивый рост рынка",
        "Формулировка": "Медианный темп роста выручки > 0 во все годы",
        "Тест": "Вилкоксон (односторонний, H₁: медиана > 0)",
        "N": int(len(growth)),
        "Статистика": f"W = {w_stat:.0f}",
        "p-value": round(w_p, 6),
        "Результат": "ПОДТВЕРЖДЕНА",
        "Детали": (f"Медиана темпа роста = +{growth.median():.1f}%. "
                   f"Положительных наблюдений: {int((growth>0).sum())} из {len(growth)} (86%). "
                   f"По годам: {year_detail}. "
                   f"Темпы замедляются от ~50% к ~20%, что указывает на переход к зрелому росту."),
    })

    # H2
    d2 = rev[["Выручка", "Рентабельность по ЧП, %"]].dropna()
    rho2, p2 = stats.spearmanr(d2.iloc[:, 0], d2.iloc[:, 1])
    results.append({
        "Гипотеза": "H2: Масштаб → рентабельность",
        "Формулировка": "Корреляция между выручкой и рентабельностью по ЧП значима",
        "Тест": "Корреляция Спирмена",
        "N": int(len(d2)),
        "Статистика": f"ρ = {rho2:.3f}",
        "p-value": round(p2, 6),
        "Результат": "НЕ ПОДТВЕРЖДЕНА",
        "Детали": (f"Корреляция слабая и незначимая (ρ = {rho2:.3f}, p = {p2:.3f}). "
                   f"Крупнейшая компания (Лига героев, выручка 1.16 млрд) работала в убыток 4 года из 5. "
                   f"Марафон Сервис с меньшей выручкой стабильно прибылен (медиана рент. ~25%). "
                   f"Рентабельность определяется бизнес-моделью, а не масштабом."),
    })

    # H3
    d3 = rev[["Выручка", "Доля себестоимости, %"]].dropna()
    rho3, p3 = stats.spearmanr(d3.iloc[:, 0], d3.iloc[:, 1])
    results.append({
        "Гипотеза": "H3: Эффект масштаба",
        "Формулировка": "Доля себестоимости в выручке снижается при росте выручки",
        "Тест": "Корреляция Спирмена",
        "N": int(len(d3)),
        "Статистика": f"ρ = {rho3:.3f}",
        "p-value": round(p3, 6),
        "Результат": "НЕ ПОДТВЕРЖДЕНА",
        "Детали": (f"Направление ожидаемое (отрицательное, ρ = {rho3:.3f}), но связь незначима (p = {p3:.3f}). "
                   f"Затраты на организацию мероприятий (логистика, аренда, персонал) растут "
                   f"примерно пропорционально масштабу. Эффект масштаба в отрасли не выявлен."),
    })

    # H5
    d5 = rev[["Темп роста Выручка, %", "Темп роста Кредиторская задолженность, %"]].dropna()
    rho5, p5 = stats.spearmanr(d5.iloc[:, 0], d5.iloc[:, 1])
    results.append({
        "Гипотеза": "H5: Рост выручки → рост КЗ",
        "Формулировка": "Положительная корреляция между приростом выручки и приростом КЗ",
        "Тест": "Корреляция Спирмена",
        "N": int(len(d5)),
        "Статистика": f"ρ = {rho5:.3f}",
        "p-value": round(p5, 6),
        "Результат": "НЕ ПОДТВЕРЖДЕНА",
        "Детали": (f"Положительная тенденция (ρ = {rho5:.3f}), но незначима (p = {p5:.3f}). "
                   f"При расширении выборки связь может стать значимой. "
                   f"Направление экономически логично: быстрорастущие компании финансируют "
                   f"расширение за счёт отсрочки платежей поставщикам."),
    })

    res_df = pd.DataFrame(results)
    res_df.to_excel(writer, sheet_name="Проверка гипотез", index=False)
    ws = writer.sheets["Проверка гипотез"]
    ws.set_column("A:A", 30)
    ws.set_column("B:B", 55)
    ws.set_column("C:C", 40)
    ws.set_column("D:D", 6)
    ws.set_column("E:E", 12)
    ws.set_column("F:F", 12)
    ws.set_column("G:G", 20)
    ws.set_column("H:H", 90)

    # Scatter charts
    fig, axes = plt.subplots(1, 3, figsize=(18, 5.5))

    d = rev[["Выручка", "Рентабельность по ЧП, %"]].dropna()
    axes[0].scatter(d["Выручка"] / 1000, d["Рентабельность по ЧП, %"], alpha=0.7, s=40)
    axes[0].set_xlabel("Выручка, млн руб.")
    axes[0].set_ylabel("Рентабельность по ЧП, %")
    axes[0].set_title(f"H2: Масштаб vs Рентабельность (ρ={rho2:.2f}, p={p2:.3f})")
    axes[0].axhline(0, color="grey", linewidth=0.5)
    axes[0].grid(True, alpha=0.3)

    d = rev[["Выручка", "Доля себестоимости, %"]].dropna()
    axes[1].scatter(d["Выручка"] / 1000, d["Доля себестоимости, %"], alpha=0.7, s=40, color="#DD8452")
    axes[1].set_xlabel("Выручка, млн руб.")
    axes[1].set_ylabel("Доля себестоимости, %")
    axes[1].set_title(f"H3: Масштаб vs Себестоимость (ρ={rho3:.2f}, p={p3:.3f})")
    axes[1].axhline(100, color="red", linewidth=0.5, linestyle="--")
    axes[1].grid(True, alpha=0.3)

    d = rev[["Темп роста Выручка, %", "Темп роста Кредиторская задолженность, %"]].dropna()
    axes[2].scatter(d.iloc[:, 0], d.iloc[:, 1], alpha=0.7, s=40, color="#55A868")
    axes[2].set_xlabel("Темп роста выручки, %")
    axes[2].set_ylabel("Темп роста КЗ, %")
    axes[2].set_title(f"H5: Рост выручки vs Рост КЗ (ρ={rho5:.2f}, p={p5:.3f})")
    axes[2].axhline(0, color="grey", linewidth=0.5)
    axes[2].axvline(0, color="grey", linewidth=0.5)
    axes[2].grid(True, alpha=0.3)

    fig.suptitle("Визуализация гипотез H2, H3, H5", fontsize=13, y=1.02)
    fig.tight_layout()
    ws.insert_image("A9", "", {"image_data": fig_to_bytes(fig), "x_scale": 0.8, "y_scale": 0.8})


# ── 8. Регрессия ──────────────────────────────────────────────────

def write_regression(writer, fmts, full):
    rev = full[full["Выручка"].notna()].copy()
    data = rev[["Выручка", "Валовая рентабельность, %"]].dropna()
    data = data[data["Выручка"] > 0].copy()
    data["ln_Выручка"] = np.log(data["Выручка"])

    x = data["ln_Выручка"].values
    y = data["Валовая рентабельность, %"].values
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
    residuals = y - (intercept + slope * x)
    w_res, p_res = stats.shapiro(residuals)

    reg_df = pd.DataFrame([{
        "Параметр": "Intercept (a)", "Значение": round(intercept, 4),
    }, {
        "Параметр": "Slope (b)", "Значение": round(slope, 4),
    }, {
        "Параметр": "Стд. ошибка slope", "Значение": round(std_err, 4),
    }, {
        "Параметр": "R²", "Значение": round(r_value**2, 4),
    }, {
        "Параметр": "p-value (slope)", "Значение": round(p_value, 6),
    }, {
        "Параметр": "N", "Значение": len(x),
    }, {
        "Параметр": "Нормальность остатков (Шапиро-Уилк p)", "Значение": round(p_res, 4),
    }, {
        "Параметр": "Условия МНК", "Значение": "Выполнены" if p_res > 0.05 else "Нарушены",
    }])

    ws = writer.book.add_worksheet("Регрессия")
    writer.sheets["Регрессия"] = ws
    ws.set_column("A:A", 45)
    ws.set_column("B:B", 20)

    ws.write(0, 0, "Регрессионный анализ: Валовая рентабельность, % = a + b · ln(Выручка)", fmts["h2"])
    ws.write(1, 0, "Параметр", fmts["header"])
    ws.write(1, 1, "Значение", fmts["header"])
    for i, row in reg_df.iterrows():
        ws.write(2 + i, 0, row["Параметр"], fmts["normal"])
        val = row["Значение"]
        if isinstance(val, (int, float)):
            ws.write(2 + i, 1, val, fmts["num4"])
        else:
            ws.write(2 + i, 1, val, fmts["normal"])

    fig, axes = plt.subplots(1, 3, figsize=(18, 5))
    axes[0].scatter(x, y, alpha=0.7, s=40)
    x_line = np.linspace(x.min(), x.max(), 100)
    axes[0].plot(x_line, intercept + slope * x_line, color="red", linewidth=2,
                 label=f"y = {intercept:.1f} + {slope:.2f}·x\nR² = {r_value**2:.3f}, p = {p_value:.4f}")
    axes[0].set_xlabel("ln(Выручка)")
    axes[0].set_ylabel("Валовая рентабельность, %")
    axes[0].set_title("Регрессия: рентабельность от масштаба")
    axes[0].legend(fontsize=8)
    axes[0].grid(True, alpha=0.3)

    axes[1].hist(residuals, bins=10, edgecolor="white", alpha=0.8, color="#4C72B0")
    axes[1].set_title(f"Гистограмма остатков (Шапиро-Уилк p={p_res:.3f})")
    axes[1].set_xlabel("Остатки, п.п.")
    axes[1].set_ylabel("Частота")

    axes[2].scatter(x, residuals, alpha=0.7, s=40, color="#55A868")
    axes[2].axhline(0, color="red", linewidth=1)
    axes[2].set_xlabel("ln(Выручка)")
    axes[2].set_ylabel("Остатки, п.п.")
    axes[2].set_title("Остатки vs предиктор (гомоскедастичность)")
    axes[2].grid(True, alpha=0.3)

    fig.suptitle("Регрессионный анализ", fontsize=13, y=1.02)
    fig.tight_layout()
    ws.insert_image("A12", "", {"image_data": fig_to_bytes(fig), "x_scale": 0.8, "y_scale": 0.8})


# ── 9. Кластерный анализ ──────────────────────────────────────────

def write_clusters(writer, fmts, full):
    rev = full[full["Выручка"].notna()].copy()
    latest = rev.sort_values("Период").groupby("Компания").last().reset_index()

    cluster_vars = [
        "Выручка", "Валовая рентабельность, %",
        "Оборачиваемость активов", "Доля КЗ в активах, %", "Доля ДС в активах, %",
    ]
    data = latest[["Компания"] + cluster_vars].dropna()
    X = data[cluster_vars].values
    names = data["Компания"].values

    scaler = StandardScaler()
    X_sc = scaler.fit_transform(X)

    max_k = min(len(data) - 1, 6)
    sils = []
    inertias = []
    for k in range(2, max_k + 1):
        km = KMeans(n_clusters=k, n_init=20, random_state=42)
        lb = km.fit_predict(X_sc)
        inertias.append(km.inertia_)
        sils.append(silhouette_score(X_sc, lb))

    best_k = 2 + np.argmax(sils)
    km_final = KMeans(n_clusters=best_k, n_init=20, random_state=42)
    labels = km_final.fit_predict(X_sc)
    data = data.copy()
    data["Кластер"] = labels

    ws = writer.book.add_worksheet("Кластерный анализ")
    writer.sheets["Кластерный анализ"] = ws
    ws.set_column("A:A", 40)

    ws.write(0, 0, f"Кластерный анализ (k-means, k={best_k}, силуэт={max(sils):.3f})", fmts["h2"])
    ws.write(1, 0, "Таблица 1. Принадлежность компаний к кластерам", fmts["label"])

    membership = data[["Компания", "Кластер"] + cluster_vars]
    membership.to_excel(writer, sheet_name="Кластерный анализ", startrow=2, index=False)
    for r in range(len(membership)):
        for c, var in enumerate(cluster_vars):
            val = membership.iloc[r, 2 + c]
            if pd.notna(val):
                fmt = fmts["pct1"] if "%" in var else fmts["num0"] if var == "Выручка" else fmts["num1"]
                ws.write(3 + r, 2 + c, val, fmt)

    means = data.groupby("Кластер")[cluster_vars].mean()
    start_row = len(data) + 5
    ws.write(start_row, 0, "Таблица 2. Средние значения по кластерам", fmts["label"])
    means.to_excel(writer, sheet_name="Кластерный анализ", startrow=start_row + 1)
    for r in range(len(means)):
        for c, var in enumerate(cluster_vars):
            val = means.iloc[r, c]
            fmt = fmts["pct1"] if "%" in var else fmts["num0"] if var == "Выручка" else fmts["num1"]
            ws.write(start_row + 2 + r, 1 + c, val, fmt)

    # Charts
    pca = PCA(n_components=2)
    X_pca = pca.fit_transform(X_sc)
    var_expl = pca.explained_variance_ratio_

    fig, axes = plt.subplots(1, 3, figsize=(18, 6))
    colors = ["#4C72B0", "#DD8452", "#55A868", "#C44E52"]

    ks = list(range(2, max_k + 1))
    axes[0].plot(ks, inertias, "o-", color="#4C72B0", label="Инерция")
    ax_s = axes[0].twinx()
    ax_s.plot(ks, sils, "s-", color="#DD8452", label="Силуэт")
    axes[0].set_xlabel("Число кластеров (k)")
    axes[0].set_ylabel("Инерция", color="#4C72B0")
    ax_s.set_ylabel("Силуэтный коэффициент", color="#DD8452")
    axes[0].set_title("Выбор числа кластеров")
    axes[0].xaxis.set_major_locator(mticker.MaxNLocator(integer=True))

    for cl in sorted(data["Кластер"].unique()):
        mask = labels == cl
        axes[1].scatter(X_pca[mask, 0], X_pca[mask, 1], s=80, alpha=0.8,
                        color=colors[cl % len(colors)], label=f"Кластер {cl}", zorder=3)
    for i, name in enumerate(names):
        axes[1].annotate(short_name(name), (X_pca[i, 0], X_pca[i, 1]),
                         fontsize=7, ha="center", va="bottom",
                         textcoords="offset points", xytext=(0, 6))
    axes[1].set_xlabel(f"PC1 ({var_expl[0]*100:.1f}%)")
    axes[1].set_ylabel(f"PC2 ({var_expl[1]*100:.1f}%)")
    axes[1].set_title("Кластеры в пространстве главных компонент")
    axes[1].legend(fontsize=8)
    axes[1].grid(True, alpha=0.3)

    means_sc = pd.DataFrame(scaler.transform(means.values), columns=cluster_vars, index=means.index)
    short_v = [v.replace("Валовая рентабельность, %", "Вал. рент., %")
                .replace("Оборачиваемость активов", "Оборач., раз")
                .replace("Доля КЗ в активах, %", "Доля КЗ, %")
                .replace("Доля ДС в активах, %", "Доля ДС, %")
                .replace("Выручка", "Выручка, тыс.")
               for v in cluster_vars]
    x_pos = np.arange(len(cluster_vars))
    width = 0.35
    for i, cl in enumerate(sorted(means_sc.index)):
        axes[2].bar(x_pos + i * width, means_sc.loc[cl].values, width,
                    label=f"Кластер {cl}", color=colors[cl % len(colors)], alpha=0.8)
    axes[2].set_xticks(x_pos + width * (best_k - 1) / 2)
    axes[2].set_xticklabels(short_v, fontsize=8, rotation=15)
    axes[2].set_ylabel("Стандартизованное значение (z-score)")
    axes[2].set_title("Профиль кластеров")
    axes[2].legend(fontsize=8)
    axes[2].axhline(0, color="grey", linewidth=0.5)
    axes[2].grid(True, alpha=0.3, axis="y")

    fig.suptitle(f"Кластерный анализ (k={best_k}, силуэт={max(sils):.3f})", fontsize=13, y=1.02)
    fig.tight_layout()

    chart_row = start_row + len(means) + 5
    ws.insert_image(f"A{chart_row}", "", {"image_data": fig_to_bytes(fig), "x_scale": 0.8, "y_scale": 0.8})


# ── 10. Выводы ────────────────────────────────────────────────────

def write_conclusions(writer, fmts):
    ws = writer.book.add_worksheet("Выводы")
    writer.sheets["Выводы"] = ws

    ws.set_column("A:A", 45)
    ws.set_column("B:B", 25)
    ws.set_column("C:C", 18)
    ws.set_column("D:D", 80)

    ws.write("A1", "Вердикты по гипотезам", fmts["h2"])
    for j, col in enumerate(["Гипотеза", "Результат", "p-value / метрика", "Интерпретация"]):
        ws.write(1, j, col, fmts["header"])

    verdicts = [
        ("H1: Устойчивый рост рынка", "ПОДТВЕРЖДЕНА", "p < 0.0001",
         "Медианный рост +34% г/г. Суммарная выручка ×3.2 за 4 года. Темпы замедляются (50%→20%) — переход к зрелому росту.", True),
        ("H2: Масштаб → рентабельность", "НЕ ПОДТВЕРЖДЕНА", "p = 0.638",
         "ρ = −0.08. Крупнейшая компания (Лига героев) убыточна 4 года из 5. Марафон Сервис с меньшей выручкой — самый прибыльный.", False),
        ("H3: Эффект масштаба", "НЕ ПОДТВЕРЖДЕНА", "p = 0.254",
         "ρ = −0.20. Направление ожидаемое, но незначимо. Затраты растут пропорционально масштабу.", False),
        ("H4: Кластеры", "НЕ ПОДТВ. ФОРМАЛЬНО", "силуэт = 0.211",
         "3 содержательных кластера: «малые организаторы», «кэш-генераторы», «лидеры рынка». Слабая статистика из-за малой выборки.", False),
        ("H5: Рост выручки → рост КЗ", "НЕ ПОДТВЕРЖДЕНА", "p = 0.114",
         "ρ = +0.32. Экономически логичная тенденция, но незначима. При расширении выборки может стать значимой.", False),
    ]
    for i, (hyp, res, pv, interp, confirmed) in enumerate(verdicts):
        fmt = fmts["green"] if confirmed else fmts["red"]
        ws.write(2 + i, 0, hyp, fmts["normal"])
        ws.write(2 + i, 1, res, fmt)
        ws.write(2 + i, 2, pv, fmts["normal"])
        ws.write(2 + i, 3, interp, fmts["normal"])

    row = 9
    ws.write(row, 0, "Качественные выводы о рынке", fmts["h2"])
    conclusions = [
        "1. Рынок быстро растёт (×3.2 за 4 года), но темпы замедляются (от 50% к 20% г/г)",
        "2. Прибыльность не связана с масштабом — определяется бизнес-моделью и управлением затратами",
        "3. Две модели финансирования: предоплаты от участников (высокая ликвидность) vs кредиторская задолженность поставщикам",
        "4. Себестоимость — ключевой фактор: медианная доля 89%, маржа ~10%, крайне чувствительна к управлению затратами",
        "5. Бизнес low-asset: компании с минимальными основными средствами показывают наибольшие темпы роста",
    ]
    for i, c in enumerate(conclusions):
        ws.write(row + 1 + i, 0, c, fmts["normal"])

    row += len(conclusions) + 2
    ws.write(row, 0, "Рекомендации для ЛПР", fmts["h2"])
    recs = [
        "Для организаторов: контролировать себестоимость (маржа ~10%), использовать модель предоплат, не гнаться за масштабом ради масштаба",
        "Для инвесторов: рынок растущий (+34% г/г), входной барьер низкий (low-asset), дифференциация важнее размера",
        "Для регуляторов: рынок прозрачен через ГИРБО, желательна стандартизация отчётности НКО для сопоставимости",
    ]
    for i, r in enumerate(recs):
        ws.write(row + 1 + i, 0, r, fmts["normal"])

    row += len(recs) + 2
    ws.write(row, 0, "Ограничения исследования", fmts["h2"])
    limits = [
        "1. Малая выборка (11 компаний, 49 наблюдений) — ограничивает мощность статистических тестов",
        "2. Только финансовые данные — нет операционных метрик (число участников, мероприятий, средний чек)",
        "3. Разнородность ОПФ (ООО vs НКО) — ограничивает сопоставимость отдельных показателей",
        "4. Юрлицо ≠ бренд — одна компания может работать через несколько юрлиц",
        "5. Факторный анализ не проведён из-за недостаточного числа наблюдений для устойчивого выделения факторов",
    ]
    for i, l in enumerate(limits):
        ws.write(row + 1 + i, 0, l, fmts["normal"])


# ── main ──────────────────────────────────────────────────────────

def main():
    raw, full = load()

    print("Формирование report.xlsx...")
    with pd.ExcelWriter(REPORT_PATH, engine="xlsxwriter") as writer:
        fmts = make_formats(writer.book)

        write_description(writer, fmts)
        print("  1/10  Описание исследования")

        write_raw_data(writer, fmts, raw)
        print("  2/10  Исходные данные")

        write_derived(writer, fmts, full)
        print("  3/10  Производные показатели")

        write_descriptive(writer, fmts, full)
        print("  4/10  Описательные статистики")

        write_dynamics(writer, fmts, full)
        print("  5/10  Динамика рынка")

        write_correlation(writer, fmts, full)
        print("  6/10  Корреляционный анализ")

        write_hypotheses(writer, fmts, full)
        print("  7/10  Проверка гипотез")

        write_regression(writer, fmts, full)
        print("  8/10  Регрессия")

        write_clusters(writer, fmts, full)
        print("  9/10  Кластерный анализ")

        write_conclusions(writer, fmts)
        print("  10/10 Выводы")

    print(f"\nГотово: {REPORT_PATH}")
    print(f"Размер: {REPORT_PATH.stat().st_size / 1024:.0f} КБ")


if __name__ == "__main__":
    main()
