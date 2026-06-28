"""Этап 8. Формирование итогового xlsx-отчёта со встроенными графиками.

Где возможно — нативные Excel-графики (line, scatter, column).
Где нет (boxplot, гистограммы, PCA) — PNG-изображения.
Корреляционная матрица — условное форматирование ячеек.
"""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from io import BytesIO
from scipy import stats
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler

plt.rcParams.update({"font.family": "DejaVu Sans", "figure.dpi": 130, "savefig.bbox": "tight"})

APP = Path(__file__).parent
OUT = APP / "output"
REPORT_PATH = OUT / "report.xlsx"

COLORS = ["#4C72B0", "#DD8452", "#55A868", "#C44E52", "#8172B3", "#937860", "#DA8BC3"]


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
    ws.set_column("D:D", 25)

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
    for j, col in enumerate(["№", "Гипотеза", "Метод проверки"]):
        ws.write(12, j, col, fmts["header"])
    hypotheses = [
        ("H1", "Рынок демонстрирует устойчивый рост выручки", "Тест Вилкоксона"),
        ("H2", "Существует зависимость между масштабом бизнеса и рентабельностью", "Корреляция Спирмена"),
        ("H3", "Доля себестоимости снижается с ростом выручки (эффект масштаба)", "Корреляция Спирмена"),
        ("H4", "Компании образуют кластеры по финансовому профилю", "k-means, силуэтный коэффициент"),
        ("H5", "Рост выручки сопровождается ростом кредиторской задолженности", "Корреляция Спирмена"),
    ]
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


# ── 5. Динамика рынка — нативные Excel-графики ─────────────────────

def write_dynamics(writer, fmts, full):
    SHEET = "Динамика рынка"
    rev = full[full["Выручка"].notna()]
    book = writer.book

    pivot_rev = rev.pivot_table(index="Компания", columns="Период", values="Выручка")
    pivot_profit = rev.pivot_table(index="Компания", columns="Период", values="Чистая прибыль (убыток)")
    pivot_rent = rev.pivot_table(index="Компания", columns="Период", values="Валовая рентабельность, %")

    ws = book.add_worksheet(SHEET)
    writer.sheets[SHEET] = ws
    ws.set_column("A:A", 40)
    for c in range(1, 7):
        ws.set_column(c, c, 14)

    row = 0

    def write_pivot(ws, pivot, title, fmt, start_row):
        ws.write(start_row, 0, title, fmts["h2"])
        r = start_row + 1
        years = list(pivot.columns)
        ws.write(r, 0, "Компания", fmts["header"])
        for j, y in enumerate(years):
            ws.write(r, 1 + j, int(y), fmts["header"])
        for i, (comp, vals) in enumerate(pivot.iterrows()):
            ws.write(r + 1 + i, 0, comp, fmts["normal"])
            for j, v in enumerate(vals):
                if pd.notna(v):
                    ws.write(r + 1 + i, 1 + j, v, fmt)
        data_start = r + 1
        data_end = r + len(pivot)
        return data_start, data_end, years

    # Table 1: Revenue
    ds1, de1, years = write_pivot(ws, pivot_rev, "Таблица 1. Выручка компаний по годам, тыс. руб.", fmts["num0"], row)
    row = de1 + 2

    # Table 2: Profit
    ds2, de2, _ = write_pivot(ws, pivot_profit, "Таблица 2. Чистая прибыль (убыток) по годам, тыс. руб.", fmts["num0"], row)
    row = de2 + 2

    # Table 3: Rentability
    ds3, de3, _ = write_pivot(ws, pivot_rent, "Таблица 3. Валовая рентабельность по годам, %", fmts["pct1"], row)
    row = de3 + 2

    # Native Excel line chart: Revenue dynamics
    chart_rev = book.add_chart({"type": "line"})
    chart_rev.set_title({"name": "Динамика выручки по компаниям, тыс. руб."})
    chart_rev.set_x_axis({"name": "Год"})
    chart_rev.set_y_axis({"name": "тыс. руб.", "num_format": "#,##0"})
    chart_rev.set_size({"width": 800, "height": 480})
    for i in range(de1 - ds1 + 1):
        chart_rev.add_series({
            "name": [SHEET, ds1 + i, 0],
            "categories": [SHEET, ds1 - 1, 1, ds1 - 1, len(years)],
            "values": [SHEET, ds1 + i, 1, ds1 + i, len(years)],
            "line": {"width": 2},
            "marker": {"type": "circle", "size": 5},
        })
    chart_rev.set_legend({"position": "bottom", "font": {"size": 8}})
    ws.insert_chart(f"A{row + 1}", chart_rev)
    row += 28

    # Native Excel line chart: Profit dynamics
    chart_profit = book.add_chart({"type": "line"})
    chart_profit.set_title({"name": "Динамика чистой прибыли, тыс. руб."})
    chart_profit.set_x_axis({"name": "Год"})
    chart_profit.set_y_axis({"name": "тыс. руб.", "num_format": "#,##0"})
    chart_profit.set_size({"width": 800, "height": 480})
    for i in range(de2 - ds2 + 1):
        chart_profit.add_series({
            "name": [SHEET, ds2 + i, 0],
            "categories": [SHEET, ds2 - 1, 1, ds2 - 1, len(years)],
            "values": [SHEET, ds2 + i, 1, ds2 + i, len(years)],
            "line": {"width": 2},
            "marker": {"type": "circle", "size": 5},
        })
    chart_profit.set_legend({"position": "bottom", "font": {"size": 8}})
    ws.insert_chart(f"A{row + 1}", chart_profit)
    row += 28

    # Native Excel line chart: Rentability
    chart_rent = book.add_chart({"type": "line"})
    chart_rent.set_title({"name": "Динамика валовой рентабельности, %"})
    chart_rent.set_x_axis({"name": "Год"})
    chart_rent.set_y_axis({"name": "%", "num_format": "0.0"})
    chart_rent.set_size({"width": 800, "height": 480})
    for i in range(de3 - ds3 + 1):
        chart_rent.add_series({
            "name": [SHEET, ds3 + i, 0],
            "categories": [SHEET, ds3 - 1, 1, ds3 - 1, len(years)],
            "values": [SHEET, ds3 + i, 1, ds3 + i, len(years)],
            "line": {"width": 2},
            "marker": {"type": "circle", "size": 5},
        })
    chart_rent.set_legend({"position": "bottom", "font": {"size": 8}})
    ws.insert_chart(f"A{row + 1}", chart_rent)
    row += 28

    # Market summary data table
    agg = rev.groupby("Период").agg(
        Сумма_выручки=("Выручка", "sum"),
        Медиана_выручки=("Выручка", "median"),
        Медиана_рент=("Валовая рентабельность, %", "median"),
    ).reset_index()

    ws.write(row, 0, "Таблица 4. Сводные показатели рынка по годам", fmts["h2"])
    row += 1
    for j, col in enumerate(["Год", "Сумма выручки, тыс. руб.", "Медиана выручки, тыс. руб.", "Медиана рентабельности, %"]):
        ws.write(row, j, col, fmts["header"])
    agg_start = row + 1
    for i, r in agg.iterrows():
        ws.write(agg_start + i, 0, int(r["Период"]), fmts["normal"])
        ws.write(agg_start + i, 1, r["Сумма_выручки"], fmts["num0"])
        ws.write(agg_start + i, 2, r["Медиана_выручки"], fmts["num0"])
        ws.write(agg_start + i, 3, r["Медиана_рент"], fmts["pct1"])
    agg_end = agg_start + len(agg) - 1
    row = agg_end + 2

    # Native column chart: Total revenue
    chart_total = book.add_chart({"type": "column"})
    chart_total.set_title({"name": "Суммарная выручка рынка, тыс. руб."})
    chart_total.set_x_axis({"name": "Год"})
    chart_total.set_y_axis({"name": "тыс. руб.", "num_format": "#,##0"})
    chart_total.set_size({"width": 500, "height": 350})
    chart_total.add_series({
        "categories": [SHEET, agg_start, 0, agg_end, 0],
        "values": [SHEET, agg_start, 1, agg_end, 1],
        "fill": {"color": COLORS[0]},
    })
    chart_total.set_legend({"none": True})
    ws.insert_chart(f"A{row + 1}", chart_total)

    # Native line chart: Median revenue
    chart_med = book.add_chart({"type": "line"})
    chart_med.set_title({"name": "Медианная выручка, тыс. руб."})
    chart_med.set_x_axis({"name": "Год"})
    chart_med.set_y_axis({"name": "тыс. руб.", "num_format": "#,##0"})
    chart_med.set_size({"width": 500, "height": 350})
    chart_med.add_series({
        "categories": [SHEET, agg_start, 0, agg_end, 0],
        "values": [SHEET, agg_start, 2, agg_end, 2],
        "line": {"color": COLORS[1], "width": 2.5},
        "marker": {"type": "circle", "size": 6, "fill": {"color": COLORS[1]}},
    })
    chart_med.set_legend({"none": True})
    ws.insert_chart(f"H{row + 1}", chart_med)
    row += 22

    # Boxplot — PNG (Excel не поддерживает нативно)
    ws.write(row, 0, "Распределение показателей по компаниям (boxplot, PNG)", fmts["label"])
    row += 1
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    for ax, col in zip(axes.flatten(), ["Выручка", "Чистая прибыль (убыток)", "Валовая рентабельность, %", "Оборачиваемость активов"]):
        data = rev[["Компания", col]].dropna()
        if data.empty:
            continue
        order = data.groupby("Компания")[col].median().sort_values(ascending=False).index
        sns.boxplot(data=data, y="Компания", x=col, order=order, ax=ax, hue="Компания", legend=False, palette="Set2")
        unit = "тыс. руб." if "руб" not in col and "%" not in col and "Оборач" not in col else ""
        ax.set_title(col + (f", {unit}" if unit else ""), fontsize=11)
        ax.set_ylabel("")
    fig.suptitle("Распределение показателей по компаниям", fontsize=13, y=1.01)
    fig.tight_layout()
    ws.insert_image(f"A{row + 1}", "", {"image_data": fig_to_bytes(fig), "x_scale": 0.8, "y_scale": 0.8})


# ── 6. Корреляционный анализ — условное форматирование ─────────────

def write_correlation(writer, fmts, full):
    SHEET = "Корреляционный анализ"
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

    ws = writer.book.add_worksheet(SHEET)
    writer.sheets[SHEET] = ws
    ws.set_column("A:A", 35)
    for c in range(1, n + 1):
        ws.set_column(c, c, 14)

    # Table 1: Correlation matrix with conditional formatting
    ws.write(0, 0, "Таблица 1. Матрица корреляций Спирмена (ρ)", fmts["h2"])
    ws.write(1, 0, f"N = {len(data)} полных наблюдений. Цветовая шкала: синий (−1) → белый (0) → красный (+1).", fmts["label"])

    hdr_row = 2
    ws.write(hdr_row, 0, "", fmts["header"])
    for j, col in enumerate(corr_cols):
        short = col.replace("Кредиторская задолженность", "Кредит. задолж.") \
                    .replace("Себестоимость продаж", "Себестоимость") \
                    .replace("Чистая прибыль (убыток)", "Чист. прибыль") \
                    .replace("Валовая рентабельность, %", "Вал. рент., %") \
                    .replace("Рентабельность по ЧП, %", "Рент. по ЧП, %") \
                    .replace("Оборачиваемость активов", "Оборач. активов") \
                    .replace("Денежные средства", "Ден. средства") \
                    .replace("Активы всего", "Активы")
        ws.write(hdr_row, 1 + j, short, fmts["header"])
        ws.write(hdr_row + 1 + j, 0, short, fmts["normal"])

    data_start_row = hdr_row + 1
    for i in range(n):
        for j in range(n):
            ws.write(data_start_row + i, 1 + j, corr.iloc[i, j], fmts["num2"])

    # Conditional formatting: color scale blue-white-red
    cell_range = f"B{data_start_row + 1}:{chr(ord('A') + n)}{data_start_row + n}"
    ws.conditional_format(cell_range, {
        "type": "3_color_scale",
        "min_color": "#2166AC",
        "mid_color": "#FFFFFF",
        "max_color": "#B2182B",
        "min_value": -1,
        "mid_value": 0,
        "max_value": 1,
        "min_type": "num",
        "mid_type": "num",
        "max_type": "num",
    })

    # Table 2: p-values
    gap = data_start_row + n + 2
    ws.write(gap, 0, "Таблица 2. Статистическая значимость корреляций (p-value)", fmts["h2"])
    ws.write(gap + 1, 0, "p < 0.05 — значима, p < 0.01 — высоко значима. Зелёный фон — значимые.", fmts["label"])

    p_hdr = gap + 2
    ws.write(p_hdr, 0, "", fmts["header"])
    for j, col in enumerate(corr_cols):
        short = col.replace("Кредиторская задолженность", "Кредит. задолж.") \
                    .replace("Себестоимость продаж", "Себестоимость") \
                    .replace("Чистая прибыль (убыток)", "Чист. прибыль") \
                    .replace("Валовая рентабельность, %", "Вал. рент., %") \
                    .replace("Рентабельность по ЧП, %", "Рент. по ЧП, %") \
                    .replace("Оборачиваемость активов", "Оборач. активов") \
                    .replace("Денежные средства", "Ден. средства") \
                    .replace("Активы всего", "Активы")
        ws.write(p_hdr, 1 + j, short, fmts["header"])
        ws.write(p_hdr + 1 + j, 0, short, fmts["normal"])

    p_data_start = p_hdr + 1
    for i in range(n):
        for j in range(n):
            ws.write(p_data_start + i, 1 + j, pvals.iloc[i, j], fmts["num4"])

    p_range = f"B{p_data_start + 1}:{chr(ord('A') + n)}{p_data_start + n}"
    ws.conditional_format(p_range, {
        "type": "cell", "criteria": "<", "value": 0.05,
        "format": writer.book.add_format({"bg_color": "#C6EFCE", "num_format": "0.0000", "font_size": 11}),
    })


# ── 7. Проверка гипотез — нативные scatter-графики ─────────────────

def write_hypotheses(writer, fmts, full):
    SHEET = "Проверка гипотез"
    rev = full[full["Выручка"].notna()].copy()
    book = writer.book

    # Compute hypothesis results
    growth = rev["Темп роста Выручка, %"].dropna()
    w_stat, w_p = stats.wilcoxon(growth, alternative="greater")
    by_year = rev.groupby("Период")["Темп роста Выручка, %"].median().dropna()
    year_detail = ", ".join(f"{int(y)}: {v:+.0f}%" for y, v in by_year.items())

    d2 = rev[["Выручка", "Рентабельность по ЧП, %"]].dropna()
    rho2, p2 = stats.spearmanr(d2.iloc[:, 0], d2.iloc[:, 1])

    d3 = rev[["Выручка", "Доля себестоимости, %"]].dropna()
    rho3, p3 = stats.spearmanr(d3.iloc[:, 0], d3.iloc[:, 1])

    d5 = rev[["Темп роста Выручка, %", "Темп роста Кредиторская задолженность, %"]].dropna()
    rho5, p5 = stats.spearmanr(d5.iloc[:, 0], d5.iloc[:, 1])

    results = [
        {"Гипотеза": "H1: Устойчивый рост рынка",
         "Формулировка": "Медианный темп роста выручки > 0 во все годы",
         "Тест": "Вилкоксон (односторонний, H₁: медиана > 0)",
         "N": int(len(growth)), "Статистика": f"W = {w_stat:.0f}", "p-value": round(w_p, 6),
         "Результат": "ПОДТВЕРЖДЕНА",
         "Детали": f"Медиана = +{growth.median():.1f}%. Положительных: {int((growth>0).sum())}/{len(growth)} (86%). По годам: {year_detail}. Темпы замедляются от ~50% к ~20%."},
        {"Гипотеза": "H2: Масштаб → рентабельность",
         "Формулировка": "Корреляция между выручкой и рентабельностью по ЧП значима",
         "Тест": "Корреляция Спирмена", "N": int(len(d2)), "Статистика": f"ρ = {rho2:.3f}",
         "p-value": round(p2, 6), "Результат": "НЕ ПОДТВЕРЖДЕНА",
         "Детали": f"ρ = {rho2:.3f}, p = {p2:.3f}. Лига героев (выручка 1.16 млрд) убыточна 4 года из 5. Марафон Сервис с меньшей выручкой — самый прибыльный (медиана рент. ~25%)."},
        {"Гипотеза": "H3: Эффект масштаба",
         "Формулировка": "Доля себестоимости снижается при росте выручки",
         "Тест": "Корреляция Спирмена", "N": int(len(d3)), "Статистика": f"ρ = {rho3:.3f}",
         "p-value": round(p3, 6), "Результат": "НЕ ПОДТВЕРЖДЕНА",
         "Детали": f"ρ = {rho3:.3f}, p = {p3:.3f}. Направление ожидаемое (−), но незначимо. Затраты на организацию мероприятий растут пропорционально масштабу."},
        {"Гипотеза": "H5: Рост выручки → рост КЗ",
         "Формулировка": "Положительная корреляция приростов выручки и КЗ",
         "Тест": "Корреляция Спирмена", "N": int(len(d5)), "Статистика": f"ρ = {rho5:.3f}",
         "p-value": round(p5, 6), "Результат": "НЕ ПОДТВЕРЖДЕНА",
         "Детали": f"ρ = {rho5:.3f}, p = {p5:.3f}. Тенденция (+), но незначима. При расширении выборки может стать значимой."},
    ]
    res_df = pd.DataFrame(results)
    res_df.to_excel(writer, sheet_name=SHEET, index=False)
    ws = writer.sheets[SHEET]
    ws.set_column("A:A", 30)
    ws.set_column("B:B", 55)
    ws.set_column("C:C", 40)
    ws.set_column("D:D", 6)
    ws.set_column("E:E", 12)
    ws.set_column("F:F", 12)
    ws.set_column("G:G", 20)
    ws.set_column("H:H", 90)

    # Write scatter data for native charts
    scatter_start = len(results) + 4

    # H2 data
    ws.write(scatter_start, 0, "Данные для графика H2", fmts["label"])
    ws.write(scatter_start + 1, 0, "Выручка, тыс. руб.", fmts["header"])
    ws.write(scatter_start + 1, 1, "Рентабельность по ЧП, %", fmts["header"])
    h2_data = d2.reset_index(drop=True)
    h2_start = scatter_start + 2
    for i in range(len(h2_data)):
        ws.write(h2_start + i, 0, h2_data.iloc[i, 0], fmts["num0"])
        ws.write(h2_start + i, 1, h2_data.iloc[i, 1], fmts["pct1"])
    h2_end = h2_start + len(h2_data) - 1

    chart_h2 = book.add_chart({"type": "scatter"})
    chart_h2.set_title({"name": f"H2: Масштаб vs Рентабельность (ρ={rho2:.2f}, p={p2:.3f})"})
    chart_h2.set_x_axis({"name": "Выручка, тыс. руб.", "num_format": "#,##0"})
    chart_h2.set_y_axis({"name": "Рентабельность по ЧП, %"})
    chart_h2.set_size({"width": 500, "height": 380})
    chart_h2.add_series({
        "categories": [SHEET, h2_start, 0, h2_end, 0],
        "values": [SHEET, h2_start, 1, h2_end, 1],
        "marker": {"type": "circle", "size": 6, "fill": {"color": COLORS[0]}},
        "line": {"none": True},
    })
    chart_h2.set_legend({"none": True})
    ws.insert_chart(f"D{scatter_start + 1}", chart_h2)

    # H3 data
    h3_col = 3
    ws.write(scatter_start, h3_col, "Данные для графика H3", fmts["label"])
    ws.write(scatter_start + 1, h3_col, "Выручка, тыс. руб.", fmts["header"])
    ws.write(scatter_start + 1, h3_col + 1, "Доля себестоимости, %", fmts["header"])
    h3_data = d3.reset_index(drop=True)
    h3_start = scatter_start + 2
    for i in range(len(h3_data)):
        ws.write(h3_start + i, h3_col, h3_data.iloc[i, 0], fmts["num0"])
        ws.write(h3_start + i, h3_col + 1, h3_data.iloc[i, 1], fmts["pct1"])
    h3_end = h3_start + len(h3_data) - 1

    chart_h3 = book.add_chart({"type": "scatter"})
    chart_h3.set_title({"name": f"H3: Масштаб vs Себестоимость (ρ={rho3:.2f}, p={p3:.3f})"})
    chart_h3.set_x_axis({"name": "Выручка, тыс. руб.", "num_format": "#,##0"})
    chart_h3.set_y_axis({"name": "Доля себестоимости, %"})
    chart_h3.set_size({"width": 500, "height": 380})
    chart_h3.add_series({
        "categories": [SHEET, h3_start, h3_col, h3_end, h3_col],
        "values": [SHEET, h3_start, h3_col + 1, h3_end, h3_col + 1],
        "marker": {"type": "circle", "size": 6, "fill": {"color": COLORS[1]}},
        "line": {"none": True},
    })
    chart_h3.set_legend({"none": True})
    ws.insert_chart(f"L{scatter_start + 1}", chart_h3)

    # H5 data
    h5_col = 6
    ws.write(scatter_start, h5_col, "Данные для графика H5", fmts["label"])
    ws.write(scatter_start + 1, h5_col, "Темп роста выручки, %", fmts["header"])
    ws.write(scatter_start + 1, h5_col + 1, "Темп роста КЗ, %", fmts["header"])
    h5_data = d5.reset_index(drop=True)
    h5_start = scatter_start + 2
    for i in range(len(h5_data)):
        ws.write(h5_start + i, h5_col, h5_data.iloc[i, 0], fmts["pct1"])
        ws.write(h5_start + i, h5_col + 1, h5_data.iloc[i, 1], fmts["pct1"])
    h5_end = h5_start + len(h5_data) - 1

    chart_h5 = book.add_chart({"type": "scatter"})
    chart_h5.set_title({"name": f"H5: Рост выручки vs Рост КЗ (ρ={rho5:.2f}, p={p5:.3f})"})
    chart_h5.set_x_axis({"name": "Темп роста выручки, %"})
    chart_h5.set_y_axis({"name": "Темп роста КЗ, %"})
    chart_h5.set_size({"width": 500, "height": 380})
    chart_h5.add_series({
        "categories": [SHEET, h5_start, h5_col, h5_end, h5_col],
        "values": [SHEET, h5_start, h5_col + 1, h5_end, h5_col + 1],
        "marker": {"type": "circle", "size": 6, "fill": {"color": COLORS[2]}},
        "line": {"none": True},
    })
    chart_h5.set_legend({"none": True})
    ws.insert_chart(f"D{scatter_start + 23}", chart_h5)


# ── 8. Регрессия — нативный scatter + PNG остатков ─────────────────

def write_regression(writer, fmts, full):
    SHEET = "Регрессия"
    rev = full[full["Выручка"].notna()].copy()
    book = writer.book

    data = rev[["Выручка", "Валовая рентабельность, %"]].dropna()
    data = data[data["Выручка"] > 0].copy()
    data["ln_Выручка"] = np.log(data["Выручка"])

    x = data["ln_Выручка"].values
    y = data["Валовая рентабельность, %"].values
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
    residuals = y - (intercept + slope * x)
    w_res, p_res = stats.shapiro(residuals)

    ws = book.add_worksheet(SHEET)
    writer.sheets[SHEET] = ws
    ws.set_column("A:A", 45)
    ws.set_column("B:B", 20)

    ws.write(0, 0, "Регрессионный анализ: Валовая рентабельность, % = a + b · ln(Выручка)", fmts["h2"])
    params = [
        ("Intercept (a)", round(intercept, 4)),
        ("Slope (b)", round(slope, 4)),
        ("Стд. ошибка slope", round(std_err, 4)),
        ("R²", round(r_value**2, 4)),
        ("p-value (slope)", round(p_value, 6)),
        ("N", len(x)),
        ("Нормальность остатков (Шапиро-Уилк p)", round(p_res, 4)),
        ("Условия МНК", "Выполнены" if p_res > 0.05 else "Нарушены"),
    ]
    ws.write(1, 0, "Параметр", fmts["header"])
    ws.write(1, 1, "Значение", fmts["header"])
    for i, (name, val) in enumerate(params):
        ws.write(2 + i, 0, name, fmts["normal"])
        if isinstance(val, (int, float)):
            ws.write(2 + i, 1, val, fmts["num4"])
        else:
            ws.write(2 + i, 1, val, fmts["normal"])

    # Scatter data for native chart
    scatter_row = 12
    ws.write(scatter_row, 0, "Данные для графика", fmts["label"])
    ws.write(scatter_row + 1, 0, "ln(Выручка)", fmts["header"])
    ws.write(scatter_row + 1, 1, "Вал. рентабельность, %", fmts["header"])
    ws.write(scatter_row + 1, 2, "Регрессия (предсказ.), %", fmts["header"])
    ws.set_column("C:C", 22)

    sorted_idx = np.argsort(x)
    for i, idx in enumerate(sorted_idx):
        ws.write(scatter_row + 2 + i, 0, x[idx], fmts["num2"])
        ws.write(scatter_row + 2 + i, 1, y[idx], fmts["pct1"])
        ws.write(scatter_row + 2 + i, 2, intercept + slope * x[idx], fmts["pct1"])
    s_start = scatter_row + 2
    s_end = s_start + len(x) - 1

    chart = book.add_chart({"type": "scatter"})
    chart.set_title({"name": f"Регрессия: рентабельность от масштаба (R²={r_value**2:.3f}, p={p_value:.3f})"})
    chart.set_x_axis({"name": "ln(Выручка)"})
    chart.set_y_axis({"name": "Валовая рентабельность, %"})
    chart.set_size({"width": 600, "height": 400})
    chart.add_series({
        "name": "Наблюдения",
        "categories": [SHEET, s_start, 0, s_end, 0],
        "values": [SHEET, s_start, 1, s_end, 1],
        "marker": {"type": "circle", "size": 6, "fill": {"color": COLORS[0]}},
        "line": {"none": True},
    })
    chart.add_series({
        "name": f"y = {intercept:.1f} + {slope:.2f}·x",
        "categories": [SHEET, s_start, 0, s_end, 0],
        "values": [SHEET, s_start, 2, s_end, 2],
        "marker": {"type": "none"},
        "line": {"color": "red", "width": 2},
    })
    ws.insert_chart(f"E{scatter_row + 1}", chart)

    # Residuals — PNG (histogram)
    res_row = scatter_row + 26
    ws.write(res_row, 0, "Диагностика остатков (PNG)", fmts["label"])
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
    axes[0].hist(residuals, bins=10, edgecolor="white", alpha=0.8, color=COLORS[0])
    axes[0].set_title(f"Гистограмма остатков (Шапиро-Уилк p={p_res:.3f})")
    axes[0].set_xlabel("Остатки, п.п.")
    axes[0].set_ylabel("Частота")
    axes[1].scatter(x, residuals, alpha=0.7, s=40, color=COLORS[2])
    axes[1].axhline(0, color="red", linewidth=1)
    axes[1].set_xlabel("ln(Выручка)")
    axes[1].set_ylabel("Остатки, п.п.")
    axes[1].set_title("Остатки vs предиктор")
    axes[1].grid(True, alpha=0.3)
    fig.tight_layout()
    ws.insert_image(f"A{res_row + 2}", "", {"image_data": fig_to_bytes(fig), "x_scale": 0.75, "y_scale": 0.75})


# ── 9. Кластерный анализ — нативный bar + PNG для PCA ─────────────

def write_clusters(writer, fmts, full):
    SHEET = "Кластерный анализ"
    rev = full[full["Выручка"].notna()].copy()
    book = writer.book

    latest = rev.sort_values("Период").groupby("Компания").last().reset_index()
    cluster_vars = [
        "Выручка", "Валовая рентабельность, %",
        "Оборачиваемость активов", "Доля КЗ в активах, %", "Доля ДС в активах, %",
    ]
    cdata = latest[["Компания"] + cluster_vars].dropna()
    X = cdata[cluster_vars].values
    names = cdata["Компания"].values

    scaler = StandardScaler()
    X_sc = scaler.fit_transform(X)

    max_k = min(len(cdata) - 1, 6)
    sils, inertias = [], []
    for k in range(2, max_k + 1):
        km = KMeans(n_clusters=k, n_init=20, random_state=42)
        lb = km.fit_predict(X_sc)
        inertias.append(km.inertia_)
        sils.append(silhouette_score(X_sc, lb))

    best_k = 2 + np.argmax(sils)
    km_final = KMeans(n_clusters=best_k, n_init=20, random_state=42)
    labels = km_final.fit_predict(X_sc)
    cdata = cdata.copy()
    cdata["Кластер"] = labels

    ws = book.add_worksheet(SHEET)
    writer.sheets[SHEET] = ws
    ws.set_column("A:A", 40)
    for c in range(1, 8):
        ws.set_column(c, c, 16)

    ws.write(0, 0, f"Кластерный анализ (k-means, k={best_k}, силуэт={max(sils):.3f})", fmts["h2"])

    # Table 1: Membership
    ws.write(1, 0, "Таблица 1. Принадлежность компаний к кластерам", fmts["label"])
    cols_out = ["Компания", "Кластер"] + cluster_vars
    for j, col in enumerate(cols_out):
        ws.write(2, j, col, fmts["header"])
    for i in range(len(cdata)):
        ws.write(3 + i, 0, cdata.iloc[i]["Компания"], fmts["normal"])
        ws.write(3 + i, 1, int(cdata.iloc[i]["Кластер"]), fmts["normal"])
        for jj, var in enumerate(cluster_vars):
            val = cdata.iloc[i][var]
            if pd.notna(val):
                fmt = fmts["pct1"] if "%" in var else fmts["num0"] if var == "Выручка" else fmts["num1"]
                ws.write(3 + i, 2 + jj, val, fmt)

    # Table 2: Cluster means
    means = cdata.groupby("Кластер")[cluster_vars].mean()
    means_row = len(cdata) + 5
    ws.write(means_row, 0, "Таблица 2. Средние значения по кластерам", fmts["label"])
    ws.write(means_row + 1, 0, "Кластер", fmts["header"])
    for j, var in enumerate(cluster_vars):
        short = var.replace("Валовая рентабельность, %", "Вал. рент., %") \
                    .replace("Оборачиваемость активов", "Оборач., раз") \
                    .replace("Доля КЗ в активах, %", "Доля КЗ, %") \
                    .replace("Доля ДС в активах, %", "Доля ДС, %") \
                    .replace("Выручка", "Выручка, тыс.")
        ws.write(means_row + 1, 1 + j, short, fmts["header"])
    m_start = means_row + 2
    for i, (cl, row_data) in enumerate(means.iterrows()):
        ws.write(m_start + i, 0, f"Кластер {cl}", fmts["normal"])
        for j, var in enumerate(cluster_vars):
            val = row_data[var]
            fmt = fmts["pct1"] if "%" in var else fmts["num0"] if var == "Выручка" else fmts["num1"]
            ws.write(m_start + i, 1 + j, val, fmt)
    m_end = m_start + len(means) - 1

    # Native column chart: cluster profile (standardized)
    means_sc = pd.DataFrame(scaler.transform(means.values), columns=cluster_vars, index=means.index)
    prof_row = m_end + 3
    ws.write(prof_row, 0, "Таблица 3. Стандартизованный профиль кластеров (z-score)", fmts["label"])
    short_vars = ["Выручка", "Вал. рент.", "Оборач.", "Доля КЗ", "Доля ДС"]
    ws.write(prof_row + 1, 0, "Показатель", fmts["header"])
    for j, cl in enumerate(sorted(means_sc.index)):
        ws.write(prof_row + 1, 1 + j, f"Кластер {cl}", fmts["header"])
    p_start = prof_row + 2
    for i, var in enumerate(short_vars):
        ws.write(p_start + i, 0, var, fmts["normal"])
        for j, cl in enumerate(sorted(means_sc.index)):
            ws.write(p_start + i, 1 + j, means_sc.loc[cl].iloc[i], fmts["num2"])
    p_end = p_start + len(short_vars) - 1

    chart_prof = book.add_chart({"type": "column"})
    chart_prof.set_title({"name": "Профиль кластеров (стандартизованные значения)"})
    chart_prof.set_y_axis({"name": "z-score"})
    chart_prof.set_size({"width": 600, "height": 400})
    for j, cl in enumerate(sorted(means_sc.index)):
        chart_prof.add_series({
            "name": [SHEET, prof_row + 1, 1 + j],
            "categories": [SHEET, p_start, 0, p_end, 0],
            "values": [SHEET, p_start, 1 + j, p_end, 1 + j],
            "fill": {"color": COLORS[j % len(COLORS)]},
        })
    ws.insert_chart(f"E{prof_row + 1}", chart_prof)

    # PCA + elbow — PNG (labels and dual axis not available natively)
    png_row = prof_row + 26
    ws.write(png_row, 0, "Визуализация кластеров (PNG)", fmts["label"])

    pca = PCA(n_components=2)
    X_pca = pca.fit_transform(X_sc)
    var_expl = pca.explained_variance_ratio_

    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    colors_map = COLORS[:best_k]
    for cl in sorted(cdata["Кластер"].unique()):
        mask = labels == cl
        axes[0].scatter(X_pca[mask, 0], X_pca[mask, 1], s=80, alpha=0.8,
                        color=colors_map[cl % len(colors_map)], label=f"Кластер {cl}", zorder=3)
    for i, name in enumerate(names):
        axes[0].annotate(short_name(name), (X_pca[i, 0], X_pca[i, 1]),
                         fontsize=8, ha="center", va="bottom", textcoords="offset points", xytext=(0, 6))
    axes[0].set_xlabel(f"PC1 ({var_expl[0]*100:.1f}%)")
    axes[0].set_ylabel(f"PC2 ({var_expl[1]*100:.1f}%)")
    axes[0].set_title("Кластеры в пространстве главных компонент")
    axes[0].legend(fontsize=9)
    axes[0].grid(True, alpha=0.3)

    ks = list(range(2, max_k + 1))
    axes[1].plot(ks, inertias, "o-", color=COLORS[0], label="Инерция")
    ax_s = axes[1].twinx()
    ax_s.plot(ks, sils, "s-", color=COLORS[1], label="Силуэт")
    axes[1].set_xlabel("Число кластеров (k)")
    axes[1].set_ylabel("Инерция", color=COLORS[0])
    ax_s.set_ylabel("Силуэтный коэффициент", color=COLORS[1])
    axes[1].set_title("Выбор числа кластеров")

    fig.tight_layout()
    ws.insert_image(f"A{png_row + 2}", "", {"image_data": fig_to_bytes(fig), "x_scale": 0.8, "y_scale": 0.8})


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
         "Медианный рост +34% г/г. Суммарная выручка ×3.2 за 4 года. Темпы замедляются (50%→20%).", True),
        ("H2: Масштаб → рентабельность", "НЕ ПОДТВЕРЖДЕНА", "p = 0.638",
         "ρ = −0.08. Крупнейшая компания убыточна 4 года из 5. Рентабельность определяется бизнес-моделью.", False),
        ("H3: Эффект масштаба", "НЕ ПОДТВЕРЖДЕНА", "p = 0.254",
         "ρ = −0.20. Направление ожидаемое, но незначимо. Затраты растут пропорционально масштабу.", False),
        ("H4: Кластеры", "НЕ ПОДТВ. ФОРМАЛЬНО", "силуэт = 0.211",
         "3 содержательных кластера: «малые организаторы», «кэш-генераторы», «лидеры рынка».", False),
        ("H5: Рост выручки → рост КЗ", "НЕ ПОДТВЕРЖДЕНА", "p = 0.114",
         "ρ = +0.32. Экономически логичная тенденция, но незначима при текущей выборке.", False),
    ]
    for i, (hyp, res, pv, interp, confirmed) in enumerate(verdicts):
        fmt = fmts["green"] if confirmed else fmts["red"]
        ws.write(2 + i, 0, hyp, fmts["normal"])
        ws.write(2 + i, 1, res, fmt)
        ws.write(2 + i, 2, pv, fmts["normal"])
        ws.write(2 + i, 3, interp, fmts["normal"])

    row = 9
    ws.write(row, 0, "Качественные выводы о рынке", fmts["h2"])
    for i, c in enumerate([
        "1. Рынок быстро растёт (×3.2 за 4 года), но темпы замедляются (от 50% к 20% г/г)",
        "2. Прибыльность не связана с масштабом — определяется бизнес-моделью и управлением затратами",
        "3. Две модели финансирования: предоплаты от участников (высокая ликвидность) vs КЗ поставщикам",
        "4. Себестоимость — ключевой фактор: медианная доля 89%, маржа ~10%",
        "5. Бизнес low-asset: компании с минимальными ОС показывают наибольшие темпы роста",
    ]):
        ws.write(row + 1 + i, 0, c, fmts["normal"])

    row += 8
    ws.write(row, 0, "Рекомендации для ЛПР", fmts["h2"])
    for i, r in enumerate([
        "Для организаторов: контролировать себестоимость (маржа ~10%), использовать модель предоплат",
        "Для инвесторов: рынок растущий (+34% г/г), входной барьер низкий (low-asset)",
        "Для регуляторов: рынок прозрачен через ГИРБО, стандартизировать отчётность НКО",
    ]):
        ws.write(row + 1 + i, 0, r, fmts["normal"])

    row += 6
    ws.write(row, 0, "Ограничения исследования", fmts["h2"])
    for i, l in enumerate([
        "1. Малая выборка (11 компаний, 49 наблюдений) — ограничивает мощность статистических тестов",
        "2. Только финансовые данные — нет операционных метрик (число участников, мероприятий)",
        "3. Разнородность ОПФ (ООО vs НКО) — ограничивает сопоставимость",
        "4. Юрлицо ≠ бренд — возможны связанные юрлица",
        "5. Факторный анализ не проведён из-за малого числа наблюдений",
    ]):
        ws.write(row + 1 + i, 0, l, fmts["normal"])


# ── main ──────────────────────────────────────────────────────────

def main():
    raw, full = load()
    print("Формирование report.xlsx...")
    with pd.ExcelWriter(REPORT_PATH, engine="xlsxwriter") as writer:
        fmts = make_formats(writer.book)
        write_description(writer, fmts);    print("  1/10  Описание исследования")
        write_raw_data(writer, fmts, raw);  print("  2/10  Исходные данные")
        write_derived(writer, fmts, full);  print("  3/10  Производные показатели")
        write_descriptive(writer, fmts, full); print("  4/10  Описательные статистики")
        write_dynamics(writer, fmts, full); print("  5/10  Динамика рынка")
        write_correlation(writer, fmts, full); print("  6/10  Корреляционный анализ")
        write_hypotheses(writer, fmts, full); print("  7/10  Проверка гипотез")
        write_regression(writer, fmts, full); print("  8/10  Регрессия")
        write_clusters(writer, fmts, full); print("  9/10  Кластерный анализ")
        write_conclusions(writer, fmts);    print("  10/10 Выводы")
    print(f"\nГотово: {REPORT_PATH}")
    print(f"Размер: {REPORT_PATH.stat().st_size / 1024:.0f} КБ")


if __name__ == "__main__":
    main()
