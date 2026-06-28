# Структура проекта и описание файлов

## Управление проектом

| Файл | Описание |
|---|---|
| [AGENT.md](AGENT.md) | Цель проекта, исходные данные, план реализации, ссылки на все артефакты |
| [IMPLEMENTATION_PLAN.md](IMPLEMENTATION_PLAN.md) | Поэтапный план (9 этапов) с чеклистами шагов и статусами выполнения |
| [REQUIREMENTS.md](REQUIREMENTS.md) | 39 требований к приёмке работы с отметками выполнения (38/39 выполнено) |
| [FILES.md](FILES.md) | Этот файл — описание всех файлов проекта |

## Исходные материалы (`sources/`)

| Файл | Описание |
|---|---|
| [data.txt](sources/data.txt) | Список компаний для анализа (бренд + ИНН) |
| [Задание на исследование рынка, ориентированного на сбыт (1).docx](sources/Задание%20на%20исследование%20рынка,%20ориентированного%20на%20сбыт%20(1).docx) | Методическое задание |
| [Примеры для отчета (1).pptx](sources/Примеры%20для%20отчета%20(1).pptx) | Рекомендации по написанию отчёта, примеры визуализации и анализа |
| [Анализ+данных+ПащенкоТВ+Банкроты+ПК.xlsx](sources/Анализ+данных+ПащенкоТВ+Банкроты+ПК.xlsx) | Пример аналогичного исследования (~250 компаний-банкротов Пермского края) |

## Результаты исследования (`output/`)

| Файл | Описание |
|---|---|
| [report.md](output/report.md) | Итоговый отчёт в markdown (цели → данные → методы → результаты → выводы) |
| [stage1_research_design.md](output/stage1_research_design.md) | Этап 1: проблема, цель, 5 гипотез, описание рынка |
| [stage2_data_sources.md](output/stage2_data_sources.md) | Этап 2: источник (ГИРБО), период, обоснование, ограничения |
| [stage3_variables.md](output/stage3_variables.md) | Этап 3: 30 переменных, типы, обоснование методов анализа |
| [stage4_descriptive.md](output/stage4_descriptive.md) | Этап 4: описательные статистики, нормальность, графики |
| [stage5_inferential.md](output/stage5_inferential.md) | Этап 5: тесты гипотез, корреляционная матрица |
| [stage6_multivariate.md](output/stage6_multivariate.md) | Этап 6: регрессия, кластерный анализ |
| [stage7_conclusions.md](output/stage7_conclusions.md) | Этап 7: вердикты по гипотезам, выводы, рекомендации |

## Python-проект (`app/`)

### Конфигурация

| Файл | Описание |
|---|---|
| [README.md](app/README.md) | Инструкция по запуску: установка, пайплайн, описание скриптов |
| [pyproject.toml](app/pyproject.toml) | Зависимости Python-проекта (pandas, scipy, sklearn, matplotlib и др.) |
| [companies.csv](app/companies.csv) | Список компаний для сбора данных (name, inn) — единственный файл для редактирования при расширении выборки |

### Скрипты

| Файл | Что делает | Команда |
|---|---|---|
| [collect_data.py](app/collect_data.py) | Сбор финансовой отчётности из API ГИРБО по списку из `companies.csv` | `uv run python collect_data.py` |
| [prepare_dataset.py](app/prepare_dataset.py) | Расчёт производных показателей (рентабельность, темпы роста, доли и др.) | `uv run python prepare_dataset.py` |
| [descriptive_analysis.py](app/descriptive_analysis.py) | Описательные статистики, тест нормальности, графики | `uv run python descriptive_analysis.py` |
| [inferential_analysis.py](app/inferential_analysis.py) | Проверка гипотез (Вилкоксон, Спирмен, хи-квадрат), корреляционная матрица | `uv run python inferential_analysis.py` |
| [multivariate_analysis.py](app/multivariate_analysis.py) | Регрессия (МНК) и кластерный анализ (k-means) | `uv run python multivariate_analysis.py` |
| [build_report.py](app/build_report.py) | Сборка итогового xlsx-отчёта (10 листов с данными, расчётами и графиками) | `uv run python build_report.py` |

### Данные (`app/data/`)

| Файл | Описание |
|---|---|
| [companies_financials.xlsx](app/data/companies_financials.xlsx) | Сырая финансовая отчётность 11 компаний за 2021–2025 (49 строк, 33 столбца) |
| [raw_bfo.json](app/data/raw_bfo.json) | Сырые данные из API ГИРБО (JSON) |
| [analysis_dataset.xlsx](app/data/analysis_dataset.xlsx) | Полный аналитический датасет с производными показателями (49 строк, 46 столбцов) |
| [analysis_revenue_subset.xlsx](app/data/analysis_revenue_subset.xlsx) | Выборка с данными по выручке (38 строк, 9 компаний) |
| [mean_changes.xlsx](app/data/mean_changes.xlsx) | Средние абсолютные изменения показателей по компаниям |

### Результаты расчётов (`app/output/`)

| Файл | Описание |
|---|---|
| [report.xlsx](app/output/report.xlsx) | **Итоговый отчёт** — 10 листов с данными, таблицами, графиками (самодостаточный для презентации) |
| [descriptive_stats.xlsx](app/output/descriptive_stats.xlsx) | Описательные статистики (среднее, медиана, σ, асимметрия, эксцесс, IQR) |
| [normality_tests.xlsx](app/output/normality_tests.xlsx) | Результаты теста Шапиро-Уилка |
| [hypothesis_tests.xlsx](app/output/hypothesis_tests.xlsx) | Сводка тестов гипотез H1–H5 |
| [correlation_matrix.xlsx](app/output/correlation_matrix.xlsx) | Корреляционная матрица Спирмена + p-values |
| [regression_results.xlsx](app/output/regression_results.xlsx) | Коэффициенты регрессии, R², проверка условий МНК |
| [cluster_results.xlsx](app/output/cluster_results.xlsx) | Принадлежность к кластерам, средние по кластерам |

### Графики (`app/output/plots/`)

| Файл | Описание |
|---|---|
| [dynamics.png](app/output/plots/dynamics.png) | Динамика выручки, прибыли, рентабельности и активов по компаниям (2021–2025) |
| [market_summary.png](app/output/plots/market_summary.png) | Сводные показатели рынка (суммарная/медианная выручка, рентабельность) |
| [boxplots.png](app/output/plots/boxplots.png) | Распределение показателей по компаниям (boxplot) |
| [histograms.png](app/output/plots/histograms.png) | Гистограммы распределения ключевых переменных |
| [revenue_growth.png](app/output/plots/revenue_growth.png) | Темпы роста выручки по годам |
| [correlation_heatmap.png](app/output/plots/correlation_heatmap.png) | Тепловая карта корреляций Спирмена |
| [hypothesis_scatters.png](app/output/plots/hypothesis_scatters.png) | Scatter-графики для гипотез H2, H3, H5 |
| [regression.png](app/output/plots/regression.png) | Регрессия + гистограмма остатков + остатки vs предиктор |
| [clusters.png](app/output/plots/clusters.png) | Выбор k + PCA scatter + профиль кластеров |
