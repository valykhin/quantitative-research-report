"""Сбор финансовой отчётности компаний-организаторов беговых мероприятий из ГИРБО (bo.nalog.gov.ru)."""

import csv
import json
import time
from pathlib import Path

import httpx
import pandas as pd

BASE = "https://bo.nalog.gov.ru"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)",
    "Accept": "application/json",
}

COMPANIES_CSV = Path(__file__).parent / "companies.csv"


def load_companies() -> list[tuple[str, str]]:
    with open(COMPANIES_CSV, encoding="utf-8") as f:
        return [(row["name"], row["inn"]) for row in csv.DictReader(f)]

BALANCE_FIELDS = {
    "current1110": "Нематериальные активы",
    "current1150": "Основные средства",
    "current1100": "Внеоборотные активы",
    "current1210": "Запасы",
    "current1230": "Дебиторская задолженность",
    "current1240": "Краткосрочные финансовые вложения",
    "current1250": "Денежные средства",
    "current1200": "Оборотные активы",
    "current1600": "Активы всего",
    "current1310": "Уставный капитал",
    "current1370": "Нераспределённая прибыль (убыток)",
    "current1300": "Капитал и резервы",
    "current1410": "Заёмные средства (долгосрочные)",
    "current1400": "Долгосрочные обязательства",
    "current1510": "Заёмные средства (краткосрочные)",
    "current1520": "Кредиторская задолженность",
    "current1500": "Краткосрочные обязательства",
}

PNL_FIELDS = {
    "current2110": "Выручка",
    "current2120": "Себестоимость продаж",
    "current2100": "Валовая прибыль",
    "current2210": "Коммерческие расходы",
    "current2220": "Управленческие расходы",
    "current2200": "Прибыль от продаж",
    "current2310": "Доходы от участия в других организациях",
    "current2320": "Проценты к получению",
    "current2330": "Проценты к уплате",
    "current2340": "Прочие доходы",
    "current2350": "Прочие расходы",
    "current2300": "Прибыль до налогообложения",
    "current2400": "Чистая прибыль (убыток)",
}


def search_org(client: httpx.Client, inn: str) -> list[dict]:
    url = f"{BASE}/advanced-search/organizations/search?query={inn}&page=0&size=20"
    r = client.get(url)
    r.raise_for_status()
    return r.json()["content"]


def get_bfo(client: httpx.Client, org_id: int) -> list[dict]:
    url = f"{BASE}/nbo/organizations/{org_id}/bfo/"
    r = client.get(url)
    r.raise_for_status()
    return r.json()


def extract_row(company_name: str, inn: str, period: str, balance: dict | None, pnl: dict | None) -> dict:
    row = {"Компания": company_name, "ИНН": inn, "Период": int(period)}

    if balance:
        for field, label in BALANCE_FIELDS.items():
            row[label] = balance.get(field)

    if pnl:
        for field, label in PNL_FIELDS.items():
            row[label] = pnl.get(field)

    return row


def main():
    out_dir = Path(__file__).parent / "data"
    out_dir.mkdir(exist_ok=True)

    all_rows = []
    seen_org_ids = set()

    companies = load_companies()
    print(f"Загружено {len(companies)} компаний из {COMPANIES_CSV.name}")

    with httpx.Client(headers=HEADERS, follow_redirects=True, timeout=20) as client:
        for name, inn in companies:
            print(f"\n{'='*60}")
            print(f"Поиск: {name} (ИНН {inn})")

            orgs = search_org(client, inn)
            if not orgs:
                print(f"  НЕ НАЙДЕНА")
                continue

            for org in orgs:
                org_id = org["id"]
                org_name = org.get("shortName", name)
                org_inn = org["inn"].replace("<strong>", "").replace("</strong>", "")

                if org_id in seen_org_ids:
                    print(f"  {org_name} (id={org_id}) — уже обработана, пропуск")
                    continue
                seen_org_ids.add(org_id)

                print(f"  Найдена: {org_name} (id={org_id}, status={org.get('statusCode')})")
                time.sleep(0.5)

                bfo_list = get_bfo(client, org_id)
                print(f"  Периодов: {len(bfo_list)}")

                for bfo in bfo_list:
                    period = bfo["period"]
                    corrections = bfo.get("typeCorrections", [])
                    if not corrections:
                        continue

                    corr = corrections[0].get("correction", {})
                    balance = corr.get("balance")
                    pnl = corr.get("financialResult")

                    row = extract_row(org_name, org_inn, period, balance, pnl)
                    all_rows.append(row)
                    print(f"    {period}: balance={'OK' if balance else 'NO'}, pnl={'OK' if pnl else 'NO'}")

                time.sleep(0.5)

    df = pd.DataFrame(all_rows)
    df = df.sort_values(["Компания", "Период"]).reset_index(drop=True)

    out_path = out_dir / "companies_financials.xlsx"
    df.to_excel(out_path, index=False, engine="xlsxwriter")
    print(f"\nСохранено: {out_path}")
    print(f"Строк: {len(df)}, Компаний: {df['Компания'].nunique()}")
    print(f"\nКолонки: {list(df.columns)}")

    raw_path = out_dir / "raw_bfo.json"
    with open(raw_path, "w", encoding="utf-8") as f:
        json.dump(all_rows, f, ensure_ascii=False, indent=2, default=str)
    print(f"Сырые данные: {raw_path}")


if __name__ == "__main__":
    main()
