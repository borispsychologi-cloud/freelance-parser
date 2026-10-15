import requests
import time
import csv
import json
import logging
import re
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, numbers
import schedule
import argparse

def get_args():
    parser = argparse.ArgumentParser(description="Парсер Wildberries с мониторингом цен")
    parser.add_argument("--query", type=str, default="кроссовки", help="Поисковый запрос (например, 'кроссовки мужские')")
    parser.add_argument("--schedule", action="store_true", help="Запускать парсер ежедневно по расписанию")

    return parser.parse_args()


# ──────────────────────────────────────────────
# НАСТРОЙКИ ПАРСЕРА
# ──────────────────────────────────────────────
MAX_PAGES = 3
MIN_PRICE = 0
MAX_PRICE = 1000000
MIN_RATING = 0.0
MIN_FEEDBACKS = 0
SORT_BY = "price_asc"
DELAY = 4

# Папка для хранения истории
HISTORY_DIR = "wb_history"
RUN_MODE = "once"  # "once" — один запуск, "schedule" — по расписанию
RUN_TIME = "09:00"  # Время запуска при режиме schedule


# ──────────────────────────────────────────────
# ЛОГИРОВАНИЕ
# ──────────────────────────────────────────────
logging.basicConfig(
    filename="wb_parser.log",
    level=logging.INFO,
    format="%(asctime)s — %(levelname)s — %(message)s",
    encoding="utf-8",
)
log = logging.getLogger()


# ──────────────────────────────────────────────
# ПАРСИНГ
# ──────────────────────────────────────────────
def fetch_with_retry(url, headers, max_retries=3):
    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=headers, timeout=15)
            if response.status_code == 200:
                return response
            log.warning(f"Попытка {attempt+1}: статус {response.status_code}")
        except requests.RequestException as e:
            log.warning(f"Попытка {attempt+1}: ошибка {e}")
        time.sleep(2)
    return None


def parse_page(query, page, headers):
    """Парсит одну страницу поиска WB."""
    url = (
        f"https://search.wb.ru/exactmatch/ru/common/v5/search?"
        f"appType=1&curr=rub&dest=-1257786&lang=ru&locale=ru"
        f"&page={page}&query={query}&resultset=catalog&sort=popular&spp=30"
        f"&suppressSpellcheck=false"
    )
    response = fetch_with_retry(url, headers)
    if not response:
        log.error(f"Не удалось получить страницу {page}")
        return []

    try:
        data = response.json()
    except Exception as e:
        log.error(f"Ошибка JSON на странице {page}: {e}")
        return []

    # WB меняет структуру: иногда products в корне, иногда внутри data
    products = data.get("products") or data.get("data", {}).get("products", [])
    results = []

    for p in products:
        # Цена: может быть в sizes или в корне
        sizes = p.get("sizes", [{}])
        if sizes and isinstance(sizes, list):
            price_raw = sizes[0].get("price", {})
        else:
            price_raw = {}

        price = price_raw.get("product", 0) / 100 if price_raw else 0
        price_total = price_raw.get("total", 0) / 100 if price_raw else 0

        # Альтернативные поля цены (WB меняет структуру)
        if not price:
            price = p.get("salePriceU", 0) / 100
        if not price_total:
            price_total = p.get("priceU", 0) / 100

        item = {
            "id": p.get("id", ""),
            "name": p.get("name", "Без названия"),
            "brand": p.get("brand", "—"),
            "price": int(price) if price else 0,
            "price_total": int(price_total) if price_total else 0,
            "sale": p.get("sale", 0),
            "cashback": p.get("cashback", 0),
            "rating": p.get("rating", 0),
            "feedbacks": p.get("feedbacks", 0),
            "review_rating": p.get("reviewRating", 0),
            "supplier": p.get("supplier", "—"),
            "supplier_rating": p.get("supplierRating", 0),
            "promo": p.get("promoTextCard", ""),
            "url": f"https://www.wildberries.ru/catalog/{p.get('id')}/detail.aspx",
        }
        results.append(item)

    return results




def apply_filters(items, min_price, max_price, min_rating, min_feedbacks):
    filtered = []
    for item in items:
        if not (min_price <= item["price"] <= max_price):
            continue
        if item["rating"] < min_rating:
            continue
        if item["feedbacks"] < min_feedbacks:
            continue
        filtered.append(item)
    return filtered


def sort_items(items, sort_by):
    sort_map = {
        "price_asc": lambda x: x["price"],
        "price_desc": lambda x: -x["price"],
        "rating_desc": lambda x: -x["rating"],
        "feedbacks_desc": lambda x: -x["feedbacks"],
    }
    if sort_by in sort_map:
        return sorted(items, key=sort_map[sort_by])
    return items


# ──────────────────────────────────────────────
# СОХРАНЕНИЕ В EXCEL
# ──────────────────────────────────────────────
def save_to_excel(items, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Товары WB"

    headers = [
        "Артикул", "Название", "Бренд", "Цена (₽)",
        "Цена со скидкой (₽)", "Скидка (%)", "Кэшбэк",
        "Рейтинг", "Отзывов", "Рейтинг отзывов",
        "Продавец", "Рейтинг продавца", "Промо", "Ссылка",
    ]

    header_fill = PatternFill(start_color="FF6700", end_color="FF6700", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, item in enumerate(items, 2):
        ws.cell(row=row_idx, column=1, value=item["id"])
        ws.cell(row=row_idx, column=2, value=item["name"])
        ws.cell(row=row_idx, column=3, value=item["brand"])
        ws.cell(row=row_idx, column=4, value=item["price"])
        ws.cell(row=row_idx, column=5, value=item["price_total"])
        ws.cell(row=row_idx, column=6, value=item["sale"])
        ws.cell(row=row_idx, column=7, value=item["cashback"])
        ws.cell(row=row_idx, column=8, value=item["rating"])
        ws.cell(row=row_idx, column=9, value=item["feedbacks"])
        ws.cell(row=row_idx, column=10, value=item["review_rating"])
        ws.cell(row=row_idx, column=11, value=item["supplier"])
        ws.cell(row=row_idx, column=12, value=item["supplier_rating"])
        ws.cell(row=row_idx, column=13, value=item["promo"])
        link_cell = ws.cell(row=row_idx, column=14, value=item["url"])
        link_cell.hyperlink = item["url"]
        link_cell.font = Font(color="0563C1", underline="single")

    col_widths = [12, 45, 20, 12, 18, 12, 10, 10, 10, 14, 25, 16, 30, 50]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=len(headers)).column_letter}1"
    wb.save(filename)


# ──────────────────────────────────────────────
# РАБОТА С ИСТОРИЕЙ (СРАВНЕНИЕ ДНЕЙ)
# ──────────────────────────────────────────────
def save_history(items, query):
    """Сохраняет данные текущего запуска в JSON."""
    os.makedirs(HISTORY_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(HISTORY_DIR, f"wb_{query}_{timestamp}.json")
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)
    log.info(f"История сохранена: {filename}")
    return filename



def load_previous_history(query, exclude_file=None):
    """Загружает самый свежий предыдущий файл истории."""
    if not os.path.exists(HISTORY_DIR):
        return None, None

    files = [f for f in os.listdir(HISTORY_DIR)
             if f.startswith(f"wb_{query}_") and f.endswith(".json")]

    if exclude_file:
        exclude_basename = os.path.basename(exclude_file)
        files = [f for f in files if f != exclude_basename]

    if not files:
        return None, None

    files.sort(reverse=True)
    latest = files[0]
    filepath = os.path.join(HISTORY_DIR, latest)

    # Извлекаем дату из имени файла (формат: wb_query_20260912_195424.json)
    date_str = latest.replace(f"wb_{query}_", "").replace(".json", "")
    try:
        file_date = datetime.strptime(date_str, "%Y%m%d_%H%M%S").strftime("%d.%m.%Y %H:%M")
    except ValueError:
        file_date = date_str

    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)

    log.info(f"Загружена история: {latest}")
    return data, file_date



def compare_with_previous(current_items, previous_items, prev_date):
    """Сравнивает текущие данные с предыдущими."""
    current_map = {item["id"]: item for item in current_items}
    previous_map = {item["id"]: item for item in previous_items}

    changes = []

    # Товары, которые были и стали — сравниваем цены
    for item_id, curr in current_map.items():
        prev = previous_map.get(item_id)
        if not prev:
            # Новые товары
            changes.append({
                "id": item_id,
                "name": curr["name"],
                "brand": curr["brand"],
                "old_price": None,
                "new_price": curr["price"],
                "diff": None,
                "diff_percent": None,
                "old_feedbacks": None,
                "new_feedbacks": curr["feedbacks"],
                "new_feedbacks_diff": None,
                "old_rating": None,
                "new_rating": curr["rating"],
                "rating_diff": None,
                "status": "🆕 Новые",
                "url": curr["url"],
            })
            continue

        price_diff = curr["price"] - prev["price"]
        price_diff_pct = round((price_diff / prev["price"] * 100), 1) if prev["price"] else 0
        feedbacks_diff = curr["feedbacks"] - prev["feedbacks"]
        rating_diff = round(curr["rating"] - prev["rating"], 2)

        if price_diff != 0 or feedbacks_diff != 0 or rating_diff != 0:
            if price_diff < 0:
                status = "📉 Подешевело"
            elif price_diff > 0:
                status = "📈 Подорожало"
            elif feedbacks_diff > 0:
                status = "💬 Новые отзывы"
            elif rating_diff != 0:
                status = "⭐ Изм. рейтинг"
            else:
                status = "📝 Изменение"

            changes.append({
                "id": item_id,
                "name": curr["name"],
                "brand": curr["brand"],
                "old_price": prev["price"],
                "new_price": curr["price"],
                "diff": price_diff,
                "diff_percent": price_diff_pct,
                "old_feedbacks": prev["feedbacks"],
                "new_feedbacks": curr["feedbacks"],
                "new_feedbacks_diff": feedbacks_diff,
                "old_rating": prev["rating"],
                "new_rating": curr["rating"],
                "rating_diff": rating_diff,
                "status": status,
                "url": curr["url"],
            })

    # Пропавшие товары
    for item_id, prev in previous_map.items():
        if item_id not in current_map:
            changes.append({
                "id": item_id,
                "name": prev["name"],
                "brand": prev["brand"],
                "old_price": prev["price"],
                "new_price": None,
                "diff": None,
                "diff_percent": None,
                "old_feedbacks": prev["feedbacks"],
                "new_feedbacks": None,
                "new_feedbacks_diff": None,
                "old_rating": prev["rating"],
                "new_rating": None,
                "rating_diff": None,
                "status": "❌ Пропал",
                "url": prev["url"],
            })

    # Сортировка: сначала изменения цены, потом остальное
    status_order = {
        "📈 Подорожало": 0,
        "📉 Подешевело": 1,
        "🆕 Новые": 2,
        "❌ Пропал": 3,
        "💬 Новые отзывы": 4,
        "⭐ Изм. рейтинг": 5,
        "📝 Изменение": 6,
    }
    changes.sort(key=lambda x: status_order.get(x["status"], 99))

    return changes


def save_changes_excel(changes, filename, prev_date):
    """Сохраняет отчёт об изменениях в Excel с подсветкой."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Изменения цен"

    today_str = datetime.now().strftime("%d.%m.%Y")
    title = f"Отчёт: {prev_date} → {today_str}"

    # Заголовок отчёта
    ws.cell(row=1, column=1, value=title)
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.merge_cells("A1:K1")

    headers = [
        "Статус", "Артикул", "Название", "Бренд",
        "Цена вчера (₽)", "Цена сегодня (₽)", "Разница (₽)", "Разница (%)",
        "Отзывов вчера", "Отзывов сегодня", "Рейтинг вчера", "Рейтинг сегодня", "Ссылка",
    ]

    header_row = 3
    header_fill = PatternFill(start_color="FF6700", end_color="FF6700", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Цвета для статусов
    fill_red = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
    fill_green = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    fill_blue = PatternFill(start_color="BBDEFB", end_color="BBDEFB", fill_type="solid")
    fill_gray = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")

    for row_idx, ch in enumerate(changes, header_row + 1):
        ws.cell(row=row_idx, column=1, value=ch["status"])
        ws.cell(row=row_idx, column=2, value=ch["id"])
        ws.cell(row=row_idx, column=3, value=ch["name"])
        ws.cell(row=row_idx, column=4, value=ch["brand"])
        ws.cell(row=row_idx, column=5, value=ch["old_price"] if ch["old_price"] is not None else "—")
        ws.cell(row=row_idx, column=6, value=ch["new_price"] if ch["new_price"] is not None else "—")
        ws.cell(row=row_idx, column=7, value=ch["diff"] if ch["diff"] is not None else "—")
        ws.cell(row=row_idx, column=8, value=f"{ch['diff_percent']}%" if ch["diff_percent"] is not None else "—")
        ws.cell(row=row_idx, column=9, value=ch["old_feedbacks"] if ch["old_feedbacks"] is not None else "—")
        ws.cell(row=row_idx, column=10, value=ch["new_feedbacks"] if ch["new_feedbacks"] is not None else "—")
        ws.cell(row=row_idx, column=11, value=ch["old_rating"] if ch["old_rating"] is not None else "—")
        ws.cell(row=row_idx, column=12, value=ch["new_rating"] if ch["new_rating"] is not None else "—")
        link_cell = ws.cell(row=row_idx, column=13, value=ch["url"])
        link_cell.hyperlink = ch["url"]
        link_cell.font = Font(color="0563C1", underline="single")

        # Подсветка по статусусу
        if "Подорожало" in ch["status"]:
            row_fill = fill_red
        elif "Подешевело" in ch["status"]:
            row_fill = fill_green
        elif "Новые" in ch["status"]:
            row_fill = fill_blue
        elif "Пропал" in ch["status"]:
            row_fill = fill_gray
        else:
            row_fill = fill_yellow

        for col in range(1, 14):
            ws.cell(row=row_idx, column=col).fill = row_fill

    col_widths = [16, 12, 45, 20, 16, 16, 14, 14, 14, 14, 14, 14, 50]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=header_row, column=i).column_letter].width = w

    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:{ws.cell(row=header_row, column=13).column_letter}{header_row}"
    wb.save(filename)


# ──────────────────────────────────────────────
# ГЛАВНЫЙ ЗАПУСК
# ──────────────────────────────────────────────
def run_parser():
    print(f"\n{'='*50}")
    print(f"Запуск парсера: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
    print(f"{'='*50}")

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36",
        "Accept": "*/*",
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "Origin": "https://www.wildberries.ru",
        "Referer": "https://www.wildberries.ru/",
    }


    all_items = []
    for page in range(1, MAX_PAGES + 1):
        print(f"Страница {page}/{MAX_PAGES}...")
        items = parse_page(SEARCH_QUERY, page, headers)

        # До 3 попыток при пустой странице
        for retry in range(3):
            if items:
                break
            wait = 5 + retry * 5  # 5, 10, 15 сек
            print(f"  Пусто. Повтор через {wait} сек (попытка {retry+1}/3)...")
            time.sleep(wait)
            items = parse_page(SEARCH_QUERY, page, headers)

        print(f"  Получено: {len(items)} товаров")
        log.info(f"Страница {page}: {len(items)} товаров")
        all_items.extend(items)
        time.sleep(4)






    filtered = apply_filters(all_items, MIN_PRICE, MAX_PRICE, MIN_RATING, MIN_FEEDBACKS)
    filtered = sort_items(filtered, SORT_BY)
    print(f"\nВсего после фильтрации: {len(filtered)} товаров")

    if not filtered:
        print("Нет товаров по фильтрам.")
        return

    # Сохраняем текущие данные в Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_file = f"wb_{SEARCH_QUERY}_{timestamp}.xlsx"
    save_to_excel(filtered, excel_file)
    print(f"✓ Excel: {excel_file}")


    # Сохраняем в историю
    hist_file = save_history(filtered, SEARCH_QUERY)


    # Сравниваем с предыдущим днём
    today = datetime.now().strftime("%Y%m%d")
    prev_data, prev_date = load_previous_history(SEARCH_QUERY, exclude_file=hist_file)


    if prev_data:
        print(f"\n--- Сравнение с предыдущим запуском ({prev_date}) ---")
        changes = compare_with_previous(filtered, prev_data, prev_date)

        if changes:
            print(f"Найдено изменений: {len(changes)}")
            for ch in changes[:15]:
                if ch["diff"] is not None:
                    sign = "+" if ch["diff"] > 0 else ""
                    print(f"  {ch['status']} | {ch['name'][:40]} | "
                          f"{ch['old_price']} → {ch['new_price']} ₽ ({sign}{ch['diff']} ₽)")
                else:
                    print(f"  {ch['status']} | {ch['name'][:40]}")

            # Сохраняем отчёт
            report_file = f"wb_{SEARCH_QUERY}_report_{timestamp}.xlsx"
            save_changes_excel(changes, report_file, prev_date)
            print(f"\n✓ Отчёт об изменениях: {report_file}")
            log.info(f"Отчёт: {len(changes)} изменений → {report_file}")
        else:
            print("Изменений не найдено. Цены стабильны.")
            log.info("Изменений не найдено.")
    else:
        print("\nПредыдущих данных нет — это первый запуск. Сравнение будет доступно со следующего запуска.")

    print(f"\nГотово! {datetime.now().strftime('%H:%M:%S')}")


def scheduled_main():
    print(f"Расписание: запуск каждый день в {RUN_TIME}")
    print("Ожидание... (Ctrl+C для остановки)\n")

    schedule.every().day.at(RUN_TIME).do(run_parser)
    run_parser()  # первый запуск сразу

    while True:
        schedule.run_pending()
        time.sleep(60)


def main():
    args = get_args()

    global SEARCH_QUERY
    SEARCH_QUERY = args.query

    print(f"Запускаю парсер для запроса: '{SEARCH_QUERY}'")

    if args.schedule:
        scheduled_main()
    else:
        run_parser()


if __name__ == "__main__":
    main()




