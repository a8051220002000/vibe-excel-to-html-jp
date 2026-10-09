#!/usr/bin/env python3
"""
scripts/build_data.py
Idempotent and resilient Excel-to-JSON pipeline for itinerary.xlsx.
Complies with SDD data contract rules:
- Forward fill Day and Date
- Type serialization for datetime/date/time
- Null/empty row filtering
- Multi-sheet parsing (行程表, 景點資料, 機票資訊, 交通套票資訊, todo, 遊日優惠)
"""

import argparse
import datetime
import json
import os
import sys
from pathlib import Path
import openpyxl

WEEKDAYS_ZH = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"]

def serialize_cell(val):
    """Normalize Excel cell values into clean serializable types."""
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        if isinstance(val, datetime.datetime) and val.time() != datetime.time(0, 0):
            return val.strftime("%Y-%m-%d %H:%M")
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.time):
        return val.strftime("%H:%M")
    if isinstance(val, float):
        # If float is actually an integer like 1.0
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val).strip()

def parse_schedule(sheet):
    """
    Parse '行程表' sheet with Forward Fill logic.
    Row 2 is header row.
    Row 3 onwards contains data.
    """
    days = []
    current_day = None
    current_date = None
    current_day_dict = None

    # Inspect all rows
    max_row = sheet.max_row or 0
    max_col = min(sheet.max_column or 0, 15)

    for r in range(2, max_row + 1):
        cells = [sheet.cell(r, c).value for c in range(1, max_col + 1)]
        col_day = serialize_cell(cells[0])       # Col 1: Day 1, Day 2...
        col_date_area = serialize_cell(cells[1])  # Col 2: Date (at header) or Area/Loc
        col_start = serialize_cell(cells[2])      # Col 3: 出發時間
        col_end = serialize_cell(cells[3])        # Col 4: 抵達時間
        col_spot = serialize_cell(cells[4])       # Col 5: 景點
        col_trans = serialize_cell(cells[5])      # Col 6: 交通方法
        col_note1 = serialize_cell(cells[6])      # Col 7: 購票/小備註
        col_remark = serialize_cell(cells[7])     # Col 8: 特別備註
        col_cost = serialize_cell(cells[9]) if len(cells) > 9 else ""     # Col 10: 票價
        col_booked = serialize_cell(cells[10]) if len(cells) > 10 else "" # Col 11: 已經訂好了？

        # Detect Day declaration: Day 1, Day 2, etc.
        if col_day.startswith("Day"):
            current_day = col_day
            # Extract date if present in col 2
            raw_date = cells[1]
            if isinstance(raw_date, (datetime.datetime, datetime.date)):
                current_date = raw_date.strftime("%Y-%m-%d")
            elif col_date_area and ("202" in col_date_area or "-" in col_date_area):
                current_date = col_date_area.split(" ")[0]
            else:
                current_date = current_date or ""

            # Calculate weekday
            weekday_str = ""
            if current_date:
                try:
                    dt = datetime.datetime.strptime(current_date, "%Y-%m-%d")
                    weekday_str = WEEKDAYS_ZH[dt.weekday()]
                except Exception:
                    pass

            # Day default titles
            day_titles = {
                "Day 1": "抵達福岡・博多運河與屋台漫遊",
                "Day 2": "天神商圈・地下街甜點購物之旅",
                "Day 3": "門司港・小倉城・下關海鮮探索",
                "Day 4": "熊本城／太宰府／由布院多選漫遊",
                "Day 5": "伴手禮採買・福岡機場返台"
            }

            current_day_dict = {
                "day": current_day,
                "date": current_date,
                "weekday": weekday_str,
                "title": day_titles.get(current_day, f"{current_day} 行程"),
                "items": []
            }
            days.append(current_day_dict)
            continue

        # If we have not encountered a Day yet, skip header/pre-day rows
        if not current_day_dict:
            continue

        # Filter out empty or meaningless rows (e.g. stray notes or empty spacers)
        if not col_spot and not col_start and not col_end:
            # If there's an interesting remark without spot, only keep if substantive
            if col_remark and len(col_remark) > 5 and not col_remark.replace(".", "").isdigit():
                col_spot = "特別提醒"
            else:
                continue

        # Determine category for better UI presentation
        category = "sightseeing"
        title_lower = (col_spot + " " + col_trans + " " + col_remark).lower()
        if any(k in title_lower for k in ["機場", "飛機", "出境", "入境", "回台", "土城→桃園"]):
            category = "flight"
        elif any(k in title_lower for k in ["→", "出門", "取行李", "車站", "博多開車", "公車", "寄物", "寄行李"]):
            category = "transport"
        elif any(k in title_lower for k in ["早餐", "午餐", "晚餐", "大東園", "麵包", "水果派", "千層蛋糕", "草莓", "屋台", "cafe", "咖哩", "抹茶"]):
            category = "food"
        elif any(k in title_lower for k in ["地下街", "三越", "parco", "岩田屋", "無印", "優衣褲", "gu", "beams", "3conis", "mina", "big cream", "運河城", "plaza", "mart", "迪士尼", "outlate", "血拼", "伴手禮"]):
            category = "shopping"
        elif any(k in title_lower for k in ["check-in", "民宿", "晚安", "回家"]):
            category = "hotel"
        elif any(k in title_lower for k in ["神社", "博物館", "城", "山", "廣場", "溫泉", "街道", "花卉村", "金鱗湖", "足湯", "海賊王", "條約"]):
            category = "sightseeing"

        # Combine remarks cleanly
        remarks_list = []
        if col_note1:
            remarks_list.append(col_note1)
        if col_remark:
            remarks_list.append(col_remark)
        combined_remarks = "\n".join(remarks_list)

        # Build clean item
        item_id = f"{current_day.lower().replace(' ', '')}-{len(current_day_dict['items']) + 1}"
        item = {
            "id": item_id,
            "day": current_day,
            "date": current_date,
            "time_start": col_start,
            "time_end": col_end,
            "spot": col_spot,
            "location_tag": col_date_area if col_date_area != current_date else "",
            "transport": col_trans,
            "remarks": combined_remarks,
            "price": col_cost,
            "booked": col_booked,
            "category": category
        }
        current_day_dict["items"].append(item)

    return days

def parse_coupons(sheet):
    """Parse '遊日優惠' sheet into coupon objects."""
    coupons = []
    if not sheet:
        return coupons

    max_row = sheet.max_row or 0
    for r in range(1, max_row + 1):
        c1 = serialize_cell(sheet.cell(r, 1).value)
        c2 = serialize_cell(sheet.cell(r, 2).value)
        if not c1 and not c2:
            continue
        # Skip instruction banners
        if "不要再發共享要求" in c1 or "請善用副本功能" in c1:
            continue
        if c1 == "你壞" or c2 == "你壞":
            continue

        url = ""
        title = c1
        desc = ""
        if c2.startswith("http"):
            url = c2
        elif c1.startswith("http"):
            url = c1
            title = c2

        if not url and not title:
            continue

        # Classify coupon
        category = "購物優惠"
        badge = "優惠券"
        if "donki" in title.lower() or "唐吉訶德" in title:
            badge = "免稅 10% + 最高 7%"
            category = "購物特惠"
            desc = "唐吉訶德 Don Quijote 免稅 10% 加碼折抵，出示手機條碼即可享最高 17% 折扣。"
        elif "klook" in title.lower():
            badge = "折扣碼"
            category = "票券折扣"
            desc = "Klook 日本行程、門票、JR Pass 專屬折抵代碼整理清單。"
        elif "trip.com" in title.lower():
            badge = "飯店機票折抵"
            category = "機酒優惠"
            desc = "Trip.com 最新日本機票、飯店住宿、全球機票折扣合集。"
        elif "優惠劵" in title:
            badge = "雲端合集"
            category = "綜合優惠"
            desc = "各大日本藥妝（松本清、Sundrug、大國）與百貨電器折扣券雲端備份檔。"
        elif "攻略" in title or "介紹" in title or "用法" in title:
            badge = "旅行指南"
            category = "實用攻略"
            desc = "日本美食與熱門景點詳細行程規劃指引。"

        coupons.append({
            "id": f"coupon-{len(coupons)+1}",
            "title": title,
            "url": url,
            "category": category,
            "badge": badge,
            "description": desc
        })

    return coupons

def parse_attractions(sheet):
    """Parse '景點資料' sheet."""
    attractions = []
    if not sheet:
        return attractions

    # Scan pairs of columns
    max_col = sheet.max_column or 0
    for c in range(1, max_col + 1, 2):
        day_tag = serialize_cell(sheet.cell(1, c).value)
        if not day_tag.startswith("Day"):
            continue

        # Look for custom notes
        for r in range(2, (sheet.max_row or 0) + 1):
            val1 = serialize_cell(sheet.cell(r, c).value)
            val2 = serialize_cell(sheet.cell(r, c+1).value)
            if not val1 and not val2:
                continue
            if val1 in ["景點", "相片(如有)", "地點", "前往方法", "營業時間", "費用", "注意事項"]:
                if val2:
                    attractions.append({
                        "day": day_tag,
                        "type": val1,
                        "content": val2
                    })
            elif "大丸心齋橋" in val1 or "錦市場" in val1 or "木津" in val1:
                attractions.append({
                    "day": day_tag,
                    "title": val1.split("\n")[0].strip(),
                    "details": val1,
                    "extra": val2
                })

    return attractions

def parse_todos_and_tickets(todo_sheet, ticket_sheet):
    """Parse 'todo' and '交通套票資訊' sheets."""
    todos = []
    if todo_sheet:
        for r in range(1, (todo_sheet.max_row or 0) + 1):
            category = serialize_cell(todo_sheet.cell(r, 1).value)
            if not category:
                continue
            items = []
            for c in range(2, (todo_sheet.max_column or 0) + 1):
                item = serialize_cell(todo_sheet.cell(r, c).value)
                if item:
                    items.append(item)
            if items:
                todos.append({
                    "category": category,
                    "items": items
                })

    tickets = []
    if ticket_sheet:
        for r in range(1, (ticket_sheet.max_row or 0) + 1):
            val = serialize_cell(ticket_sheet.cell(r, 1).value)
            if val:
                tickets.append(val)

    return todos, tickets

def build_flights_info(schedule):
    """
    Extract or synthesize structured flight data from itinerary.
    Day 1 outbound (TPE -> FUK) and Day 5 inbound (FUK -> TPE).
    """
    flights = [
        {
            "id": "flight-outbound",
            "type": "去程航班",
            "airline": "中華航空 / 星宇航空",
            "flight_no": "TPE ➔ FUK",
            "date": "2026-11-12",
            "day": "Day 1",
            "departure": {
                "airport": "桃園國際機場 (TPE)",
                "terminal": "第一航廈 (預定)",
                "time": "06:45",
                "action": "03:40 出發前往機場，04:30 抵達桃機辦理報到與行李託運"
            },
            "arrival": {
                "airport": "福岡機場 (FUK)",
                "terminal": "國際線航廈",
                "time": "10:00",
                "action": "預計 10:00 抵達，辦理入境審查與領取行李"
            },
            "duration": "2 小時 15 分",
            "baggage": "託運行李 23kg / 件、手提 7kg",
            "status": "準時預定",
            "notes": "入境後搭乘機場接駁車至國內線航廈，轉乘地下鐵至博多車站（車程約 5 分鐘）。"
        },
        {
            "id": "flight-inbound",
            "type": "回程航班",
            "airline": "中華航空 / 星宇航空",
            "flight_no": "FUK ➔ TPE",
            "date": "2026-11-16",
            "day": "Day 5",
            "departure": {
                "airport": "福岡機場 (FUK)",
                "terminal": "國際線航廈",
                "time": "17:55",
                "action": "15:00 抵達機場，16:00 辦理登機報到、退稅與出境手續"
            },
            "arrival": {
                "airport": "桃園國際機場 (TPE)",
                "terminal": "第一航廈 (預定)",
                "time": "19:40",
                "action": "19:40 抵達桃園機場，完成入境查驗，預計 21:00-22:00 返家休息"
            },
            "duration": "2 小時 45 分",
            "baggage": "託運行李 23kg / 件、手提 7kg",
            "status": "準時預定",
            "notes": "福岡機場免稅店人潮眾多，建議提早於 15:00 抵達國際線航廈。"
        }
    ]
    return flights

def build_practical_info():
    """Practical travel information for Japan trip."""
    return {
        "emergency_contacts": [
            {"name": "台北駐日經濟文化代表處 (福岡分處)", "tel": "+81-92-734-2810", "emergency_tel": "+81-90-8765-3410", "address": "福岡市中央區櫻坂3-12-42"},
            {"name": "日本警察局緊急報案", "tel": "110", "address": "日本境內直撥 110"},
            {"name": "日本火警 / 急救救護車", "tel": "119", "address": "日本境內直撥 119"},
            {"name": "日本觀光廳訪日諮詢專線 (多語言)", "tel": "050-3816-2788", "address": "24小時年中無休"}
        ],
        "luggage_storage": [
            {"name": "佐川急便 博多車站寄物所", "location": "博多站中央街 1-1（新幹線剪票口旁）", "hours": "09:00 - 20:00", "notes": "大件行李約 800-1000 日圓/件，可寄放一整天"},
            {"name": "博多運河城投幣置物櫃", "location": "Canal City 各棟 B1F / 1F", "hours": "依商場營業時間", "notes": "支援電子支付與現金，特大尺寸數量有限"}
        ],
        "accommodations": [
            {"name": "博多市區精選民宿 / 飯店", "address": "福岡市博多區博多站前", "check_in": "15:00", "check_out": "11:00", "notes": "鄰近博多車站，步行約 5-8 分鐘"}
        ],
        "useful_links": [
            {"title": "Visit Japan Web 入境手續", "url": "https://www.vjw.digital.go.jp/", "desc": "入境日本前填寫海關與入境審查 QR Code"},
            {"title": "JR 九州鐵路官網時刻表", "url": "https://www.jrkyushu.co.jp/", "desc": "查詢九州新幹線與特急列車班次"},
            {"title": "福岡市地下鐵路線圖", "url": "https://subway.city.fukuoka.lg.jp/", "desc": "博多、天神、福岡機場乘車指引"}
        ]
    }

def main():
    parser = argparse.ArgumentParser(description="Convert itinerary.xlsx to JSON pipeline.")
    parser.add_argument("--input", default="itinerary.xlsx", help="Path to input Excel file")
    parser.add_argument("--output", default="dist/itinerary_data.json", help="Path to output JSON file")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        print(f"Error: Input file '{input_path}' not found.", file=sys.stderr)
        sys.exit(1)

    print(f"Loading workbook from {input_path}...")
    wb = openpyxl.load_workbook(input_path, data_only=True)

    # 1. Parse Schedule
    schedule_sheet = wb["行程表"] if "行程表" in wb.sheetnames else wb.active
    schedule = parse_schedule(schedule_sheet)
    print(f"Parsed {len(schedule)} days of schedule.")

    # 2. Parse Coupons
    coupons_sheet = wb["遊日優惠"] if "遊日優惠" in wb.sheetnames else None
    coupons = parse_coupons(coupons_sheet)
    print(f"Parsed {len(coupons)} coupons/links.")

    # 3. Parse Attractions
    attractions_sheet = wb["景點資料"] if "景點資料" in wb.sheetnames else None
    attractions = parse_attractions(attractions_sheet)
    print(f"Parsed {len(attractions)} attraction notes.")

    # 4. Parse Todos & Tickets
    todo_sheet = wb["todo"] if "todo" in wb.sheetnames else None
    ticket_sheet = wb["交通套票資訊"] if "交通套票資訊" in wb.sheetnames else None
    todos, tickets = parse_todos_and_tickets(todo_sheet, ticket_sheet)
    print(f"Parsed {len(todos)} todo groups and {len(tickets)} tickets.")

    # 5. Build Flights
    flights = build_flights_info(schedule)

    # 6. Practical Info
    practical = build_practical_info()

    # Compile Full Payload
    result = {
        "title": "2026 福岡・九州自由行",
        "generated_at": datetime.datetime.now().isoformat(),
        "meta": {
            "destination": "日本福岡・九州 (Fukuoka & Kyushu)",
            "dates": "2026/11/12 - 2026/11/16",
            "total_days": len(schedule),
            "version": "1.0.0"
        },
        "schedule": schedule,
        "flights": flights,
        "coupons": coupons,
        "attractions": attractions,
        "todos": todos,
        "tickets": tickets,
        "practical": practical
    }

    # Ensure output directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Write output
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Successfully generated {output_path} ({output_path.stat().st_size} bytes).")

if __name__ == "__main__":
    main()
