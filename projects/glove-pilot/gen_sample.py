"""ساخت داده آزمایشی (ساختگی) برای اجرای آزمایشی اسکریپت تطبیق.

همه اعداد ساختگی هستند و فقط برای تست مدل استفاده می‌شوند. چند انحراف عمدی
در داده گذاشته شده تا معلوم شود اسکریپت آن‌ها را پیدا می‌کند:
  - کسری در انبار مرکزی برای GLV-NIT-M
  - مصرف بالای وینیل در جراحی با ثبت پایین روی پرونده
  - ثبت ناقص مصرف در اورژانس
  - کسورات بیمه به دلیل نبود مستند مصرف
  - یک انتقال بین بخش (داخلی به اورژانس) و یک ضایعات ثبت‌شده در جراحی
  - گاز: مصرف جراحی بالاتر از معیار و کسری ۴ بسته در انبار مرکزی
"""
import csv
import random
from pathlib import Path

random.seed(1405)
OUT = Path(__file__).parent / "sample_data"
OUT.mkdir(exist_ok=True)
PERIOD = "1405/07"
DAYS = 30

wards = {"ICU": 290, "INT": 810, "SURG": 640, "ER": 1350}  # فعالیت ماه
hoteling = {"ICU": 4_350_000_000, "INT": 6_480_000_000, "SURG": 5_120_000_000, "ER": 2_700_000_000}
# انتقال بین بخش و ضایعات ثبت‌شده: (بخش، قلم) -> تغییر موجودی پایان
transfers = [["1405/07/12", "TR-1", "GLV-VIN-M", "INT", "ER", 2, "بسته"]]
waste = [["1405/07/18", "ZA-1", "SURG", "GLV-NIT-M", 100, "عدد", "پارگی بسته در اثر نشت مایع"]]
adj = {("INT", "GLV-VIN-M"): -200, ("ER", "GLV-VIN-M"): 200, ("SURG", "GLV-NIT-M"): -100}
# مصرف واقعی به ازای فعالیت (عدد) و نسبت ثبت روی پرونده
true_rate = {
    ("ICU", "nitrile"): (42, 0.85), ("INT", "nitrile"): (15, 0.80),
    ("SURG", "nitrile"): (12, 0.78), ("ER", "nitrile"): (6.5, 0.45),
    ("ICU", "vinyl"): (8, 0.80), ("INT", "vinyl"): (6, 0.75),
    ("SURG", "vinyl"): (11, 0.30), ("ER", "vinyl"): (4, 0.40),
}
items = {"GLV-NIT-M": ("nitrile", 0.6), "GLV-NIT-L": ("nitrile", 0.4),
         "GLV-VIN-M": ("vinyl", 0.55), "GLV-VIN-L": ("vinyl", 0.45)}
price = {"nitrile": 24000, "vinyl": 14000}
insurers = [("تامین اجتماعی", 0.45), ("سلامت", 0.25), ("نیروهای مسلح", 0.10),
            ("تکمیلی", 0.10), ("آزاد", 0.10)]


def d(day):
    return f"{PERIOD}/{day:02d}"


def write(name, header, rows):
    with open(OUT / name, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


receipts, issues, returns, stock, charges, claims = [], [], [], [], [], []
central_issue_total = {i: 0 for i in items}

for item, (grp, share) in items.items():
    central_open = random.randint(30, 60) * 100
    for ward, act in wards.items():
        rate, capture = true_rate[(ward, grp)]
        used = round(rate * act * share)
        opening = random.randint(5, 15) * 100
        issued_boxes = -(-(used + random.randint(2, 6) * 100 - opening) // 100)
        issued_boxes = max(issued_boxes, 0)
        returned = 100 if ward == "INT" and grp == "vinyl" else 0
        closing = opening + issued_boxes * 100 - returned - used + adj.get((ward, item), 0)
        stock += [[d(1), ward, item, opening, "opening", "شمارنده نمونه"],
                  [d(DAYS), ward, item, closing, "closing", "شمارنده نمونه"]]
        # تقسیم حواله‌ها در طول ماه
        remaining = issued_boxes
        n = 1
        while remaining > 0:
            q = min(remaining, random.randint(3, 12))
            issues.append([d(random.randint(1, DAYS)), f"HV-{ward}-{item[-5:]}-{n}", item, ward, q, "بسته"])
            remaining -= q
            n += 1
        central_issue_total[item] += issued_boxes * 100
        if returned:
            returns.append([d(20), f"BR-{ward}-{item[-5:]}", item, ward, returned, "عدد"])
            central_issue_total[item] -= returned
        # ثبت روی پرونده بیمار
        to_charge = round(used * capture)
        k = 0
        while to_charge > 0:
            q = min(to_charge, random.choice([2, 4, 6, 10, 20]))
            ins = random.choices([x[0] for x in insurers], [x[1] for x in insurers])[0]
            line = f"CH-{ward}-{item[-5:]}-{k}"
            adm = f"ADM{random.randint(10000, 99999)}"
            tariff = 28000 if grp == "nitrile" else 17000
            charges.append([d(random.randint(1, DAYS)), line, adm, ward, item, q, "عدد", tariff, ins])
            if ins != "آزاد":
                ded = 0
                reason = ""
                r = random.random()
                if r < (0.25 if ward == "ER" else 0.08):
                    ded, reason = q, "نبود دستور یا مستند مصرف در پرونده"
                elif r < (0.30 if ward == "ER" else 0.12):
                    ded, reason = q // 2, "مازاد بر سقف مجاز"
                claims.append([line, adm, ins, q * tariff, ded * tariff, reason])
            to_charge -= q
            k += 1

    # انبار مرکزی: دریافت از تامین‌کننده و شمارش
    received = -(-(central_issue_total[item] + 2000 - central_open) // 1000) * 1000
    unit_price = price[grp] + random.randint(-1000, 1000)
    cartons = received // 1000
    for j, part in enumerate([cartons // 2, cartons - cartons // 2]):
        receipts.append([d(3 + j * 14), f"INV-{item[-5:]}-{j}", "تامین‌کننده نمونه", item, part, "کارتن", unit_price])
    shrink = 600 if item == "GLV-NIT-M" else 0  # کسری عمدی
    central_close = central_open + received - central_issue_total[item] - shrink
    stock += [[d(1), "CENTRAL", item, central_open, "opening", "انباردار نمونه"],
              [d(DAYS), "CENTRAL", item, central_close, "closing", "انباردار نمونه"]]

# ---------- گاز طبی (واحد پایه: بسته نیم‌کیلویی؛ روی پرونده ثبت نمی‌شود) ----------
gauze_rate = {"ICU": 0.095, "INT": 0.04, "SURG": 0.09, "ER": 0.02}
g_central_open, g_issued = 60, 0
for ward, act in wards.items():
    used = round(gauze_rate[ward] * act)
    opening = random.randint(3, 8)
    issued = max(used + random.randint(2, 5) - opening, 0)
    closing = opening + issued - used
    stock += [[d(1), ward, "GAU-500", opening, "opening", "شمارنده نمونه"],
              [d(DAYS), ward, "GAU-500", closing, "closing", "شمارنده نمونه"]]
    half = issued // 2
    for n, q in enumerate([half, issued - half], 1):
        if q:
            issues.append([d(5 if n == 1 else 20), f"HV-{ward}-GAU-{n}", "GAU-500", ward, q, "بسته"])
    g_issued += issued
g_received = -(-(g_issued + 20 - g_central_open) // 20) * 20
receipts.append([d(4), "INV-GAU-0", "تامین‌کننده نمونه", "GAU-500", g_received // 20, "کارتن", 9900000])
stock += [[d(1), "CENTRAL", "GAU-500", g_central_open, "opening", "انباردار نمونه"],
          [d(DAYS), "CENTRAL", "GAU-500", g_central_open + g_received - g_issued - 4, "closing", "انباردار نمونه"]]

write("receipts.csv", ["date", "doc_no", "supplier", "item_code", "qty", "unit", "unit_price_per_base_rial"], receipts)
write("issues.csv", ["date", "doc_no", "item_code", "ward", "qty", "unit"], issues)
write("returns.csv", ["date", "doc_no", "item_code", "ward", "qty", "unit"], returns)
write("stocktake.csv", ["date", "location", "item_code", "qty_base", "count_point", "counted_by"], stock)
write("charges.csv", ["date", "line_id", "admission_id", "ward", "item_code", "qty", "unit", "tariff_price_per_base_rial", "insurer"], charges)
write("claims.csv", ["line_id", "admission_id", "insurer", "billed_rial", "deducted_rial", "deduction_reason"], claims)
write("transfers.csv", ["date", "doc_no", "item_code", "from_ward", "to_ward", "qty", "unit"], transfers)
write("waste.csv", ["date", "doc_no", "location", "item_code", "qty", "unit", "reason"], waste)
write("activity.csv", ["period", "ward", "activity", "hoteling_revenue_rial"],
      [[PERIOD, w, a, hoteling[w]] for w, a in wards.items()])
print("sample data written to", OUT)
