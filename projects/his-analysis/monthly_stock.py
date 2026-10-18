"""تحلیل ماهانه گزارش‌های استوک بخش‌ها (خروج بدون صورتحساب) از HIS.

هر فایل ورودی گزارش اقلام بخش‌ها برای یک ماه است (همان قالب «طراح داده پیشرو»)
که سهم بیمه و بیمار آن صفر است، یعنی مصرف استوک بخش.

نمونه:
    python monthly_stock.py --wards ward_indicators.xlsx --out output \
        --file فروردین=far.xlsx --file اردیبهشت=ord.xlsx ...

ترتیب --fileها همان ترتیب ماه‌ها در خروجی است.
"""
import argparse
from pathlib import Path

import pandas as pd

from analyze_his import WARD_MAP, glove_type, load_beddays

EXAM = ["معاینه وینیل", "معاینه نیتریل", "معاینه لاتکس"]


def load_month(path):
    d = pd.read_excel(path)
    d = d[pd.to_numeric(d["كد"], errors="coerce").notna()].copy()  # ردیف جمع حذف
    d["كد"] = d["كد"].astype(int)
    d["تعداد"] = pd.to_numeric(d["تعداد"].astype(str).str.replace(",", ""), errors="coerce")
    for c in ["جمع خريد", "في خريد", "جمع كل فروش", "سهم بيمه", "سهم بيمار"]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0)
    d["بخش"] = d["بخش"].astype(str).str.strip()
    d["نوع_دستکش"] = (d["نام فارسي"].fillna("").astype(str) + " " + d["نام"].fillna("").astype(str)).map(glove_type)
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", action="append", required=True, help="ماه=مسیر فایل")
    ap.add_argument("--wards", help="فایل عملکرد بخش‌ها (تخت‌روز ماهانه)")
    ap.add_argument("--out", default=str(Path(__file__).parent / "output"))
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)

    months, frames = [], []
    for spec in a.file:
        m, p = spec.split("=", 1)
        d = load_month(p)
        d["ماه"] = m
        months.append(m)
        frames.append(d)
    df = pd.concat(frames, ignore_index=True)
    order = pd.CategoricalDtype(months, ordered=True)
    df["ماه"] = df["ماه"].astype(order)

    # کنترل: این گزارش‌ها باید بدون سهم بیمه و بیمار باشند
    billed = df[(df["سهم بيمه"] != 0) | (df["سهم بيمار"] != 0)]

    g = df[df["نوع_دستکش"].isin(EXAM)]
    gm = g.pivot_table(index="بخش", columns="ماه", values="تعداد", aggfunc="sum", fill_value=0, observed=False)
    gm["جمع"] = gm.sum(axis=1)
    gm = gm.sort_values("جمع", ascending=False)
    gm.loc["جمع کل"] = gm.sum()

    gt = g.pivot_table(index="نوع_دستکش", columns="ماه", values="تعداد", aggfunc="sum", fill_value=0, observed=False)
    gt.loc["جمع"] = gt.sum()

    price = (g.groupby(["ماه", "نوع_دستکش"], observed=True)
             .agg(تعداد=("تعداد", "sum"), ارزش=("جمع خريد", "sum")).reset_index())
    price["میانگین_قیمت_خرید"] = (price["ارزش"] / price["تعداد"]).round(0)

    val = df.pivot_table(index="بخش", columns="ماه", values="جمع خريد", aggfunc="sum", fill_value=0, observed=False)
    val["جمع"] = val.sum(axis=1)
    val = (val / 1e6).round(0).sort_values("جمع", ascending=False)
    val.loc["جمع کل"] = val.sum()

    top = (df.groupby(["كد", "نام"]).agg(تعداد=("تعداد", "sum"), ارزش_خرید=("جمع خريد", "sum"))
           .sort_values("ارزش_خرید", ascending=False).head(40).reset_index())
    top["ارزش_خرید"] = (top["ارزش_خرید"] / 1e6).round(0)

    sheets = {"دستکش_بخش_ماه": gm, "دستکش_نوع_ماه": gt, "قیمت_دستکش": price,
              "ارزش_استوک_بخش_میلیون": val, "۴۰_قلم_پرهزینه": top}

    if a.wards:
        bd, _ = load_beddays(a.wards)
        bd["month"] = bd["month"].str.strip().str.replace("ي", "ی")
        bd = bd.dropna(subset=["bed_days"])
        rows = []
        for (ward, m), q in g.groupby(["بخش", "ماه"], observed=True)["تعداد"].sum().items():
            key = WARD_MAP.get(ward)
            hit = bd[(bd["ward"] == key) & (bd["month"] == str(m).replace("ي", "ی"))]
            if key and len(hit):
                rows.append({"بخش": ward, "ماه": m, "دستکش": q, "تخت_روز": hit["bed_days"].iloc[0]})
        if rows:
            r = pd.DataFrame(rows)
            r["ماه"] = r["ماه"].astype(order)
            r["به_ازای_تخت‌روز"] = (r["دستکش"] / r["تخت_روز"]).round(1)
            sheets["دستکش_به_ازای_تخت‌روز"] = r.pivot_table(
                index="بخش", columns="ماه", values="به_ازای_تخت‌روز", observed=False)

    sheets["کنترل"] = pd.DataFrame({
        "شاخص": ["تعداد ماه", "ردیف", "ردیف دارای سهم بیمه یا بیمار (باید صفر باشد)"],
        "مقدار": [len(months), len(df), len(billed)]})

    path = out / "monthly_stock.xlsx"
    with pd.ExcelWriter(path) as w:
        for name, t in sheets.items():
            t.to_excel(w, sheet_name=name[:31])
    print(gm.to_string())
    print(gt.to_string())
    if "دستکش_به_ازای_تخت‌روز" in sheets:
        print(sheets["دستکش_به_ازای_تخت‌روز"].to_string())
    print(price.to_string())
    print(val.head(12).to_string())
    print("->", path)


if __name__ == "__main__":
    main()
