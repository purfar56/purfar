"""تطبیق ورود، تحویل، مصرف، ثبت روی پرونده و بیمه برای اقلام مصرفی.

اجرا:
    python reconcile.py --data sample_data --period 1405/07              # ماهانه
    python reconcile.py --data sample_data --period 1405/07/01-1405/07/07  # هفتگی یا هر بازه

ورودی‌ها (CSV با UTF-8) در data_dictionary.md تعریف شده‌اند.
خروجی: فایل Excel با جزئیات و گزارش مدیریتی یک‌صفحه‌ای در پوشه output.
"""
import argparse
from pathlib import Path

import pandas as pd

HERE = Path(__file__).parent
GROUP_FA = {"nitrile": "دستکش نیتریل", "vinyl": "دستکش وینیل", "gauze": "گاز طبی"}
UNIT_FA = {"gauze": "بسته"}  # بقیه گروه‌ها: عدد


def load(path, name, required=True):
    f = path / name
    if not f.exists():
        if required:
            raise SystemExit(f"فایل لازم پیدا نشد: {f}")
        return pd.DataFrame()
    return pd.read_csv(f, dtype={"item_code": str, "ward": str}, encoding="utf-8")


def in_period(df, period):
    """دوره یا یک ماه است (1405/07) یا یک بازه (1405/07/01-1405/07/07)."""
    if df.empty or "date" not in df:
        return df
    d = df["date"].astype(str)
    if "-" in period:
        start, end = period.split("-")
        return df[(d >= start) & (d <= end)].copy()
    return df[d.str.startswith(period)].copy()


def jalali_ordinal(date):
    y, m, d = (int(x) for x in date.split("/"))
    return y * 366 + (m - 1) * 31 - max(m - 7, 0) + d  # ماه‌های ۷ به بعد ۳۰ روزه‌اند


def period_days(period):
    if "-" in period:
        start, end = period.split("-")
        return jalali_ordinal(end) - jalali_ordinal(start) + 1
    m = int(period.split("/")[1])
    return 31 if m <= 6 else 30 if m <= 11 else 29


def to_base(df, units, issues):
    """تبدیل مقدار به واحد پایه (عدد) با جدول units؛ ردیف‌های بدون ضریب گزارش می‌شوند."""
    if df.empty:
        df["qty_base"] = []
        return df
    specific = units[units["item_code"] != "*"]
    generic = units[units["item_code"] == "*"][["unit", "factor_to_base"]]
    df = df.merge(specific, on=["item_code", "unit"], how="left")
    df = df.merge(generic, on="unit", how="left", suffixes=("", "_g"))
    df["factor"] = df["factor_to_base"].fillna(df["factor_to_base_g"])
    for _, r in df[df["factor"].isna()].iterrows():
        issues.append(("بحرانی", "واحد بدون ضریب تبدیل", f"{r['item_code']} / واحد «{r['unit']}»"))
    df["qty_base"] = df["qty"] * df["factor"]
    return df.drop(columns=["factor_to_base", "factor_to_base_g", "factor"])


def stock_at(stock, location, point):
    s = stock[(stock["location"] == location) & (stock["count_point"] == point)]
    return s.groupby("item_code")["qty_base"].sum()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default="sample_data")
    ap.add_argument("--config", default=str(HERE / "config"))
    ap.add_argument("--period", required=True, help="مثل 1405/07 یا 1405/07/01-1405/07/07")
    ap.add_argument("--days", type=int, help="تعداد روز دوره (پیش‌فرض: محاسبه از دوره)")
    ap.add_argument("--out", default=str(HERE / "output"))
    a = ap.parse_args()

    data, cfg, out = Path(a.data), Path(a.config), Path(a.out)
    if not data.is_absolute() and not data.exists():
        data = HERE / data
    out.mkdir(parents=True, exist_ok=True)
    p = a.period
    dq = []  # مشکلات کیفیت داده: (شدت، نوع، جزئیات)

    items = load(cfg, "items.csv")
    units = load(cfg, "units.csv")
    billing = load(cfg, "billing_rules.csv")
    norms = load(cfg, "norms.csv", required=False)
    wards = load(cfg, "wards.csv")

    receipts = to_base(in_period(load(data, "receipts.csv"), p), units, dq)
    issues = to_base(in_period(load(data, "issues.csv"), p), units, dq)
    returns = to_base(in_period(load(data, "returns.csv", required=False), p), units, dq)
    charges = to_base(in_period(load(data, "charges.csv"), p), units, dq)
    transfers = to_base(in_period(load(data, "transfers.csv", required=False), p), units, dq)
    waste = to_base(in_period(load(data, "waste.csv", required=False), p), units, dq)
    days = a.days or period_days(p)
    claims = load(data, "claims.csv", required=False)
    stock = in_period(load(data, "stocktake.csv"), p)
    activity = load(data, "activity.csv")
    activity = activity[activity["period"] == p]
    if "hoteling_revenue_rial" not in activity:
        activity = activity.assign(hoteling_revenue_rial=pd.NA)

    known = set(items["item_code"])
    for name, df in [("ورود", receipts), ("حواله", issues), ("ثبت پرونده", charges)]:
        for code in sorted(set(df.get("item_code", [])) - known):
            dq.append(("عمده", f"کد کالای ناشناخته در {name}", code))

    # قیمت خرید: میانگین موزون دوره؛ در نبود خرید، از تعرفه به‌عنوان جایگزین
    if not receipts.empty:
        receipts["value"] = receipts["qty_base"] * receipts["unit_price_per_base_rial"]
        g = receipts.groupby("item_code")[["value", "qty_base"]].sum()
        cost_price = g["value"] / g["qty_base"]
    else:
        cost_price = pd.Series(dtype=float)
    tariff = billing[billing["insurer"] == "*"].set_index("item_code")["tariff_price_per_base_rial"]
    billable = billing[billing["insurer"] == "*"].set_index("item_code")["billable"].eq("yes")
    for code in items["item_code"]:
        if code not in cost_price.index:
            dq.append(("عمده", "قیمت خرید در دوره نیست؛ از تعرفه استفاده شد", code))
            cost_price[code] = tariff.get(code, 0)

    # ---------- انبار مرکزی ----------
    c_open = stock_at(stock, "CENTRAL", "opening")
    c_close = stock_at(stock, "CENTRAL", "closing")
    rec = receipts.groupby("item_code")["qty_base"].sum()
    iss = issues.groupby("item_code")["qty_base"].sum()
    ret = returns.groupby("item_code")["qty_base"].sum() if not returns.empty else pd.Series(dtype=float)
    central = pd.DataFrame(index=items["item_code"])
    central["موجودی_اول"] = c_open
    central["ورود"] = rec
    central["حواله_به_بخش"] = iss
    central["برگشت_از_بخش"] = ret
    central["ضایعات_ثبت‌شده"] = sum_by(waste[waste["location"] == "CENTRAL"] if not waste.empty else waste)
    central["موجودی_پایان_شمارش"] = c_close
    central = central.fillna(0)
    central["موجودی_پایان_دفتری"] = (central["موجودی_اول"] + central["ورود"] - central["حواله_به_بخش"]
                                     + central["برگشت_از_بخش"] - central["ضایعات_ثبت‌شده"])
    central["کسری(+)/اضافه(-)"] = central["موجودی_پایان_دفتری"] - central["موجودی_پایان_شمارش"]
    central["ریال_کسری"] = central["کسری(+)/اضافه(-)"] * cost_price.reindex(central.index)
    central.index.name = "item_code"

    # ---------- بخش × قلم ----------
    rows = []
    for ward in wards["ward"]:
        o = stock_at(stock, ward, "opening")
        c = stock_at(stock, ward, "closing")
        wi = issues[issues["ward"] == ward].groupby("item_code")["qty_base"].sum()
        wr = returns[returns["ward"] == ward].groupby("item_code")["qty_base"].sum() if not returns.empty else {}
        wc = charges[charges["ward"] == ward].groupby("item_code")["qty_base"].sum()
        t_in = sum_by(transfers[transfers["to_ward"] == ward] if not transfers.empty else transfers)
        t_out = sum_by(transfers[transfers["from_ward"] == ward] if not transfers.empty else transfers)
        ww = sum_by(waste[waste["location"] == ward] if not waste.empty else waste)
        for code in items["item_code"]:
            if code not in o.index or code not in c.index:
                if wi.get(code, 0) or wc.get(code, 0):
                    dq.append(("بحرانی", "شمارش اول یا پایان دوره ثبت نشده", f"{ward} / {code}"))
            # مصرف واقعی شامل کسری نامشخص بخش هم هست؛ بدون شمارش در سطح بیمار قابل جداسازی نیست
            used = (o.get(code, 0) + wi.get(code, 0) + t_in.get(code, 0) - wr.get(code, 0)
                    - t_out.get(code, 0) - ww.get(code, 0) - c.get(code, 0))
            rows.append({"ward": ward, "item_code": code, "موجودی_اول": o.get(code, 0),
                         "تحویل": wi.get(code, 0), "انتقال_ورودی": t_in.get(code, 0),
                         "برگشت": wr.get(code, 0), "انتقال_خروجی": t_out.get(code, 0),
                         "ضایعات_ثبت‌شده": ww.get(code, 0),
                         "موجودی_پایان": c.get(code, 0), "مصرف_واقعی": used,
                         "ثبت_روی_پرونده": wc.get(code, 0)})
    wd = pd.DataFrame(rows).merge(items[["item_code", "item_group"]], on="item_code")
    for _, r in wd[wd["مصرف_واقعی"] < 0].iterrows():
        dq.append(("بحرانی", "مصرف منفی (شمارش یا حواله اشتباه)", f"{r['ward']} / {r['item_code']}"))
    for _, r in wd[(wd["مصرف_واقعی"] == 0) & (wd["ثبت_روی_پرونده"] > 0)].iterrows():
        dq.append(("عمده", "ثبت روی پرونده بدون مصرف", f"{r['ward']} / {r['item_code']}"))

    wd["قیمت_خرید"] = wd["item_code"].map(cost_price)
    wd["هزینه_مصرف_ریال"] = wd["مصرف_واقعی"] * wd["قیمت_خرید"]
    wd["ضایعات_ریال"] = wd["ضایعات_ثبت‌شده"] * wd["قیمت_خرید"]
    wd["قابل_اخذ"] = wd["item_code"].map(billable).fillna(False)
    wd["نرخ_ثبت_درصد"] = (wd["ثبت_روی_پرونده"] / wd["مصرف_واقعی"].where(wd["مصرف_واقعی"] > 0) * 100).round(1)
    gap = (wd["مصرف_واقعی"] - wd["ثبت_روی_پرونده"]).clip(lower=0)
    wd["ثبت‌نشده_تعداد"] = gap.where(wd["قابل_اخذ"], 0)
    wd["درآمد_ازدست‌رفته_ریال"] = wd["ثبت‌نشده_تعداد"] * wd["item_code"].map(tariff).fillna(0)

    # ---------- مقایسه با معیار مصرف (در سطح گروه کالا) ----------
    grp = wd.groupby(["ward", "item_group"], as_index=False)[["مصرف_واقعی", "هزینه_مصرف_ریال"]].sum()
    grp = grp.merge(activity[["ward", "activity"]], on="ward", how="left")
    grp = grp.merge(wards[["ward", "ward_name", "activity_type"]], on="ward", how="left")
    for w in grp[grp["activity"].isna()]["ward"].unique():
        dq.append(("بحرانی", "تخت‌روز/ویزیت دوره ثبت نشده", w))
    grp["مصرف_به_ازای_فعالیت"] = (grp["مصرف_واقعی"] / grp["activity"]).round(3)
    if not norms.empty:
        grp = grp.merge(norms[["ward", "item_group", "norm_per_activity", "source"]],
                        on=["ward", "item_group"], how="left")
    else:
        grp["norm_per_activity"], grp["source"] = pd.NA, pd.NA
    grp["مصرف_مورد_انتظار"] = grp["norm_per_activity"] * grp["activity"]
    grp["مازاد_بر_معیار_تعداد"] = (grp["مصرف_واقعی"] - grp["مصرف_مورد_انتظار"]).clip(lower=0)
    avg_price = (grp["هزینه_مصرف_ریال"] / grp["مصرف_واقعی"]).fillna(0)
    grp["مازاد_بر_معیار_ریال"] = grp["مازاد_بر_معیار_تعداد"] * avg_price
    hot = activity.set_index("ward")["hoteling_revenue_rial"]
    grp["درآمد_هتلینگ_ریال"] = grp["ward"].map(hot)
    ward_cost = grp.groupby("ward")["هزینه_مصرف_ریال"].transform("sum")
    grp["سهم_هزینه_دستکش_از_هتلینگ_درصد"] = (ward_cost / pd.to_numeric(grp["درآمد_هتلینگ_ریال"])
                                               * 100).round(2)

    # ---------- استوک پویا: ROP = d × L + SS ----------
    st = wd[["ward", "item_code", "مصرف_واقعی", "موجودی_پایان", "قیمت_خرید"]].merge(wards, on="ward")
    st["مصرف_روزانه"] = (st["مصرف_واقعی"] / days).round(1)
    st["موجودی_اطمینان"] = (st["مصرف_روزانه"] * st["safety_days"]).round()
    st["نقطه_سفارش"] = (st["مصرف_روزانه"] * st["lead_days"] + st["موجودی_اطمینان"]).round()
    st["حداکثر_موجودی"] = (st["مصرف_روزانه"] * (st["review_days"] + st["lead_days"])
                           + st["موجودی_اطمینان"]).round()
    st["روز_پوشش_موجودی_فعلی"] = (st["موجودی_پایان"] / st["مصرف_روزانه"].where(st["مصرف_روزانه"] > 0)).round(1)
    st["مازاد_بر_حداکثر"] = (st["موجودی_پایان"] - st["حداکثر_موجودی"]).clip(lower=0)
    st["مازاد_بر_حداکثر_ریال"] = st["مازاد_بر_حداکثر"] * st["قیمت_خرید"]
    st["پیشنهاد_درخواست_بعدی"] = (st["حداکثر_موجودی"] - st["موجودی_پایان"]).clip(lower=0)
    st = st.drop(columns=["ward_name", "activity_type"])

    # ---------- بیمه ----------
    if not claims.empty and not charges.empty:
        cl = claims.merge(charges[["line_id", "ward", "item_code"]], on="line_id", how="left")
        orphan = cl["ward"].isna().sum()
        if orphan:
            dq.append(("عمده", "سطر بیمه بدون سطر ثبت پرونده متناظر", f"{orphan} سطر"))
        ins = cl.groupby("insurer", as_index=False)[["billed_rial", "deducted_rial"]].sum()
        ins["نرخ_کسورات_درصد"] = (ins["deducted_rial"] / ins["billed_rial"] * 100).round(1)
        reasons = (cl[cl["deducted_rial"] > 0].groupby("deduction_reason", as_index=False)["deducted_rial"]
                   .sum().sort_values("deducted_rial", ascending=False))
        ins_ward = cl.groupby("ward", as_index=False)[["billed_rial", "deducted_rial"]].sum()
        ins_ward["نرخ_کسورات_درصد"] = (ins_ward["deducted_rial"] / ins_ward["billed_rial"] * 100).round(1)
    else:
        ins = reasons = ins_ward = pd.DataFrame()
        dq.append(("عمده", "داده بیمه/کسورات وارد نشده", "claims.csv"))

    # ---------- خلاصه ----------
    total_in = (rec * cost_price.reindex(rec.index)).sum()
    total_cost = wd["هزینه_مصرف_ریال"].sum()
    summary = pd.DataFrame([
        ("ارزش ورودی به انبار مرکزی", total_in),
        ("هزینه مصرف واقعی بخش‌ها", total_cost),
        ("الف) کسری انبار مرکزی", central["ریال_کسری"].clip(lower=0).sum()),
        ("ب) مصرف مازاد بر معیار (به قیمت خرید)", grp["مازاد_بر_معیار_ریال"].sum()),
        ("ج) مصرف ثبت‌نشده روی پرونده (به تعرفه)", wd["درآمد_ازدست‌رفته_ریال"].sum()),
        ("ضایعات ثبت‌شده بخش‌ها", wd["ضایعات_ریال"].sum()),
        ("موجودی بخش‌ها بیش از حداکثر پیشنهادی", st["مازاد_بر_حداکثر_ریال"].sum()),
        ("د) کسورات بیمه", ins["deducted_rial"].sum() if not ins.empty else 0),
        ("نرخ ثبت مصرف روی پرونده (درصد)",  # فقط اقلام قابل اخذ
         round(wd.loc[wd["قابل_اخذ"], "ثبت_روی_پرونده"].sum()
               / max(wd.loc[wd["قابل_اخذ"], "مصرف_واقعی"].sum(), 1) * 100, 1)),
    ], columns=["شاخص", "مقدار"])

    dqdf = pd.DataFrame(dq, columns=["شدت", "مشکل", "جزئیات"]).drop_duplicates()
    tag = p.replace("/", "-")
    xlsx = out / f"glove_pilot_{tag}.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="خلاصه", index=False)
        grp.to_excel(xw, sheet_name="بخش_گروه_معیار", index=False)
        wd.to_excel(xw, sheet_name="بخش_قلم", index=False)
        st.to_excel(xw, sheet_name="استوک_پویا", index=False)
        central.reset_index().to_excel(xw, sheet_name="انبار_مرکزی", index=False)
        if not ins.empty:
            ins.to_excel(xw, sheet_name="بیمه_بیمه‌گر", index=False)
            ins_ward.to_excel(xw, sheet_name="بیمه_بخش", index=False)
            reasons.to_excel(xw, sheet_name="علل_کسورات", index=False)
        dqdf.to_excel(xw, sheet_name="کیفیت_داده", index=False)
        for ws in xw.book.worksheets:
            ws.sheet_view.rightToLeft = True
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 22

    prev = update_history(out / "history.csv", p, summary)
    report = build_report(p, summary, grp, wd, central, ins_ward, reasons, dqdf, prev)
    md = out / f"report_{tag}.md"
    md.write_text(report, encoding="utf-8")
    print(f"خروجی: {xlsx}\nگزارش: {md}")


def sum_by(df):
    if df.empty:
        return pd.Series(dtype=float)
    return df.groupby("item_code")["qty_base"].sum()


def update_history(path, p, summary):
    """خلاصه دوره را در history.csv ثبت می‌کند و خلاصه دوره قبلی هم‌نوع (ماهانه/بازه) را برمی‌گرداند."""
    hist = pd.read_csv(path, dtype={"period": str}) if path.exists() else pd.DataFrame(
        columns=["period", "شاخص", "مقدار"])
    hist = hist[hist["period"] != p]
    same_kind = hist[hist["period"].str.contains("-") == ("-" in p)]
    earlier = sorted(x for x in same_kind["period"].unique() if x < p)
    prev = None
    if earlier:
        prev = dict(zip(*same_kind[same_kind["period"] == earlier[-1]][["شاخص", "مقدار"]].values.T))
        prev["_period"] = earlier[-1]
    new = summary.assign(period=p)[["period", "شاخص", "مقدار"]]
    pd.concat([hist, new]).sort_values("period").to_csv(path, index=False, encoding="utf-8")
    return prev


def fmt(x):
    return f"{x:,.0f}"


def build_report(p, summary, grp, wd, central, ins_ward, reasons, dqdf, prev=None):
    s = dict(zip(summary["شاخص"], summary["مقدار"]))
    prev = prev or {}

    def cmp(key, pct=False):
        if key not in prev:
            return "—"
        v = float(prev[key])
        return f"{v}٪" if pct else fmt(v)
    groups = "، ".join(GROUP_FA.get(g, g) for g in grp["item_group"].unique())
    L = [f"# گزارش مدیریتی پایلوت ردیابی اقلام ({groups}) — دوره {p}", "",
         "> این گزارش به‌صورت خودکار از داده‌ها ساخته شده است. بخش «تصمیم موردنیاز» را ناظر تکمیل می‌کند.", ""]
    crit = dqdf[dqdf["شدت"] == "بحرانی"]
    if len(crit):
        L += [f"**هشدار کیفیت داده:** {len(crit)} مشکل بحرانی در داده وجود دارد؛ "
              "اعداد زیر تا رفع آن‌ها قطعی نیستند (برگه «کیفیت_داده»).", ""]
    L += ["## وضعیت", "",
          f"| شاخص | این دوره | دوره قبل ({prev.get('_period', 'ندارد')}) |", "|---|---|---|",
          f"| هزینه مصرف واقعی (ریال) | {fmt(s['هزینه مصرف واقعی بخش‌ها'])} | {cmp('هزینه مصرف واقعی بخش‌ها')} |",
          f"| نرخ ثبت مصرف روی پرونده | {s['نرخ ثبت مصرف روی پرونده (درصد)']}٪ | "
          f"{cmp('نرخ ثبت مصرف روی پرونده (درصد)', pct=True)} |",
          f"| کسورات بیمه (ریال) | {fmt(s['د) کسورات بیمه'])} | {cmp('د) کسورات بیمه')} |", "",
          "## مصرف به ازای فعالیت", "",
          "| بخش | گروه | مصرف به ازای فعالیت | معیار (موقت) | هزینه (ریال) |", "|---|---|---|---|---|"]
    for _, r in grp.iterrows():
        L.append(f"| {r['ward_name']} | {GROUP_FA.get(r['item_group'], r['item_group'])} | {r['مصرف_به_ازای_فعالیت']} در هر {r['activity_type']} "
                 f"| {r['norm_per_activity']} | {fmt(r['هزینه_مصرف_ریال'])} |")
    L += ["",
          "## انحراف ریالی", "",
          "| نوع انحراف | ریال |", "|---|---|"]
    for k in ["الف) کسری انبار مرکزی", "ب) مصرف مازاد بر معیار (به قیمت خرید)",
              "ج) مصرف ثبت‌نشده روی پرونده (به تعرفه)", "د) کسورات بیمه",
              "ضایعات ثبت‌شده بخش‌ها", "موجودی بخش‌ها بیش از حداکثر پیشنهادی"]:
        L.append(f"| {k} | {fmt(s[k])} |")
    L += ["", "ردیف‌های ب و ج هم‌پوشانی دارند (مصرف مازاد معمولاً ثبت هم نمی‌شود) و نباید با هم جمع شوند. "
          "موجودی مازاد هزینه از دست‌رفته نیست، بلکه سرمایه راکد است.", "",
          "## خطرهای مهم (سه مورد بزرگ‌تر)", ""]
    risks = []
    for code, r in central[central["ریال_کسری"] > 0].iterrows():
        risks.append((r["ریال_کسری"], f"کسری انبار مرکزی در {code}: {fmt(r['کسری(+)/اضافه(-)'])} واحد"))
    for _, r in grp[grp["مازاد_بر_معیار_ریال"] > 0].iterrows():
        risks.append((r["مازاد_بر_معیار_ریال"],
                      f"مصرف {GROUP_FA.get(r['item_group'], r['item_group'])} در {r['ward_name']}: {r['مصرف_به_ازای_فعالیت']} {UNIT_FA.get(r['item_group'], 'عدد')} به ازای هر "
                      f"{r['activity_type']} در برابر معیار {r['norm_per_activity']}"))
    by_ward = wd[wd["قابل_اخذ"]].groupby("ward")[["مصرف_واقعی", "ثبت_روی_پرونده", "درآمد_ازدست‌رفته_ریال"]].sum()
    for w, r in by_ward.iterrows():
        rate = r["ثبت_روی_پرونده"] / max(r["مصرف_واقعی"], 1) * 100
        if rate < 70:
            risks.append((r["درآمد_ازدست‌رفته_ریال"], f"نرخ ثبت روی پرونده در {w} فقط {rate:.0f}٪"))
    if not ins_ward.empty:
        top = ins_ward.sort_values("نرخ_کسورات_درصد", ascending=False).iloc[0]
        risks.append((top["deducted_rial"], f"بیشترین نرخ کسورات در {top['ward']}: {top['نرخ_کسورات_درصد']}٪"))
    for i, (v, t) in enumerate(sorted(risks, reverse=True)[:3], 1):
        L.append(f"{i}. {t} — اثر ریالی تقریبی {fmt(v)} ریال")
    if not reasons.empty:
        r0 = reasons.iloc[0]
        L += ["", f"عمده‌ترین علت کسورات: «{r0['deduction_reason']}» ({fmt(r0['deducted_rial'])} ریال)."]
    L += ["", "## اقدام‌ها", "", "| اقدام | مسئول | موعد | وضعیت |", "|---|---|---|---|", "| | | | |", "",
          "## تصمیم موردنیاز", "", "1. ", ""]
    return "\n".join(L)


if __name__ == "__main__":
    main()
