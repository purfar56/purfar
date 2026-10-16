"""تحلیل گزارش اقلام بخش‌ها از HIS (طراح داده / پیشرو) همراه با تخت‌روز بخش‌ها.

اجرا:
    python analyze_his.py --his <گزارش اقلام HIS.xlsx> --wards <عملکرد بخش‌ها.xlsx> --label "1405 فروردین تا شهریور"

خروجی (پوشه output، خارج از git چون داده واقعی بیمارستان است):
    his_analysis.xlsx و his_report.md
"""
import argparse
import re
from pathlib import Path

import openpyxl
import pandas as pd

HERE = Path(__file__).parent

# نام بخش در گزارش اقلام HIS -> نام بخش در فایل عملکرد بخش‌ها
WARD_MAP = {"بخش 2": "2", "بخش 3A": "3A", "بخش 3B": "3B", "بخش 3C": "3C", "بخش 4A": "4A",
            "بخش 4B": "4B", "C.C.U": "CCU", "ICU": "ICU", "I.C.U OH": "OH", "NICU": "NICU",
            "بخش نوزادان": "نوزادان", "POST CATH": "post cath"}


# اقلام پایلوت غیر از دستکش: کد HIS -> نام کوتاه
PILOT_ITEMS = {118490: "گاز طبی ساده نیم‌کیلویی"}


def glove_type(name):
    n = str(name)
    if "دستكش" not in n and "دستکش" not in n and "glove" not in n.lower():
        return None
    if "وينيل" in n or "وینیل" in n:
        return "معاینه وینیل"
    if "نيتريل" in n or "نیتریل" in n:
        return "معاینه نیتریل"
    if "لاتكس" in n or "لاتکس" in n:
        return "معاینه لاتکس"
    if "نايلوني" in n or "نایلونی" in n:
        return "نایلونی یکبارمصرف"
    return "جراحی/استریل"


def load_his(path):
    df = pd.read_excel(path)
    df = df[df["كد"].notna()].copy()  # ردیف جمع کل حذف می‌شود
    for c in ["تعداد", "مانده"]:
        df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", ""), errors="coerce")
    df["بخش"] = df["بخش"].astype(str).str.strip()
    df["سهم_بیمه_کل"] = df["سهم بيمه"] + df["سهم بيمه دوم"] + df["يارانه سلامت"]
    df["محل_هزینه"] = "بیمه/بیمار"
    df.loc[(df["سهم_بیمه_کل"] == 0) & (df["سهم بيمار"] == 0), "محل_هزینه"] = "هزینه بخش (بدون صورتحساب)"
    return df


def load_beddays(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = next(s for s in wb.worksheets if "عملكرد" in s.title or "عملکرد" in s.title)
    rows = list(ws.iter_rows(values_only=True))
    months = [str(m).strip() for m in rows[0][1::4] if m]
    out = []
    for r in rows[2:]:
        if not r[0] or str(r[0]).startswith("ميانگين"):
            continue
        for i, m in enumerate(months):
            v = r[1 + i * 4]
            out.append({"ward": str(r[0]).strip(), "month": m,
                        "bed_days": v if isinstance(v, (int, float)) else None})
    bd = pd.DataFrame(out)
    return bd, months


def fmt(x):
    return f"{x:,.0f}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--his", required=True)
    ap.add_argument("--wards", required=True)
    ap.add_argument("--label", default="")
    ap.add_argument("--out", default=str(HERE / "output"))
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)

    df = load_his(a.his)
    bd, months = load_beddays(a.wards)
    bd_tot = bd.groupby("ward")["bed_days"].agg(["sum", "count"]).rename(
        columns={"sum": "تخت_روز", "count": "ماه_دارای_داده"})
    df["ward_bd"] = df["بخش"].map(WARD_MAP)
    df = df.join(bd_tot, on="ward_bd")

    # ---------- دستکش ----------
    df["نوع_دستکش"] = df["نام"].map(glove_type)
    g = df[df["نوع_دستکش"].notna()].copy()
    gw = g.pivot_table(index="بخش", columns="نوع_دستکش", values="تعداد", aggfunc="sum", fill_value=0)
    gw["ارزش_خرید_ریال"] = g.groupby("بخش")["جمع خريد"].sum()
    gw["تخت_روز"] = gw.index.map(lambda w: bd_tot["تخت_روز"].get(WARD_MAP.get(w)))
    exam = [c for c in ["معاینه وینیل", "معاینه نیتریل", "معاینه لاتکس"] if c in gw]
    gw["معاینه_به_ازای_تخت‌روز"] = (gw[exam].sum(axis=1) / gw["تخت_روز"]).round(1)
    gw = gw.sort_values("ارزش_خرید_ریال", ascending=False)

    price = (g[g["نوع_دستکش"].isin(exam)].groupby(["نام", "في خريد"])
             .agg(تعداد=("تعداد", "sum"), ارزش=("جمع خريد", "sum")).reset_index())
    billing = (g.groupby(["نوع_دستکش", "بخش", "محل_هزینه"])["تعداد"].sum().unstack(fill_value=0))
    ge = g[g["نوع_دستکش"].isin(["معاینه وینیل", "معاینه نیتریل", "معاینه لاتکس"])]
    # یک ردیف HIS ممکن است فقط بخشی از مقدارش روی پرونده باشد؛ پس سهم ریالی مبنا است
    eb = ge.assign(billed=ge["سهم_بیمه_کل"] + ge["سهم بيمار"]).groupby("بخش")[["billed", "جمع كل فروش"]].sum()
    eb["روی_پرونده"] = eb["billed"]
    eb["درصد"] = (eb["billed"] / eb["جمع كل فروش"] * 100).round(1)
    billed_wards = [f"{w} ({r['درصد']}٪)" for w, r in eb[eb["روی_پرونده"] > 0]
                    .sort_values("درصد", ascending=False).iterrows()]

    # ---------- آلبومین ----------
    al = df[df["نام"].astype(str).str.contains("albumin|آلبومين|آلبومین", case=False)].copy()
    alw = al.groupby("بخش").agg(ویال=("تعداد", "sum"), ارزش_خرید=("جمع خريد", "sum"),
                                 فروش=("جمع كل فروش", "sum"), سهم_بیمه=("سهم_بیمه_کل", "sum"),
                                 سهم_بیمار=("سهم بيمار", "sum"))
    alw["تخت_روز"] = alw.index.map(lambda w: bd_tot["تخت_روز"].get(WARD_MAP.get(w)))
    alw["ویال_در_۱۰۰_تخت‌روز"] = (alw["ویال"] / alw["تخت_روز"] * 100).round(1)
    alw["درصد_سهم_بیمه"] = (alw["سهم_بیمه"] / alw["فروش"] * 100).round(1)
    alw = alw.sort_values("ویال", ascending=False)
    al_brand = al.groupby(["كد", "نام", "في خريد"]).agg(ویال=("تعداد", "sum"), ارزش_خرید=("جمع خريد", "sum"),
                                                     سهم_بیمه=("سهم_بیمه_کل", "sum")).reset_index()

    # ---------- سایر اقلام پایلوت (مثل گاز) ----------
    pil = df[df["كد"].isin(PILOT_ITEMS)].copy()
    pil["قلم"] = pil["كد"].map(PILOT_ITEMS)
    pw = pil.groupby(["قلم", "بخش"]).agg(تعداد=("تعداد", "sum"), ارزش_خرید=("جمع خريد", "sum"),
                                       سهم_بیمه_و_بیمار=("سهم_بیمه_کل", "sum")).reset_index()
    pw["سهم_بیمه_و_بیمار"] += pil.groupby(["قلم", "بخش"])["سهم بيمار"].sum().values
    pw["تخت_روز"] = pw["بخش"].map(lambda w: bd_tot["تخت_روز"].get(WARD_MAP.get(w)))
    pw["در_۱۰۰_تخت‌روز"] = (pw["تعداد"] / pw["تخت_روز"] * 100).round(1)
    pw = pw.sort_values(["قلم", "ارزش_خرید"], ascending=[True, False])
    pprice = pil.groupby(["قلم", "في خريد"])["تعداد"].sum().reset_index()

    # ---------- کیفیت داده ----------
    dq = []
    unmapped = sorted(set(g["بخش"]) - set(WARD_MAP) - {"CSR", "اتاق عمل"})
    if unmapped:
        dq.append(("بخش‌هایی که تخت‌روز ندارند (غیربستری یا پشتیبانی)", "، ".join(unmapped)))
    if (df["حواله از انبار"] == 0).all():
        dq.append(("ستون «حواله از انبار» همه صفر است", "تحویل به بخش از این گزارش قابل جداسازی از مصرف نیست"))
    dq.append(("ستون «مانده» در سطح کل داروخانه است، نه بخش", "برای مصرف واقعی بخش، شمارش موجودی بخش لازم است"))
    dqdf = pd.DataFrame(dq, columns=["مشکل", "توضیح"])

    with pd.ExcelWriter(out / "his_analysis.xlsx", engine="openpyxl") as xw:
        gw.reset_index().to_excel(xw, sheet_name="دستکش_بخش", index=False)
        price.to_excel(xw, sheet_name="دستکش_قیمت", index=False)
        billing.reset_index().to_excel(xw, sheet_name="دستکش_صورتحساب", index=False)
        alw.reset_index().to_excel(xw, sheet_name="آلبومین_بخش", index=False)
        pw.to_excel(xw, sheet_name="سایر_اقلام_پایلوت", index=False)
        al_brand.to_excel(xw, sheet_name="آلبومین_برند", index=False)
        bd.pivot_table(index="ward", columns="month", values="bed_days", aggfunc="sum").reset_index() \
          .to_excel(xw, sheet_name="تخت_روز", index=False)
        dqdf.to_excel(xw, sheet_name="کیفیت_داده", index=False)
        for ws in xw.book.worksheets:
            ws.sheet_view.rightToLeft = True
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 20

    # ---------- گزارش ----------
    vin = g[g["نوع_دستکش"] == "معاینه وینیل"]
    nit = g[g["نوع_دستکش"] == "معاینه نیتریل"]
    L = [f"# تحلیل داده HIS: دستکش و آلبومین — {a.label}", "",
         f"تخت‌روز از فایل عملکرد بخش‌ها ({'، '.join(months)}). اعداد مصرف «خروج ثبت‌شده در HIS برای بخش» است، "
         "نه مصرف واقعی شمارش‌شده.", "",
         "## دستکش", "",
         f"- وینیل: {fmt(vin['تعداد'].sum())} عدد، ارزش خرید {fmt(vin['جمع خريد'].sum())} ریال",
         f"- نیتریل: {fmt(nit['تعداد'].sum())} عدد، ارزش خرید {fmt(nit['جمع خريد'].sum())} ریال",
         f"- کل دستکش (همه انواع): ارزش خرید {fmt(g['جمع خريد'].sum())} ریال", "",
         "| بخش | وینیل | نیتریل | تخت‌روز | دستکش معاینه به ازای تخت‌روز | ارزش خرید (ریال) |",
         "|---|---|---|---|---|---|"]
    for w, r in gw.head(15).iterrows():
        tb = "—" if pd.isna(r["تخت_روز"]) else fmt(r["تخت_روز"])
        pr = "—" if pd.isna(r["معاینه_به_ازای_تخت‌روز"]) else r["معاینه_به_ازای_تخت‌روز"]
        L.append(f"| {w} | {fmt(r.get('معاینه وینیل', 0))} | {fmt(r.get('معاینه نیتریل', 0))} | {tb} | {pr} | "
                 f"{fmt(r['ارزش_خرید_ریال'])} |")
    L += ["", "**قیمت خرید هر عدد در دوره:**", ""]
    for nm, s in price.groupby("نام"):
        L.append(f"- {nm.strip()}: " + "، ".join(f"{fmt(p)} ریال ({fmt(q)} عدد)"
                                               for p, q in zip(s["في خريد"], s["تعداد"])))
    L += ["", f"**صورتحساب دستکش معاینه:** فقط در این بخش‌ها بخشی از دستکش معاینه روی پرونده بیمار ثبت شده است "
          f"(درصد ریالی فروش): {'، '.join(billed_wards) or 'هیچ'}. در بقیه بخش‌ها سهم بیمه و بیمار صفر است و "
          "دستکش هزینه بخش محسوب شده است.", "",
          "## آلبومین ۲۰٪", "",
          f"- کل: {fmt(al['تعداد'].sum())} ویال، ارزش خرید {fmt(al['جمع خريد'].sum())} ریال، "
          f"فروش {fmt(al['جمع كل فروش'].sum())} ریال",
          f"- سهم بیمه: {fmt(al['سهم_بیمه_کل'].sum())} ریال "
          f"({al['سهم_بیمه_کل'].sum() / al['جمع كل فروش'].sum() * 100:.1f}٪ فروش)؛ "
          f"سهم بیمار: {fmt(al['سهم بيمار'].sum())} ریال", "",
          "**قیمت خرید هر ویال به تفکیک کد:**", ""] + [
          f"- کد {int(r['كد'])} — {str(r['نام']).strip()}: {fmt(r['في خريد'])} ریال × {fmt(r['ویال'])} ویال"
          for _, r in al_brand.sort_values("في خريد", ascending=False).iterrows()] + ["",
          "| بخش | ویال | تخت‌روز | ویال در ۱۰۰ تخت‌روز | ارزش خرید (ریال) | سهم بیمه ٪ |", "|---|---|---|---|---|---|"]
    for w, r in alw.iterrows():
        tb = "—" if pd.isna(r["تخت_روز"]) else fmt(r["تخت_روز"])
        rt = "—" if pd.isna(r["ویال_در_۱۰۰_تخت‌روز"]) else r["ویال_در_۱۰۰_تخت‌روز"]
        L.append(f"| {w} | {fmt(r['ویال'])} | {tb} | {rt} | {fmt(r['ارزش_خرید'])} | {r['درصد_سهم_بیمه']} |")
    for item, part in pw.groupby("قلم"):
        L += ["", f"## {item}", "",
              f"- کل: {fmt(part['تعداد'].sum())} واحد، ارزش خرید {fmt(part['ارزش_خرید'].sum())} ریال؛ "
              f"سهم ثبت‌شده روی پرونده: {fmt(part['سهم_بیمه_و_بیمار'].sum())} ریال",
              "- قیمت خرید هر واحد: " + "، ".join(
                  f"{fmt(r['في خريد'])} ریال ({fmt(r['تعداد'])})"
                  for _, r in pprice[pprice["قلم"] == item].iterrows()), "",
              "| بخش | تعداد | تخت‌روز | در ۱۰۰ تخت‌روز | ارزش خرید (ریال) |", "|---|---|---|---|---|"]
        for _, r in part.iterrows():
            tb = "—" if pd.isna(r["تخت_روز"]) else fmt(r["تخت_روز"])
            rt = "—" if pd.isna(r["در_۱۰۰_تخت‌روز"]) else r["در_۱۰۰_تخت‌روز"]
            L.append(f"| {r['بخش']} | {fmt(r['تعداد'])} | {tb} | {rt} | {fmt(r['ارزش_خرید'])} |")
    L += ["", "## محدودیت‌های داده", ""] + [f"- {m}: {t}" for m, t in dq]
    (out / "his_report.md").write_text("\n".join(L) + "\n", encoding="utf-8")
    print(f"خروجی: {out / 'his_analysis.xlsx'}\nگزارش: {out / 'his_report.md'}")


if __name__ == "__main__":
    main()
