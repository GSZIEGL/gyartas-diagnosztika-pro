
def safe_completion_pct(planned, demand):
    """Pandas/NumPy kompatibilis teljesítési százalék 0-100 között."""
    planned_s = pd.to_numeric(planned, errors="coerce").fillna(0)
    demand_s = pd.to_numeric(demand, errors="coerce").fillna(0)
    pct = pd.Series(0.0, index=demand_s.index)
    mask = demand_s > 0
    pct.loc[mask] = planned_s.loc[mask] / demand_s.loc[mask] * 100
    return pct.clip(lower=0, upper=100).round(1)


import io
import os
import json
from datetime import datetime, date as dt_date
from pathlib import Path

try:
    from supabase import create_client
except Exception:
    create_client = None
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.graphics.shapes import Drawing, Rect, String, Line
except Exception:
    SimpleDocTemplate = None



st.set_page_config(
    page_title="Gyártási Diagnosztika PRO SaaS V4 V4.1 V2 SaaS.7.5.4.4.3.3.2.2",
    page_icon="🏭",
    layout="wide"
)


# ------------------------------------------------------------
# Stílus
# ------------------------------------------------------------
st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.1rem;
        font-weight: 900;
        color: #0f172a;
        margin-bottom: .2rem;
    }
    .subtitle {
        color: #475569;
        font-size: 1rem;
        margin-bottom: 1.2rem;
    }
    .kpi-card {
        background: linear-gradient(135deg, #f8fafc, #eef2ff);
        border: 1px solid #cbd5e1;
        border-radius: 18px;
        padding: 18px;
        box-shadow: 0 8px 22px rgba(15,23,42,.08);
        min-height: 125px;
    }
    .kpi-label {
        font-size: .82rem;
        color: #475569;
        text-transform: uppercase;
        font-weight: 800;
        letter-spacing: .04em;
    }
    .kpi-value {
        font-size: 1.85rem;
        color: #0f172a;
        font-weight: 950;
        margin-top: 6px;
    }
    .kpi-note {
        font-size: .86rem;
        color: #334155;
        margin-top: 6px;
    }
    .insight-card {
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-left: 7px solid #2563eb;
        border-radius: 16px;
        padding: 14px 16px;
        margin-bottom: 10px;
        box-shadow: 0 6px 18px rgba(15,23,42,.06);
        color: #0f172a !important;
        font-weight: 650;
        line-height: 1.45;
    }
    .insight-card * {
        color: #0f172a !important;
    }
    .danger {
        border-left-color: #dc2626;
        background: #fff7f7;
    }
    .warning {
        border-left-color: #f59e0b;
        background: #fffbeb;
    }
    .success {
        border-left-color: #16a34a;
        background: #f0fdf4;
    }
    .small-muted {
        color:#64748b;
        font-size:.86rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ------------------------------------------------------------
# Segédfüggvények
# ------------------------------------------------------------
REQUIRED_PROD_COLS = [
    "Dátum", "Műszak", "Dolgozó", "Gép", "Termék",
    "Gyártott_db", "Selejt_db", "Állásidő_perc"
]

REQUIRED_MACHINE_COLS = ["Gép", "Kapacitás_db_óra", "Óradíj", "Kritikus_gép"]
REQUIRED_PRODUCT_COLS = ["Termék", "Eladási_ár", "Anyagköltség"]
OPTIONAL_ORDER_COLS = ["Rendelés_ID", "Vevő", "Termék", "Rendelt_db", "Határidő", "Prioritás"]


def fmt_num(x, digits=0):
    if pd.isna(x):
        return "-"
    if digits == 0:
        return f"{x:,.0f}".replace(",", " ")
    return f"{x:,.{digits}f}".replace(",", " ")


def fmt_pct(x, digits=1):
    if pd.isna(x):
        return "-"
    return f"{x:.{digits}f}%"


def fmt_huf(x):
    if pd.isna(x):
        return "-"
    try:
        x = float(x)
    except Exception:
        return "-"
    if abs(x) < 0.5:
        x = 0
    return f"{x:,.0f} Ft".replace(",", " ")


def show_kpi(label: str, value: str, note: str = ""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def safe_read_excel(uploaded_file) -> Dict[str, pd.DataFrame]:
    return pd.read_excel(uploaded_file, sheet_name=None)


def find_sheet(sheets: Dict[str, pd.DataFrame], possible_names: List[str]) -> pd.DataFrame:
    lower_map = {name.lower(): name for name in sheets.keys()}
    for name in possible_names:
        if name.lower() in lower_map:
            return sheets[lower_map[name.lower()]]
    # fallback: first sheet
    return list(sheets.values())[0]


def validate_columns(df: pd.DataFrame, required: List[str], sheet_name: str):
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"A(z) {sheet_name} munkalapon hiányzó oszlopok: {', '.join(missing)}")
        st.stop()


def prepare_data(prod: pd.DataFrame, machines: pd.DataFrame, products: pd.DataFrame) -> pd.DataFrame:
    prod = prod.copy()
    machines = machines.copy()
    products = products.copy()

    prod["Dátum"] = pd.to_datetime(prod["Dátum"], errors="coerce")
    for col in ["Gyártott_db", "Selejt_db", "Állásidő_perc"]:
        prod[col] = pd.to_numeric(prod[col], errors="coerce").fillna(0)

    machines["Kapacitás_db_óra"] = pd.to_numeric(machines["Kapacitás_db_óra"], errors="coerce").fillna(0)
    machines["Óradíj"] = pd.to_numeric(machines["Óradíj"], errors="coerce").fillna(0)
    if "Elérhető_óra_nap" not in machines.columns:
        machines["Elérhető_óra_nap"] = 8
    machines["Elérhető_óra_nap"] = pd.to_numeric(machines["Elérhető_óra_nap"], errors="coerce").fillna(8)

    products["Eladási_ár"] = pd.to_numeric(products["Eladási_ár"], errors="coerce").fillna(0)
    products["Anyagköltség"] = pd.to_numeric(products["Anyagköltség"], errors="coerce").fillna(0)
    if "Prioritási_súly" not in products.columns:
        products["Prioritási_súly"] = 3
    products["Prioritási_súly"] = pd.to_numeric(products["Prioritási_súly"], errors="coerce").fillna(3)

    df = prod.merge(machines, on="Gép", how="left").merge(products, on="Termék", how="left")

    # V1 feltételezés: egy sor egy körülbelül 1 órás termelési blokk.
    df["Munkaóra"] = 1.0
    df["Jó_db"] = (df["Gyártott_db"] - df["Selejt_db"]).clip(lower=0)
    df["Selejt_%"] = np.where(df["Gyártott_db"] > 0, df["Selejt_db"] / df["Gyártott_db"] * 100, 0)
    df["Állásidő_%"] = np.minimum(df["Állásidő_perc"] / 60 * 100, 100)
    df["Elérhetőség_%"] = (100 - df["Állásidő_%"]).clip(lower=0)
    df["Teljesítmény_%"] = np.where(
        df["Kapacitás_db_óra"] > 0,
        df["Gyártott_db"] / df["Kapacitás_db_óra"] * 100,
        0
    )
    df["Teljesítmény_%"] = np.minimum(df["Teljesítmény_%"], 140)
    df["Minőség_%"] = np.where(df["Gyártott_db"] > 0, df["Jó_db"] / df["Gyártott_db"] * 100, 0)
    df["OEE_light_%"] = df["Elérhetőség_%"] * df["Teljesítmény_%"] * df["Minőség_%"] / 10000

    df["Árbevétel"] = df["Jó_db"] * df["Eladási_ár"]
    df["Anyagköltség_össz"] = df["Gyártott_db"] * df["Anyagköltség"]

    # PRO: gépköltség korrekció.
    # Korábban minden sorra teljes óradíj ment, ami irreálisan negatív fedezetot okozhatott.
    df["Becsült_gépóra"] = np.where(
        df["Kapacitás_db_óra"] > 0,
        df["Gyártott_db"] / df["Kapacitás_db_óra"],
        df["Munkaóra"]
    )
    df["Becsült_gépóra"] = df["Becsült_gépóra"].clip(lower=0.05, upper=12)
    df["Gépköltség"] = df["Óradíj"] * df["Becsült_gépóra"]

    df["Becsült_fedezet"] = df["Árbevétel"] - df["Anyagköltség_össz"] - df["Gépköltség"]

    return df


def aggregate_metrics(df: pd.DataFrame, group_cols: List[str]) -> pd.DataFrame:
    out = df.groupby(group_cols, as_index=False).agg(
        Gyártott_db=("Gyártott_db", "sum"),
        Jó_db=("Jó_db", "sum"),
        Selejt_db=("Selejt_db", "sum"),
        Állásidő_perc=("Állásidő_perc", "sum"),
        Árbevétel=("Árbevétel", "sum"),
        Becsült_fedezet=("Becsült_fedezet", "sum"),
        Átlag_OEE=("OEE_light_%", "mean"),
        Átlag_teljesítmény=("Teljesítmény_%", "mean"),
        Sorok=("Gyártott_db", "count")
    )
    out["Selejt_%"] = np.where(out["Gyártott_db"] > 0, out["Selejt_db"] / out["Gyártott_db"] * 100, 0)
    out["Fedezet/db"] = np.where(out["Jó_db"] > 0, out["Becsült_fedezet"] / out["Jó_db"], 0)
    return out


def build_worker_machine_matrix(df: pd.DataFrame) -> pd.DataFrame:
    pair = aggregate_metrics(df, ["Dolgozó", "Gép"])
    # Kompatibilitási pont: teljesítmény + minőség + fedezet/db, 0-100 környék.
    perf = pair["Átlag_teljesítmény"].clip(0, 120) / 120 * 50
    quality = (100 - pair["Selejt_%"].clip(0, 20) * 5).clip(0, 100) / 100 * 25
    fedezet_norm = pair["Fedezet/db"]
    if fedezet_norm.max() != fedezet_norm.min():
        fedezet_score = (fedezet_norm - fedezet_norm.min()) / (fedezet_norm.max() - fedezet_norm.min()) * 25
    else:
        fedezet_score = 12.5
    pair["Kompatibilitási_pont"] = (perf + quality + fedezet_score).round(1)
    matrix = pair.pivot_table(
        index="Dolgozó",
        columns="Gép",
        values="Kompatibilitási_pont",
        aggfunc="mean"
    ).round(1)
    return matrix, pair


def generate_recommendations(df: pd.DataFrame, pair: pd.DataFrame) -> List[Tuple[str, str]]:
    """Mindig adjon vezetői javaslatokat, ne csak extrém eltérésnél."""
    recs = []

    shift = aggregate_metrics(df, ["Műszak"])
    if len(shift) >= 2:
        worst_shift = shift.sort_values("Átlag_OEE").iloc[0]
        best_shift = shift.sort_values("Átlag_OEE", ascending=False).iloc[0]
        diff = best_shift["Átlag_OEE"] - worst_shift["Átlag_OEE"]
        recs.append((
            "warning" if diff >= 3 else "success",
            f"Műszakhatás: a(z) {best_shift['Műszak']} műszak OEE-je {diff:.1f} ponttal jobb, mint a(z) {worst_shift['Műszak']} műszaké. "
            f"Érdemes megnézni, hogy ember-, gép- vagy termékösszetétel okozza-e."
        ))

    machine = aggregate_metrics(df, ["Gép"])
    if not machine.empty:
        worst_machine = machine.sort_values(["Állásidő_perc", "Selejt_%"], ascending=False).iloc[0]
        best_machine = machine.sort_values("Átlag_OEE", ascending=False).iloc[0]
        recs.append((
            "danger" if worst_machine["Állásidő_perc"] > machine["Állásidő_perc"].median() else "warning",
            f"Gépdiagnosztika: a(z) {worst_machine['Gép']} gépen a legmagasabb az állásidő/selejt kombináció. "
            f"A legjobb OEE-t jelenleg a(z) {best_machine['Gép']} hozza."
        ))

    worker = aggregate_metrics(df, ["Dolgozó"])
    if not worker.empty:
        top_worker = worker.sort_values("Átlag_OEE", ascending=False).iloc[0]
        low_worker = worker.sort_values("Átlag_OEE").iloc[0]
        recs.append((
            "success",
            f"Dolgozói teljesítmény: {top_worker['Dolgozó']} hozza a legjobb átlagos OEE-t ({top_worker['Átlag_OEE']:.1f}%). "
            f"{low_worker['Dolgozó']} esetében érdemes megnézni, hogy rossz gépen vagy nehezebb terméken dolgozik-e."
        ))

    best_pairs = pair.sort_values("Kompatibilitási_pont", ascending=False).head(3)
    if not best_pairs.empty:
        text = "; ".join([f"{r['Dolgozó']} → {r['Gép']} ({r['Kompatibilitási_pont']:.0f} pont)" for _, r in best_pairs.iterrows()])
        recs.append(("success", f"Legjobb dolgozó–gép párosok: {text}. Ezeket a párosokat érdemes preferálni beosztáskor."))

    weak_pairs = pair[pair["Sorok"] >= 3].sort_values("Kompatibilitási_pont").head(3)
    if not weak_pairs.empty:
        text = "; ".join([f"{r['Dolgozó']} + {r['Gép']} ({r['Kompatibilitási_pont']:.0f} pont)" for _, r in weak_pairs.iterrows()])
        recs.append(("warning", f"Figyelendő párosítások: {text}. Nem biztos, hogy rossz dolgozókról van szó, lehet, hogy rossz gép–ember párosítás."))

    product = aggregate_metrics(df, ["Termék"])
    if not product.empty:
        best_product = product.sort_values("Fedezet/db", ascending=False).iloc[0]
        worst_product = product.sort_values("Fedezet/db").iloc[0]
        recs.append((
            "warning",
            f"Termék/fedezet: a(z) {best_product['Termék']} termék fedezet/db alapján a legerősebb, "
            f"a(z) {worst_product['Termék']} a leggyengébb. Gyártási prioritásnál ezt érdemes figyelembe venni."
        ))

    if len(recs) == 0:
        recs.append(("warning", "Még kevés adat van, de az app már felépítette az alap mutatókat. Tölts fel több sort vagy hosszabb időszakot."))

    return recs


def recommended_assignment(pair: pd.DataFrame) -> pd.DataFrame:
    # Egyszerű V1: minden gépre a legjobb kompatibilitású dolgozót ajánlja,
    # egy dolgozó több gépre is ajánlható lehet. V2-ben jöhet optimalizáló algoritmus.
    best = pair.sort_values("Kompatibilitási_pont", ascending=False).groupby("Gép", as_index=False).head(1)
    best = best[["Gép", "Dolgozó", "Kompatibilitási_pont", "Átlag_teljesítmény", "Selejt_%", "Fedezet/db"]]
    return best.sort_values("Gép")





def calculate_advisor_scores(df: pd.DataFrame, fulfillment_df: pd.DataFrame, capacity_df: pd.DataFrame, impact_df: pd.DataFrame) -> Dict[str, float]:
    """PRO.4.3.2 vezetői score-ok 0-100 skálán."""
    if df is None or df.empty:
        return {"Egészségpont": 0, "Kapacitáskockázat": 0, "Határidőkockázat": 0, "Fedezetveszteség_Ft": 0, "OEE": 0, "Selejt_%": 0}
    avg_oee = float(df["OEE_light_%"].mean()) if "OEE_light_%" in df.columns else 0
    total_qty = df["Gyártott_db"].sum() if "Gyártott_db" in df.columns else 0
    scrap_pct = df["Selejt_db"].sum() / total_qty * 100 if total_qty else 0
    capacity_risk = 0
    if capacity_df is not None and not capacity_df.empty and "Kihasználtság_%" in capacity_df.columns:
        max_util = float(capacity_df["Kihasználtság_%"].max())
        capacity_risk = min(100, max(0, (max_util - 70) * 2.5))
    deadline_risk = 0
    if fulfillment_df is not None and not fulfillment_df.empty and "Teljesítés_%" in fulfillment_df.columns:
        avg_fulfillment = float(fulfillment_df["Teljesítés_%"].mean())
        deadline_risk = max(0, 100 - avg_fulfillment)
    lost_fedezet = float(impact_df["Becsült_havi_hatás_Ft"].clip(lower=0).sum()) if impact_df is not None and not impact_df.empty and "Becsült_havi_hatás_Ft" in impact_df.columns else 0
    health = avg_oee * 0.45 + max(0, 100 - scrap_pct * 10) * 0.25 + max(0, 100 - capacity_risk) * 0.15 + max(0, 100 - deadline_risk) * 0.15
    return {"Egészségpont": round(max(0, min(100, health)), 1), "Kapacitáskockázat": round(max(0, min(100, capacity_risk)), 1), "Határidőkockázat": round(max(0, min(100, deadline_risk)), 1), "Fedezetveszteség_Ft": round(lost_fedezet, 0), "OEE": round(avg_oee, 1), "Selejt_%": round(scrap_pct, 2)}


def score_label(value: float, inverse: bool = False) -> Tuple[str, str]:
    try:
        v = float(value)
    except Exception:
        v = 0
    if inverse:
        if v < 35: return "🟢 Alacsony", "success"
        if v < 70: return "🟡 Közepes", "warning"
        return "🔴 Magas", "danger"
    if v >= 75: return "🟢 Jó", "success"
    if v >= 50: return "🟡 Közepes", "warning"
    return "🔴 Gyenge", "danger"


def build_action_plan(df: pd.DataFrame, pair: pd.DataFrame, impact_df: pd.DataFrame, capacity_df: pd.DataFrame, fulfillment_df: pd.DataFrame) -> pd.DataFrame:
    actions = []
    if fulfillment_df is not None and not fulfillment_df.empty and "Hiány_db" in fulfillment_df.columns:
        shortage = fulfillment_df[fulfillment_df["Hiány_db"] > 0].sort_values("Hiány_db", ascending=False)
        if not shortage.empty:
            r = shortage.iloc[0]
            actions.append({"Prioritás":"Magas","Akció":f"Kapacitásbővítés vagy átütemezés a(z) {r['Termék']} termékre","Érintett":r["Termék"],"Miért?":f"{r['Hiány_db']:.0f} db hiány a tervben","Becsült_hatás":0})
    if capacity_df is not None and not capacity_df.empty and "Kihasználtság_%" in capacity_df.columns:
        bottleneck = capacity_df.sort_values("Kihasználtság_%", ascending=False).iloc[0]
        if bottleneck["Kihasználtság_%"] >= 90:
            actions.append({"Prioritás":"Magas","Akció":f"Szűk keresztmetszet kezelése: {bottleneck['Gép']}","Érintett":bottleneck["Gép"],"Miért?":f"{bottleneck['Kihasználtság_%']:.1f}% kapacitáskihasználtság","Becsült_hatás":0})
    if impact_df is not None and not impact_df.empty:
        med = impact_df["Becsült_havi_hatás_Ft"].median() if "Becsült_havi_hatás_Ft" in impact_df.columns else 0
        for _, r in impact_df.head(4).iterrows():
            actions.append({"Prioritás":"Magas" if r.get("Becsült_havi_hatás_Ft",0)>med else "Közepes","Akció":r.get("Javaslat","Beavatkozási pont vizsgálata"),"Érintett":r.get("Elem",""),"Miért?":r.get("Probléma",""),"Becsült_hatás":r.get("Becsült_havi_hatás_Ft",0)})
    if pair is not None and not pair.empty:
        for _, r in pair.sort_values("Kompatibilitási_pont", ascending=False).head(3).iterrows():
            actions.append({"Prioritás":"Közepes","Akció":f"{r['Dolgozó']} kerüljön gyakrabban erre a gépre: {r['Gép']}","Érintett":f"{r['Dolgozó']} + {r['Gép']}","Miért?":f"Erős dolgozó-gép kompatibilitás: {r['Kompatibilitási_pont']:.0f} pont","Becsült_hatás":max(0, r.get("Fedezet/db",0))*100})
    out=pd.DataFrame(actions)
    if out.empty: return out
    order={"Magas":0,"Közepes":1,"Alacsony":2}
    out["_sort"]=out["Prioritás"].map(order).fillna(9)
    return out.sort_values(["_sort","Becsült_hatás"],ascending=[True,False]).drop(columns=["_sort"]).head(8)




def build_trend_insights(history_df: pd.DataFrame) -> List[Tuple[str, str]]:
    """PRO V4: vezetői trendmegállapítások több időszak alapján."""
    if history_df is None or history_df.empty or len(history_df) < 2:
        return [("info", "Ments el legalább két időszakot, hogy trendmegállapítás készüljön.")]

    h = history_df.sort_values("week").copy()
    for col in ["oee", "selejt_pct", "allasido_perc", "javitasi_potencial_ft", "rendeles_teljesites_pct"]:
        if col in h.columns:
            h[col] = pd.to_numeric(h[col], errors="coerce")

    first = h.iloc[0]
    prev = h.iloc[-2]
    last = h.iloc[-1]
    notes = []

    def delta(col):
        if col not in h.columns or pd.isna(last.get(col)) or pd.isna(prev.get(col)):
            return None
        return float(last[col]) - float(prev[col])

    d_oee = delta("oee")
    if d_oee is not None:
        notes.append(("success" if d_oee >= 0 else "warning", f"OEE előző időszakhoz képest: {d_oee:+.1f} pont."))

    d_scrap = delta("selejt_pct")
    if d_scrap is not None:
        notes.append(("success" if d_scrap <= 0 else "danger", f"Selejt változás előző időszakhoz képest: {d_scrap:+.2f} százalékpont."))

    d_down = delta("allasido_perc")
    if d_down is not None:
        notes.append(("success" if d_down <= 0 else "warning", f"Állásidő változás előző időszakhoz képest: {d_down:+.0f} perc."))

    d_pot = delta("javitasi_potencial_ft")
    if d_pot is not None:
        notes.append(("success" if d_pot <= 0 else "warning", f"Javítási potenciál változás: {fmt_huf(d_pot)}. Ha nő, több pénz maradhat az asztalon."))

    if "oee" in h.columns and h["oee"].notna().sum() >= 3:
        last3 = h["oee"].dropna().tail(3).tolist()
        if len(last3) == 3 and last3[0] > last3[1] > last3[2]:
            notes.append(("danger", "Az OEE három egymást követő mentett időszakban romlott. Ez PRO szintű beavatkozási jel."))
        elif len(last3) == 3 and last3[0] < last3[1] < last3[2]:
            notes.append(("success", "Az OEE három egymást követő időszakban javult. Érdemes az aktuális működést standardizálni."))

    return notes[:6]


def build_prev_period_delta_table(history_df: pd.DataFrame) -> pd.DataFrame:
    if history_df is None or history_df.empty or len(history_df) < 2:
        return pd.DataFrame()
    h = history_df.sort_values("week").copy()
    metrics = [
        ("OEE", "oee", "pont"),
        ("Selejt %", "selejt_pct", "százalékpont"),
        ("Állásidő", "allasido_perc", "perc"),
        ("Rendelésteljesítés", "rendeles_teljesites_pct", "%"),
        ("Javítási potenciál", "javitasi_potencial_ft", "Ft"),
    ]
    rows = []
    prev = h.iloc[-2]
    last = h.iloc[-1]
    for label, col, unit in metrics:
        if col not in h.columns:
            continue
        a = pd.to_numeric(pd.Series([prev.get(col)]), errors="coerce").iloc[0]
        b = pd.to_numeric(pd.Series([last.get(col)]), errors="coerce").iloc[0]
        if pd.isna(a) or pd.isna(b):
            continue
        diff = b - a
        rows.append({
            "Mutató": label,
            "Előző": round(float(a), 2),
            "Aktuális": round(float(b), 2),
            "Változás": round(float(diff), 2),
            "Egység": unit,
        })
    return pd.DataFrame(rows)


def normalized_pair_score_table(pair_df: pd.DataFrame) -> pd.DataFrame:
    """Dolgozó-gép mátrix 0-100 normalizált pontszámmal.

    A korábbi abszolút pontozás miatt sok cella sárga/narancs lett.
    Itt a legjobb párosok zöldek lesznek, mert a tényleges mezőnyön belüli relatív helyzetet is nézzük.
    """
    if pair_df is None or pair_df.empty or "Dolgozó" not in pair_df.columns or "Gép" not in pair_df.columns:
        return pd.DataFrame()

    work = pair_df.copy()
    score_col = "Kompatibilitási_pont" if "Kompatibilitási_pont" in work.columns else None
    if score_col is None:
        return pd.DataFrame()

    work[score_col] = pd.to_numeric(work[score_col], errors="coerce").fillna(0)
    lo = float(work[score_col].min())
    hi = float(work[score_col].max())
    if hi > lo:
        work["Relatív_pont"] = 45 + (work[score_col] - lo) / (hi - lo) * 55
    else:
        work["Relatív_pont"] = work[score_col]

    # Ha az abszolút pont eleve magas, ne húzzuk le túlzottan.
    work["Vizuális_pont"] = np.maximum(work[score_col], work["Relatív_pont"]).clip(0, 100).round(0)

    return work.pivot_table(
        index="Dolgozó",
        columns="Gép",
        values="Vizuális_pont",
        aggfunc="mean"
    ).round(0)


def heatmap_symbol_from_score(value):
    try:
        x = float(value)
    except Exception:
        return "⚪"
    if x >= 80:
        return "🟢"
    if x >= 65:
        return "🟡"
    if x >= 50:
        return "🟠"
    if x > 0:
        return "🔴"
    return "⚪"


def make_symbol_heatmap_from_matrix(score_matrix: pd.DataFrame) -> pd.DataFrame:
    if score_matrix is None or score_matrix.empty:
        return pd.DataFrame()
    return score_matrix.apply(lambda col: col.map(heatmap_symbol_from_score))


def build_heatmap_symbols(matrix: pd.DataFrame) -> pd.DataFrame:
    if matrix is None or matrix.empty: return pd.DataFrame()
    def sym(v):
        try: x=float(v)
        except Exception: return "⚪"
        if x>=80: return "🟢"
        if x>=65: return "🟡"
        if x>0: return "🔴"
        return "⚪"
    return matrix.apply(lambda col: col.map(sym))


def simulate_what_if(df: pd.DataFrame, fulfillment_df: pd.DataFrame, capacity_df: pd.DataFrame, impact_df: pd.DataFrame, extra_capacity_pct: float=0, scrap_reduction_pct: float=0, oee_improvement_pct: float=0) -> pd.DataFrame:
    base_fedezet=float(df["Becsült_fedezet"].sum()) if df is not None and not df.empty and "Becsült_fedezet" in df.columns else 0
    total_qty=float(df["Gyártott_db"].sum()) if df is not None and not df.empty and "Gyártott_db" in df.columns else 0
    scrap_qty=float(df["Selejt_db"].sum()) if df is not None and not df.empty and "Selejt_db" in df.columns else 0
    avg_fedezet_per_good=float(df["Becsült_fedezet"].sum()/max(df["Jó_db"].sum(),1)) if df is not None and not df.empty and "Jó_db" in df.columns else 0
    cap_gain=total_qty*(extra_capacity_pct/100)*avg_fedezet_per_good*0.4
    oee_gain=total_qty*(oee_improvement_pct/100)*avg_fedezet_per_good*0.5
    scrap_gain=scrap_qty*(scrap_reduction_pct/100)*max(avg_fedezet_per_good,0)
    shortage_before=fulfillment_df["Hiány_db"].sum() if fulfillment_df is not None and not fulfillment_df.empty and "Hiány_db" in fulfillment_df.columns else 0
    shortage_after=max(0, shortage_before*(1-(extra_capacity_pct+oee_improvement_pct)/100))
    return pd.DataFrame([
        {"Mutató":"Becsült fedezet jelenleg","Érték":base_fedezet},
        {"Mutató":"Kapacitásnövelés becsült hatása","Érték":cap_gain},
        {"Mutató":"OEE-javulás becsült hatása","Érték":oee_gain},
        {"Mutató":"Selejtcsökkentés becsült hatása","Érték":scrap_gain},
        {"Mutató":"Becsült fedezet what-if után","Érték":base_fedezet+cap_gain+oee_gain+scrap_gain},
        {"Mutató":"Hiány előtte db","Érték":shortage_before},
        {"Mutató":"Hiány what-if után db","Érték":shortage_after},
    ])


def make_pdf_action_card(action, width=500):
    pr=str(action.get("Prioritás","Közepes"))
    color="#dc2626" if pr=="Magas" else "#f59e0b" if pr=="Közepes" else "#16a34a"
    bg="#fee2e2" if pr=="Magas" else "#fef3c7" if pr=="Közepes" else "#dcfce7"
    text=f"<b>{pdf_safe_text(action.get('Akció',''))}</b><br/>{pdf_safe_text(action.get('Miért?',''))}<br/><b>Becsült hatás:</b> {fmt_huf(action.get('Becsült_hatás',0))}"
    t=Table([[Paragraph(pr, ParagraphStyle("Pr", fontSize=9, textColor=colors.HexColor(color), alignment=1)), Paragraph(text, ParagraphStyle("Act", fontSize=8.5, leading=11, textColor=colors.HexColor("#0f172a")))]], colWidths=[2.2*cm, width-2.2*cm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor(bg)),("BOX",(0,0),(-1,-1),0.5,colors.HexColor(color)),("LINEBEFORE",(0,0),(0,-1),5,colors.HexColor(color)),("VALIGN",(0,0),(-1,-1),"TOP"),("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7)]))
    return t




def estimate_lost_revenue_by_product(summary_df: pd.DataFrame, df: pd.DataFrame = None) -> pd.DataFrame:
    """Termékszintű hiányhoz becsült kieső árbevétel/fedezet."""
    if summary_df is None or summary_df.empty:
        return pd.DataFrame()

    out = summary_df.copy()

    price_map = {}
    margin_map = {}
    if df is not None and not df.empty:
        if "Eladási_ár" in df.columns:
            price_map = df.groupby("Termék")["Eladási_ár"].mean().to_dict()
        if "Eladási_ár" in df.columns and "Anyagköltség" in df.columns:
            margin_map = (df.groupby("Termék")["Eladási_ár"].mean() - df.groupby("Termék")["Anyagköltség"].mean()).to_dict()

    out["Egységár"] = out["Termék"].map(price_map).fillna(0)
    out["Fedezet_db"] = out["Termék"].map(margin_map).fillna(out["Egységár"] * 0.3)
    out["Kieső_árbevétel_Ft"] = out["Hiány_db"].fillna(0) * out["Egységár"].fillna(0)
    out["Kieső_fedezet_Ft"] = out["Hiány_db"].fillna(0) * out["Fedezet_db"].fillna(0)
    return out.sort_values("Kieső_fedezet_Ft", ascending=False)


def build_causal_chain(plan_df: pd.DataFrame, fulfillment_df: pd.DataFrame, pair: pd.DataFrame, capacity_df: pd.DataFrame, df: pd.DataFrame) -> pd.DataFrame:
    """Termék → fő szűk keresztmetszet gép → javasolt dolgozó/akció."""
    summary = summarize_plan_by_product(plan_df, fulfillment_df)
    if summary is None or summary.empty:
        return pd.DataFrame()

    rows = []
    for _, s in summary.iterrows():
        product = s.get("Termék")
        shortage = float(s.get("Hiány_db", 0) or 0)
        fulfill = float(s.get("Teljesítés_%", 0) or 0)

        product_plan = plan_df[plan_df["Termék"].eq(product)] if plan_df is not None and not plan_df.empty else pd.DataFrame()
        active = product_plan[~product_plan["Gép"].isin(["Kapacitáshiány", "Nincs adat"])] if not product_plan.empty else pd.DataFrame()

        if not active.empty:
            main_machine = active.groupby("Gép", as_index=False).agg(Tervezett_db=("Tervezett_db", "sum")).sort_values("Tervezett_db", ascending=False).iloc[0]["Gép"]
        else:
            main_machine = capacity_df.sort_values("Kihasználtság_%", ascending=False).iloc[0]["Gép"] if capacity_df is not None and not capacity_df.empty else "Nincs adat"

        best_worker = "Nincs adat"
        best_point = np.nan
        if pair is not None and not pair.empty and main_machine != "Nincs adat":
            cand = pair[pair["Gép"].eq(main_machine)].sort_values("Kompatibilitási_pont", ascending=False)
            if not cand.empty:
                best_worker = cand.iloc[0]["Dolgozó"]
                best_point = cand.iloc[0]["Kompatibilitási_pont"]

        if shortage > 0:
            cause = f"{main_machine} kapacitása / termék-gép teljesítménye korlátozza" if main_machine != "Nincs adat" else "Nincs megfelelő termék-gép múltbeli adat"
            action = f"{product}: kapacitásbővítés, átütemezés vagy termék átterhelése"
        else:
            cause = "A jelenlegi kapacitás fedezi az igényt"
            action = f"{product}: jelenlegi terv tartható"

        rows.append({
            "Termék": product,
            "Teljesítés_%": round(fulfill, 1),
            "Hiány_db": round(shortage, 0),
            "Fő_gép": main_machine,
            "Javasolt_dolgozó": best_worker,
            "Dolgozó_gép_pont": round(float(best_point), 1) if pd.notna(best_point) else "",
            "Valószínű_ok": cause,
            "Javasolt_akció": action
        })

    return pd.DataFrame(rows).sort_values(["Teljesítés_%", "Hiány_db"], ascending=[True, False])


def build_top_critical_orders(orders_df: pd.DataFrame, plan_df: pd.DataFrame) -> pd.DataFrame:
    """Top kritikus rendelések: rendelés szinten hol van hiány."""
    if orders_df is None or orders_df.empty:
        return pd.DataFrame()

    if plan_df is not None and not plan_df.empty and "Rendelés_ID" in plan_df.columns:
        active = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])]
        planned_by_order = active.groupby("Rendelés_ID", as_index=False).agg(Tervezett_db=("Tervezett_db", "sum"))
    else:
        planned_by_order = pd.DataFrame(columns=["Rendelés_ID", "Tervezett_db"])

    out = orders_df.copy()
    out = out.merge(planned_by_order, on="Rendelés_ID", how="left")
    out["Tervezett_db"] = out["Tervezett_db"].fillna(0)
    out["Hiány_db"] = (out["Rendelt_db"] - out["Tervezett_db"]).clip(lower=0)
    out["Teljesítés_%"] = safe_completion_pct(out["Tervezett_db"], out["Rendelt_db"])
    out["Kritikusság"] = (100 - out["Teljesítés_%"]) + (6 - out["Prioritás"].clip(1,5)) * 10
    return out.sort_values(["Kritikusság", "Határidő"], ascending=[False, True]).head(10)


def make_pdf_executive_cover(advisor_scores, fulfillment_summary, impact_df, causal_df, action_plan_df):
    """Vezetői címlap blokk."""
    parts = []
    parts.append(Paragraph("Ügyvezetői összefoglaló", ParagraphStyle("CoverTitle", fontSize=16, leading=20, textColor=colors.HexColor("#1e3a8a"))))

    health = advisor_scores.get("Egészségpont", 0) if advisor_scores else 0
    lost = advisor_scores.get("Fedezetveszteség_Ft", 0) if advisor_scores else 0
    fulfillment_rate = 0
    if fulfillment_summary is not None and not fulfillment_summary.empty:
        total_need = fulfillment_summary["Igényelt_db"].sum()
        total_plan = fulfillment_summary["Tervezett_db"].sum()
        fulfillment_rate = min(100, total_plan / total_need * 100) if total_need else 0

    critical = causal_df[causal_df["Hiány_db"] > 0].head(1) if causal_df is not None and not causal_df.empty else pd.DataFrame()
    critical_txt = "Nincs kritikus termék"
    if not critical.empty:
        r = critical.iloc[0]
        critical_txt = f"{r['Termék']} ({r['Hiány_db']:.0f} db hiány)"

    data = [
        ["Termelési egészség", f"{health:.1f}/100"],
        ["Rendelésteljesítés", f"{fulfillment_rate:.1f}%"],
        ["Becsült javítási potenciál", fmt_huf(lost)],
        ["Legkritikusabb termék", pdf_safe_text(critical_txt)],
    ]
    t = Table([[Paragraph(str(a), ParagraphStyle("C1", fontSize=9, textColor=colors.HexColor("#475569"))),
                Paragraph(str(b), ParagraphStyle("C2", fontSize=12, textColor=colors.HexColor("#0f172a"), leading=14))] for a,b in data],
              colWidths=[6*cm, 10*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#e5e7eb")),
        ("TOPPADDING", (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
    ]))
    parts.append(t)
    parts.append(Spacer(1, 0.22*cm))

    if action_plan_df is not None and not action_plan_df.empty:
        parts.append(Paragraph("Top 3 teendő", ParagraphStyle("CoverH2", fontSize=12, textColor=colors.HexColor("#1e3a8a"))))
        for _, a in action_plan_df.head(3).iterrows():
            parts.append(make_pdf_action_card(a))
            parts.append(Spacer(1, 0.06*cm))
    return parts


def make_pdf_real_heatmap(matrix: pd.DataFrame, width=520):
    """Valódi színes dolgozó-gép heatmap pontszámokkal, nem emoji négyzetekkel."""
    if matrix is None or matrix.empty:
        return Paragraph("Nincs dolgozó-gép mátrix adat.", ParagraphStyle("Empty", fontSize=8))

    show = matrix.copy().head(9)
    cols = ["Dolgozó"] + list(show.columns)
    data = [cols]

    for idx, row in show.iterrows():
        vals = [str(idx)]
        for v in row.tolist():
            try:
                vals.append(f"{float(v):.0f}")
            except Exception:
                vals.append("")
        data.append(vals)

    first_w = 3.2 * cm
    other_w = max(1.55 * cm, (width - first_w) / max(len(cols) - 1, 1))
    table = Table(data, colWidths=[first_w] + [other_w] * (len(cols) - 1), repeatRows=1)

    style = [
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0f172a")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("FONTSIZE", (0,0), (-1,-1), 7.5),
        ("BACKGROUND", (0,1), (0,-1), colors.HexColor("#f8fafc")),
    ]

    for r_i in range(1, len(data)):
        for c_i in range(1, len(cols)):
            raw = data[r_i][c_i]
            try:
                val = float(raw)
            except Exception:
                val = 0
            if val >= 85:
                bg, fg = "#16a34a", "#ffffff"
            elif val >= 70:
                bg, fg = "#facc15", "#111827"
            elif val >= 55:
                bg, fg = "#fb923c", "#111827"
            elif val > 0:
                bg, fg = "#ef4444", "#ffffff"
            else:
                bg, fg = "#e5e7eb", "#6b7280"
            style.append(("BACKGROUND", (c_i, r_i), (c_i, r_i), colors.HexColor(bg)))
            style.append(("TEXTCOLOR", (c_i, r_i), (c_i, r_i), colors.HexColor(fg)))

    table.setStyle(TableStyle(style))
    return table


def summarize_plan_by_product(plan_df: pd.DataFrame, fulfillment_df: pd.DataFrame = None) -> pd.DataFrame:
    """Vezetői termékszintű tervösszefoglaló a nyers gépsoros lista helyett."""
    if fulfillment_df is not None and not fulfillment_df.empty:
        out = fulfillment_df.copy()
        if "Teljesítés_%" not in out.columns:
            out["Teljesítés_%"] = safe_completion_pct(out["Tervezett_db"], out["Igényelt_db"])
    elif plan_df is not None and not plan_df.empty:
        planned = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])].groupby("Termék", as_index=False).agg(Tervezett_db=("Tervezett_db", "sum"))
        shortage = plan_df[plan_df["Gép"].eq("Kapacitáshiány")].groupby("Termék", as_index=False).agg(Hiány_db=("Tervezett_db", "sum"))
        out = planned.merge(shortage, on="Termék", how="outer").fillna(0)
        out["Igényelt_db"] = out["Tervezett_db"] + out["Hiány_db"]
        out["Teljesítés_%"] = safe_completion_pct(out["Tervezett_db"], out["Igényelt_db"])
    else:
        return pd.DataFrame()

    def risk(row):
        t = float(row.get("Teljesítés_%", 0))
        if t >= 95:
            return "🟢 Rendben"
        if t >= 75:
            return "🟡 Figyelendő"
        return "🔴 Kritikus"
    out["Kockázat"] = out.apply(risk, axis=1)
    return out.sort_values("Teljesítés_%")


def make_pdf_fulfillment_cards(summary_df: pd.DataFrame, width=500):
    """Termékenkénti rendelésteljesítési kártyák."""
    if summary_df is None or summary_df.empty:
        return Paragraph("Nincs rendelésteljesítési adat.", ParagraphStyle("Empty", fontSize=8))

    rows = []
    for _, r in summary_df.head(8).iterrows():
        pct = float(r.get("Teljesítés_%", 0) or 0)
        if pct >= 95:
            color, bg = "#16a34a", "#dcfce7"
        elif pct >= 75:
            color, bg = "#f59e0b", "#fef3c7"
        else:
            color, bg = "#dc2626", "#fee2e2"

        text = (
            f"<b>{pdf_safe_text(r.get('Termék',''))}</b><br/>"
            f"Igény: {fmt_num(r.get('Igényelt_db', 0))} db | "
            f"Tervezett: {fmt_num(r.get('Tervezett_db', 0))} db | "
            f"Hiány: {fmt_num(r.get('Hiány_db', 0))} db<br/>"
            f"<b>Teljesítés:</b> {pct:.1f}%"
        )
        rows.append([
            Paragraph(pdf_safe_text(r.get("Kockázat", "")), ParagraphStyle("Risk", fontSize=8.5, textColor=colors.HexColor(color), alignment=1)),
            Paragraph(text, ParagraphStyle("Fulfill", fontSize=8.5, leading=11, textColor=colors.HexColor("#0f172a"))),
        ])

    table = Table(rows, colWidths=[2.6 * cm, width - 2.6 * cm])
    style = [
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]
    for i, (_, r) in enumerate(summary_df.head(8).iterrows()):
        pct = float(r.get("Teljesítés_%", 0) or 0)
        bg = "#dcfce7" if pct >= 95 else "#fef3c7" if pct >= 75 else "#fee2e2"
        style.append(("BACKGROUND", (0,i), (-1,i), colors.HexColor(bg)))
    table.setStyle(TableStyle(style))
    return table


def make_pdf_capacity_heatmap(capacity_df: pd.DataFrame, width=500):
    if capacity_df is None or capacity_df.empty:
        return Paragraph("Nincs kapacitásadat.", ParagraphStyle("Empty", fontSize=8))

    data = [["Gép", "Kihasználtság", "Státusz"]]
    for _, r in capacity_df.head(10).iterrows():
        data.append([str(r.get("Gép","")), f"{float(r.get('Kihasználtság_%', 0)):.0f}%", str(r.get("Státusz",""))])

    table = Table(data, colWidths=[3.2*cm, 3.2*cm, width - 6.4*cm], repeatRows=1)
    style = [
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0f172a")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("ALIGN", (1,1), (1,-1), "CENTER"),
    ]
    for i, r in enumerate(data[1:], start=1):
        try:
            val = float(str(r[1]).replace("%",""))
        except Exception:
            val = 0
        if val >= 95:
            bg = "#fee2e2"
        elif val >= 75:
            bg = "#fef3c7"
        else:
            bg = "#dcfce7"
        style.append(("BACKGROUND", (0,i), (-1,i), colors.HexColor(bg)))
    table.setStyle(TableStyle(style))
    return table


def make_pdf_symbol_matrix(symbol_df: pd.DataFrame, width=500):
    if symbol_df is None or symbol_df.empty:
        return Paragraph("Nincs dolgozó-gép mátrix adat.", ParagraphStyle("Empty", fontSize=8))
    data=[["Dolgozó"]+list(symbol_df.columns)]
    for idx,row in symbol_df.head(10).iterrows():
        data.append([str(idx)]+[str(v) for v in row.tolist()])
    col_width=width/max(len(data[0]),1)
    t=Table(data, colWidths=[col_width]*len(data[0]))
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0f172a")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#cbd5e1")),("ALIGN",(1,1),(-1,-1),"CENTER"),("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#f8fafc")),("FONTSIZE",(0,0),(-1,-1),8)]))
    return t


def pdf_safe_text(value):
    """ReportLab alap fontokhoz biztonságos, ékezet-kímélő szöveg."""
    return (
        str(value or "")
        .replace("ő", "ö").replace("Ő", "Ö")
        .replace("ű", "ü").replace("Ű", "Ü")
    )


def _pdf_color_for_score(value, inverse=False):
    try:
        v = float(value)
    except Exception:
        v = 0
    if inverse:
        if v <= 35:
            return "#16a34a", "Alacsony"
        if v <= 70:
            return "#f59e0b", "Figyelendő"
        return "#dc2626", "Magas"
    if v >= 75:
        return "#16a34a", "Erős"
    if v >= 50:
        return "#f59e0b", "Közepes"
    return "#dc2626", "Gyenge"


def make_pdf_gauge(title, value, suffix="%", inverse=False, width=165, height=90):
    """Egyszerű, stabil PDF gauge Drawing objektummal."""
    try:
        value = float(value)
    except Exception:
        value = 0
    value = max(0, min(100, value))
    color, label = _pdf_color_for_score(value, inverse=inverse)

    d = Drawing(width, height)
    d.add(Rect(0, 0, width, height, fillColor=colors.HexColor("#f8fafc"), strokeColor=colors.HexColor("#cbd5e1"), rx=10, ry=10))
    d.add(String(10, height - 18, pdf_safe_text(title), fontSize=8, fillColor=colors.HexColor("#334155")))

    # Track
    x0, y0, w, h = 10, 34, width - 20, 12
    d.add(Rect(x0, y0, w, h, fillColor=colors.HexColor("#e5e7eb"), strokeColor=colors.HexColor("#e5e7eb"), rx=5, ry=5))
    d.add(Rect(x0, y0, w * value / 100, h, fillColor=colors.HexColor(color), strokeColor=colors.HexColor(color), rx=5, ry=5))

    d.add(String(10, 14, f"{value:.1f}{suffix}", fontSize=15, fillColor=colors.HexColor(color)))
    d.add(String(78, 17, pdf_safe_text(label), fontSize=8, fillColor=colors.HexColor("#475569")))
    return d


def make_pdf_bar_chart(title, df, label_col, value_col, value_suffix="", width=500, height=175, top_n=8):
    """PDF-be rajzolt egyszerű horizontális bar chart."""
    d = Drawing(width, height)
    d.add(String(0, height - 12, pdf_safe_text(title), fontSize=10, fillColor=colors.HexColor("#0f172a")))

    if df is None or df.empty or label_col not in df.columns or value_col not in df.columns:
        d.add(String(0, height / 2, pdf_safe_text("Nincs adat a diagramhoz."), fontSize=8, fillColor=colors.HexColor("#64748b")))
        return d

    data = df[[label_col, value_col]].dropna().copy().head(top_n)
    if data.empty:
        d.add(String(0, height / 2, pdf_safe_text("Nincs adat a diagramhoz."), fontSize=8, fillColor=colors.HexColor("#64748b")))
        return d

    max_val = max(float(data[value_col].max()), 1)
    chart_top = height - 28
    row_h = min(16, (height - 38) / max(len(data), 1))
    label_w = 110
    bar_w = width - label_w - 70

    for i, (_, r) in enumerate(data.iterrows()):
        y = chart_top - (i + 1) * row_h
        label = pdf_safe_text(str(r[label_col]))[:24]
        val = float(r[value_col])
        bw = bar_w * val / max_val

        d.add(String(0, y + 3, label, fontSize=7, fillColor=colors.HexColor("#334155")))
        d.add(Rect(label_w, y + 2, bar_w, 8, fillColor=colors.HexColor("#e5e7eb"), strokeColor=colors.HexColor("#e5e7eb")))
        d.add(Rect(label_w, y + 2, bw, 8, fillColor=colors.HexColor("#2563eb"), strokeColor=colors.HexColor("#2563eb")))
        d.add(String(label_w + bar_w + 6, y + 2, pdf_safe_text(f"{val:.0f}{value_suffix}"), fontSize=7, fillColor=colors.HexColor("#0f172a")))
    return d


def make_pdf_capacity_chart(capacity_df, width=500, height=165):
    d = Drawing(width, height)
    d.add(String(0, height - 12, "Gépkapacitás kihasználtság", fontSize=10, fillColor=colors.HexColor("#0f172a")))

    if capacity_df is None or capacity_df.empty or "Gép" not in capacity_df.columns or "Kihasználtság_%" not in capacity_df.columns:
        d.add(String(0, height / 2, "Nincs kapacitásadat.", fontSize=8, fillColor=colors.HexColor("#64748b")))
        return d

    data = capacity_df[["Gép", "Kihasználtság_%"]].copy().head(8)
    chart_top = height - 30
    row_h = min(18, (height - 42) / max(len(data), 1))
    label_w = 90
    bar_w = width - label_w - 70

    for i, (_, r) in enumerate(data.iterrows()):
        y = chart_top - (i + 1) * row_h
        val = max(0, min(120, float(r["Kihasználtság_%"])))
        color = "#16a34a" if val < 75 else "#f59e0b" if val < 95 else "#dc2626"

        d.add(String(0, y + 4, pdf_safe_text(r["Gép"]), fontSize=8, fillColor=colors.HexColor("#334155")))
        d.add(Rect(label_w, y + 2, bar_w, 10, fillColor=colors.HexColor("#e5e7eb"), strokeColor=colors.HexColor("#e5e7eb")))
        d.add(Rect(label_w, y + 2, bar_w * min(val, 100) / 100, 10, fillColor=colors.HexColor(color), strokeColor=colors.HexColor(color)))
        d.add(String(label_w + bar_w + 6, y + 2, f"{val:.0f}%", fontSize=8, fillColor=colors.HexColor(color)))
    return d


def pdf_insight_card(text, kind="info", width=500):
    """Színes PDF insight kártya ikonokkal."""
    palette = {
        "danger": ("#fee2e2", "#dc2626", "!!"),
        "warning": ("#fef3c7", "#f59e0b", "!"),
        "success": ("#dcfce7", "#16a34a", "+"),
        "info": ("#dbeafe", "#2563eb", "i"),
    }
    bg, accent, icon = palette.get(kind, palette["info"])
    t = Table(
        [[Paragraph(f"<b>{icon}</b>", ParagraphStyle("IconStyle", fontSize=13, textColor=colors.HexColor(accent))),
          Paragraph(pdf_safe_text(text), ParagraphStyle("CardText", fontSize=8.5, leading=11, textColor=colors.HexColor("#0f172a")))]],
        colWidths=[0.8 * cm, width - 0.8 * cm]
    )
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor(bg)),
        ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor(accent)),
        ("LINEBEFORE", (0,0), (0,-1), 5, colors.HexColor(accent)),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("TOPPADDING", (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
    ]))
    return t



def pdf_section_header(title, subtitle=None):
    elems = []
    elems.append(Paragraph(pdf_safe_text(title), ParagraphStyle("CoolSection", fontSize=15, leading=18, textColor=colors.HexColor("#1e3a8a"))))
    if subtitle:
        elems.append(Paragraph(pdf_safe_text(subtitle), ParagraphStyle("CoolSub", fontSize=8.5, leading=10.5, textColor=colors.HexColor("#475569"))))
    elems.append(Spacer(1, 0.12 * cm))
    return elems


def compact_insight_table(recs, max_items=8):
    rows = []
    icon_map = {"danger": "■", "warning": "▲", "success": "●", "info": "i"}
    color_map = {"danger": "#dc2626", "warning": "#f59e0b", "success": "#16a34a", "info": "#2563eb"}
    for cls, text in (recs or [])[:max_items]:
        rows.append([
            Paragraph(icon_map.get(cls, "i"), ParagraphStyle("Ico", fontSize=11, textColor=colors.HexColor(color_map.get(cls, "#2563eb")), alignment=1)),
            Paragraph(pdf_safe_text(text), ParagraphStyle("InsightSmall", fontSize=8.3, leading=10.5, textColor=colors.HexColor("#0f172a")))
        ])
    if not rows:
        rows = [[Paragraph("i", ParagraphStyle("Ico", fontSize=10)), Paragraph("Nincs megállapítás.", ParagraphStyle("InsightSmall", fontSize=8))]]
    t = Table(rows, colWidths=[0.55 * cm, 16.1 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ("BOX", (0,0), (-1,-1), 0.45, colors.HexColor("#cbd5e1")),
        ("INNERGRID", (0,0), (-1,-1), 0.2, colors.HexColor("#e5e7eb")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    return t


def make_pdf_impact_table(impact_df, width=500):
    if impact_df is None or impact_df.empty:
        return Paragraph("Nincs javítási potenciál adat.", ParagraphStyle("Empty", fontSize=8))
    show = impact_df.head(8).copy()
    show["Becsült_havi_hatás_Ft"] = show["Becsült_havi_hatás_Ft"].apply(fmt_huf)
    data = [["Terület", "Elem", "Probléma", "Becsült hatás", "Javaslat"]]
    for _, r in show.iterrows():
        data.append([
            pdf_safe_text(r.get("Terület", "")),
            pdf_safe_text(r.get("Elem", "")),
            pdf_safe_text(r.get("Probléma", "")),
            pdf_safe_text(r.get("Becsült_havi_hatás_Ft", "")),
            pdf_safe_text(r.get("Javaslat", "")),
        ])
    col_widths = [2.4*cm, 2.8*cm, 4.3*cm, 2.8*cm, 4.3*cm]
    t = Table([[Paragraph(str(c), ParagraphStyle("Tbl", fontSize=7.2, leading=8.5, textColor=colors.HexColor("#0f172a"))) for c in row] for row in data], colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0f172a")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]
    for i in range(1, len(data)):
        bg = "#fee2e2" if i == 1 else "#fff7ed" if i <= 3 else "#f8fafc"
        style.append(("BACKGROUND", (0,i), (-1,i), colors.HexColor(bg)))
    t.setStyle(TableStyle(style))
    return t


def make_pdf_top_pairs_table(pair, width=500):
    if pair is None or pair.empty:
        return Paragraph("Nincs dolgozó-gép páros adat.", ParagraphStyle("Empty", fontSize=8))
    top = pair.sort_values("Kompatibilitási_pont", ascending=False).head(8).copy()
    data = [["Páros", "Pont", "Teljesítmény", "Selejt"]]
    for _, r in top.iterrows():
        data.append([
            pdf_safe_text(f"{r['Dolgozó']} - {r['Gép']}"),
            f"{float(r['Kompatibilitási_pont']):.0f}",
            f"{float(r['Átlag_teljesítmény']):.1f}%",
            f"{float(r['Selejt_%']):.1f}%"
        ])
    t = Table([[Paragraph(str(c), ParagraphStyle("PairTbl", fontSize=8, leading=9.5, textColor=colors.HexColor("#0f172a"))) for c in row] for row in data],
              colWidths=[7.4*cm, 2.4*cm, 3.4*cm, 3.4*cm], repeatRows=1)
    style = [
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0f172a")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
    ]
    for i in range(1, len(data)):
        style.append(("BACKGROUND", (0,i), (-1,i), colors.HexColor("#dcfce7" if i <= 3 else "#f8fafc")))
    t.setStyle(TableStyle(style))
    return t


def build_pdf_report(
    df: pd.DataFrame,
    pair: pd.DataFrame,
    recs: List[Tuple[str, str]],
    assignment: pd.DataFrame = None,
    plan_df: pd.DataFrame = None,
    worker_plan: pd.DataFrame = None,
    orders_df: pd.DataFrame = None,
    fulfillment_df: pd.DataFrame = None,
    capacity_df: pd.DataFrame = None,
    plan_recs: List[Tuple[str, str]] = None,
    root_cause_recs: List[Tuple[str, str]] = None,
    impact_df: pd.DataFrame = None,
    advisor_scores: Dict[str, float] = None,
    action_plan_df: pd.DataFrame = None,
    symbol_matrix: pd.DataFrame = None,
    causal_chain_df: pd.DataFrame = None,
    lost_revenue_df: pd.DataFrame = None,
    critical_orders_df: pd.DataFrame = None
) -> bytes:
    """PRO: rövidebb, informatívabb, kevésbé szöveges vezetői PDF."""
    if SimpleDocTemplate is None:
        return None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.0 * cm,
        leftMargin=1.0 * cm,
        topMargin=0.85 * cm,
        bottomMargin=0.85 * cm,
    )

    title_style = ParagraphStyle("TitleHU", fontSize=18, leading=22, textColor=colors.HexColor("#0f172a"))
    body = ParagraphStyle("BodyHU", fontSize=8.2, leading=10.2, textColor=colors.HexColor("#334155"))
    h2 = ParagraphStyle("H2HU", fontSize=13, leading=16, textColor=colors.HexColor("#1e3a8a"))

    def P(x, style=body):
        return Paragraph(pdf_safe_text(x), style)

    total_qty = df["Gyártott_db"].sum()
    downtime = df["Állásidő_perc"].sum()
    fedezet = df["Becsült_fedezet"].sum()
    fedezet_m = fedezet / 1_000_000

    story = []
    story.append(Paragraph("Gyártási Diagnosztika PRO SaaS V4 V4.1 V2 SaaS.7 - vezetői riport", title_style))
    story.append(P("Rövid döntéstámogató riport: fő megállapítások, javítási potenciál, dolgozó-gép párosítások."))
    story.append(Spacer(1, 0.20 * cm))

    # Compact KPI row only, no gauges
    kpi_header_style = ParagraphStyle("KPIHeader", fontSize=8.5, leading=10, textColor=colors.white, alignment=1)
    kpi_value_style = ParagraphStyle("KPIValue", fontSize=9, leading=11, textColor=colors.HexColor("#0f172a"), alignment=1)
    kpi_data = [
        [Paragraph("Gyártott db", kpi_header_style), Paragraph("Állásidő", kpi_header_style), Paragraph("Becsült fedezet", kpi_header_style), Paragraph("Fedezet millió Ft", kpi_header_style)],
        [Paragraph(fmt_num(total_qty), kpi_value_style), Paragraph(f"{fmt_num(downtime)} perc", kpi_value_style), Paragraph(fmt_huf(fedezet), kpi_value_style), Paragraph(f"{fedezet_m:.1f} M Ft", kpi_value_style)],
    ]
    table = Table(kpi_data, colWidths=[4.2 * cm] * 4)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1e3a8a")),
        ("BACKGROUND", (0,1), (-1,1), colors.HexColor("#eff6ff")),
        ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#cbd5e1")),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("TOPPADDING", (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
    ]))
    story.append(table)
    story.append(Spacer(1, 0.22 * cm))

    # Keep sections user liked
    story.extend(pdf_section_header("Fő megállapítások"))
    story.append(compact_insight_table((recs or [])[:4]))
    story.append(Spacer(1, 0.18 * cm))

    story.extend(pdf_section_header("Gyökérok és pénzügyi fókusz"))
    story.append(compact_insight_table((root_cause_recs or [])[:4]))
    story.append(Spacer(1, 0.18 * cm))

    story.extend(pdf_section_header("Becsült javítási potenciál", "Top javítási lehetőségek havi becsült értékkel."))
    if impact_df is not None and not impact_df.empty:
        story.append(make_pdf_bar_chart("Top potenciál", impact_df.head(5), "Elem", "Becsült_havi_hatás_Ft", " Ft", width=520, height=135, top_n=5))
        story.append(Spacer(1, 0.10 * cm))
    story.append(make_pdf_impact_table(impact_df, width=520))
    story.append(Spacer(1, 0.22 * cm))

    story.extend(pdf_section_header("Dolgozó-gép hőtérkép", "Pontszám: zöld = kiemelkedő, sárga = jó, narancs = fejleszthető, piros = kerülendő."))
    if pair is not None and not pair.empty:
        try:
            heat_matrix = normalized_pair_score_table(pair)
            story.append(make_pdf_real_heatmap(heat_matrix, width=520))
        except Exception:
            story.append(P("A hőtérkép nem készült el."))
    story.append(Spacer(1, 0.18 * cm))

    story.extend(pdf_section_header("Top dolgozó-gép párosok"))
    story.append(make_pdf_top_pairs_table(pair, width=520))
    story.append(Spacer(1, 0.18 * cm))

    # Optional concise action list, max 3, only if meaningful
    if action_plan_df is not None and not action_plan_df.empty:
        story.extend(pdf_section_header("Top 3 akció"))
        for _, act in action_plan_df.head(3).iterrows():
            story.append(make_pdf_action_card(act))
            story.append(Spacer(1, 0.06 * cm))

    doc.build(story)
    return buffer.getvalue()


def optimized_assignment(
    pair: pd.DataFrame,
    unavailable_workers: List[str] = None,
    unavailable_machines: List[str] = None,
    one_worker_once: bool = True
) -> pd.DataFrame:
    """V2 optimalizáló: dolgozó/gép kizárás és egyszerű greedy beosztás.

    Cél: minden elérhető gépre a lehető legjobb dolgozó-gép párosítás,
    opcionálisan úgy, hogy egy dolgozó csak egyszer szerepelhet.
    """
    unavailable_workers = unavailable_workers or []
    unavailable_machines = unavailable_machines or []

    available = pair[
        ~pair["Dolgozó"].isin(unavailable_workers)
        & ~pair["Gép"].isin(unavailable_machines)
    ].copy()

    if available.empty:
        return pd.DataFrame(columns=[
            "Gép", "Ajánlott dolgozó", "Kompatibilitási_pont",
            "Várható teljesítmény_%", "Várható selejt_%", "Várható fedezet/db"
        ])

    assignments = []
    used_workers = set()

    # A legfontosabb gépekkel kezdünk: ahol magasabb átlagfedezet/db vagy teljesítmény látszik.
    machine_priority = (
        available.groupby("Gép", as_index=False)
        .agg(
            Átlag_pont=("Kompatibilitási_pont", "mean"),
            Átlag_fedezet=("Fedezet/db", "mean"),
            Sorok=("Sorok", "sum")
        )
        .sort_values(["Átlag_fedezet", "Átlag_pont"], ascending=False)
    )

    for machine in machine_priority["Gép"].tolist():
        candidates = available[available["Gép"] == machine].sort_values("Kompatibilitási_pont", ascending=False)
        if one_worker_once:
            candidates = candidates[~candidates["Dolgozó"].isin(used_workers)]
        if candidates.empty:
            continue

        best = candidates.iloc[0]
        used_workers.add(best["Dolgozó"])
        assignments.append({
            "Gép": best["Gép"],
            "Ajánlott dolgozó": best["Dolgozó"],
            "Kompatibilitási_pont": round(best["Kompatibilitási_pont"], 1),
            "Várható teljesítmény_%": round(best["Átlag_teljesítmény"], 1),
            "Várható selejt_%": round(best["Selejt_%"], 2),
            "Várható fedezet/db": round(best["Fedezet/db"], 0),
        })

    return pd.DataFrame(assignments).sort_values("Gép")


def compare_assignment_scenarios(pair: pd.DataFrame, current_assignment: pd.DataFrame, optimized: pd.DataFrame) -> pd.DataFrame:
    """Egyszerű összehasonlítás a mostani V1 és optimalizált beosztás között."""
    if optimized is None or optimized.empty:
        return pd.DataFrame([{
            "Mutató": "Optimalizált beosztás",
            "Érték": "Nincs elég adat / túl sok kizárás"
        }])

    current_score = current_assignment["Kompatibilitási_pont"].mean() if not current_assignment.empty else np.nan
    opt_score = optimized["Kompatibilitási_pont"].mean()
    current_fedezet = current_assignment["Fedezet/db"].mean() if "Fedezet/db" in current_assignment.columns and not current_assignment.empty else np.nan
    opt_fedezet = optimized["Várható fedezet/db"].mean()

    rows = [
        {"Mutató": "Átlag kompatibilitási pont", "Jelenlegi egyszerű ajánlás": round(current_score, 1), "Optimalizált": round(opt_score, 1), "Változás": round(opt_score - current_score, 1) if pd.notna(current_score) else "-"},
        {"Mutató": "Átlag fedezet/db", "Jelenlegi egyszerű ajánlás": round(current_fedezet, 0) if pd.notna(current_fedezet) else "-", "Optimalizált": round(opt_fedezet, 0), "Változás": round(opt_fedezet - current_fedezet, 0) if pd.notna(current_fedezet) else "-"},
        {"Mutató": "Beosztott gépek száma", "Jelenlegi egyszerű ajánlás": len(current_assignment), "Optimalizált": len(optimized), "Változás": len(optimized) - len(current_assignment)},
    ]
    return pd.DataFrame(rows)




def normalize_orders(orders_raw: pd.DataFrame) -> pd.DataFrame:
    """Opcionális Megrendelesek munkalap feldolgozása."""
    if orders_raw is None or orders_raw.empty:
        return pd.DataFrame(columns=OPTIONAL_ORDER_COLS)

    orders = orders_raw.copy()
    missing = [c for c in OPTIONAL_ORDER_COLS if c not in orders.columns]
    if missing:
        # Nem állítjuk meg az appot, csak üres rendelésállományként kezeljük.
        return pd.DataFrame(columns=OPTIONAL_ORDER_COLS)

    orders["Rendelt_db"] = pd.to_numeric(orders["Rendelt_db"], errors="coerce").fillna(0).astype(int)
    orders["Határidő"] = pd.to_datetime(orders["Határidő"], errors="coerce")
    orders["Prioritás"] = pd.to_numeric(orders["Prioritás"], errors="coerce").fillna(3).astype(int)
    return orders


def demand_from_orders(orders_df: pd.DataFrame) -> Dict[str, int]:
    """Rendelésállományból termékenkénti igény."""
    if orders_df is None or orders_df.empty:
        return {}
    d = orders_df.groupby("Termék")["Rendelt_db"].sum().to_dict()
    return {str(k): int(v) for k, v in d.items()}


def order_priority_view(orders_df: pd.DataFrame) -> pd.DataFrame:
    """Rendelések priorizált nézete."""
    if orders_df is None or orders_df.empty:
        return pd.DataFrame()
    out = orders_df.copy()
    today = pd.Timestamp.today().normalize()
    out["Napok_határidőig"] = (out["Határidő"] - today).dt.days
    out["Sürgősségi_pont"] = (
        (6 - out["Prioritás"].clip(1, 5)) * 20
        + np.where(out["Napok_határidőig"] <= 3, 40, 0)
        + np.where(out["Napok_határidőig"] <= 7, 20, 0)
    )
    return out.sort_values(["Sürgősségi_pont", "Határidő"], ascending=[False, True])


def build_order_fulfillment(plan_df: pd.DataFrame, orders_df: pd.DataFrame) -> pd.DataFrame:
    """Megmutatja, hogy a javasolt gyártási terv mennyire fedezi a rendelésállományt."""
    if orders_df is None or orders_df.empty:
        return pd.DataFrame()

    demand = orders_df.groupby("Termék", as_index=False).agg(Rendelt_db=("Rendelt_db", "sum"))
    if plan_df is None or plan_df.empty:
        demand["Tervezett_db"] = 0
    else:
        planned = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])].groupby("Termék", as_index=False).agg(
            Tervezett_db=("Tervezett_db", "sum")
        )
        demand = demand.merge(planned, on="Termék", how="left")
        demand["Tervezett_db"] = demand["Tervezett_db"].fillna(0)

    demand["Hiány_db"] = (demand["Rendelt_db"] - demand["Tervezett_db"]).clip(lower=0)
    demand["Teljesítés_%"] = safe_completion_pct(demand["Tervezett_db"], demand["Rendelt_db"])
    return demand.sort_values("Teljesítés_%")


def generate_order_insights(orders_df: pd.DataFrame, fulfillment_df: pd.DataFrame) -> List[Tuple[str, str]]:
    if orders_df is None or orders_df.empty:
        return [("warning", "Nincs Megrendelesek munkalap, ezért a terv kézi darabszámokból indul.")]

    recs = []
    total_orders = int(orders_df["Rendelt_db"].sum())
    recs.append(("success", f"A feltöltött rendelésállomány összesen {fmt_num(total_orders)} db gyártási igényt tartalmaz."))

    urgent = order_priority_view(orders_df)
    if not urgent.empty:
        top = urgent.iloc[0]
        recs.append(("warning", f"Legsürgősebb rendelés: {top['Rendelés_ID']} / {top['Vevő']} / {top['Termék']} / {fmt_num(top['Rendelt_db'])} db."))

    if fulfillment_df is not None and not fulfillment_df.empty:
        worst = fulfillment_df.sort_values("Teljesítés_%").iloc[0]
        if worst["Teljesítés_%"] < 100:
            recs.append(("danger", f"Rendelésteljesítési hiány: {worst['Termék']} termékből {fmt_num(worst['Hiány_db'])} db még nem fér bele a tervbe."))
        else:
            recs.append(("success", "A jelenlegi terv termékszinten fedezi a rendelésállományt."))

    return recs


def product_machine_priority(df: pd.DataFrame) -> pd.DataFrame:
    """Termék-gép prioritási tábla: melyik termék melyik gépen hozza a legtöbb értéket."""
    if df.empty:
        return pd.DataFrame()

    out = df.groupby(["Termék", "Gép"], as_index=False).agg(
        Gyártott_db=("Gyártott_db", "sum"),
        Jó_db=("Jó_db", "sum"),
        Selejt_db=("Selejt_db", "sum"),
        Átlag_OEE=("OEE_light_%", "mean"),
        Átlag_teljesítmény=("Teljesítmény_%", "mean"),
        Fedezet=("Becsült_fedezet", "sum"),
        Kapacitás_db_óra=("Kapacitás_db_óra", "mean"),
        Sorok=("Gyártott_db", "count")
    )
    out["Selejt_%"] = np.where(out["Gyártott_db"] > 0, out["Selejt_db"] / out["Gyártott_db"] * 100, 0)
    out["Fedezet/db"] = np.where(out["Jó_db"] > 0, out["Fedezet"] / out["Jó_db"], 0)

    # V4 prioritás: fedezet/db + OEE + alacsony selejt
    fedezet = out["Fedezet/db"]
    if fedezet.max() != fedezet.min():
        fedezet_score = (fedezet - fedezet.min()) / (fedezet.max() - fedezet.min()) * 55
    else:
        fedezet_score = 27.5

    oee_score = out["Átlag_OEE"].clip(0, 100) / 100 * 30
    quality_score = (100 - out["Selejt_%"].clip(0, 20) * 5).clip(0, 100) / 100 * 15
    out["Termék_gép_pont"] = (fedezet_score + oee_score + quality_score).round(1)
    return out.sort_values("Termék_gép_pont", ascending=False)



def build_order_level_plan(
    df: pd.DataFrame,
    orders_df: pd.DataFrame,
    manual_demand: Dict[str, int],
    planning_days: int = 5,
    hours_per_machine_day: float = 8.0,
    unavailable_machines: List[str] = None
) -> pd.DataFrame:
    """PRO.4.3.2: rendelésalapú gyártási terv.

    A Tervezett_db nem önálló becslés: az Igényelt_db-ből indul,
    majd a tervezési horizont, a gépórák, a gépenkénti kapacitás és a
    múltbeli termék-gép teljesítmény alapján korlátozódik.
    """
    unavailable_machines = unavailable_machines or []
    priority = product_machine_priority(df)
    if priority.empty:
        return pd.DataFrame()

    # Kapacitás oszlopok biztosítása
    if "Kapacitás_db_óra" not in priority.columns:
        cap = df.groupby("Gép")["Kapacitás_db_óra"].mean().reset_index()
        priority = priority.merge(cap, on="Gép", how="left")

    priority = priority[~priority["Gép"].isin(unavailable_machines)].copy()
    if priority.empty:
        return pd.DataFrame()

    machines = df[["Gép", "Kapacitás_db_óra", "Elérhető_óra_nap"]].drop_duplicates("Gép")
    available_hours = {}
    for _, r in machines.iterrows():
        machine = r["Gép"]
        if machine in unavailable_machines:
            continue
        daily_hours = min(float(hours_per_machine_day), float(r.get("Elérhető_óra_nap", hours_per_machine_day)))
        available_hours[machine] = max(0.0, daily_hours * float(planning_days))

    if orders_df is not None and not orders_df.empty:
        order_rows = order_priority_view(orders_df).copy()
        order_rows["Igényelt_db"] = order_rows["Rendelt_db"]
    else:
        order_rows = pd.DataFrame([
            {
                "Rendelés_ID": f"MANUAL-{product}",
                "Vevő": "Kézi igény",
                "Termék": product,
                "Rendelt_db": int(qty or 0),
                "Igényelt_db": int(qty or 0),
                "Határidő": pd.NaT,
                "Prioritás": 3,
                "Sürgősségi_pont": 0
            }
            for product, qty in manual_demand.items()
            if int(qty or 0) > 0
        ])

    rows = []
    for _, order in order_rows.iterrows():
        product = order["Termék"]
        requested = int(order.get("Igényelt_db", order.get("Rendelt_db", 0)) or 0)
        remaining = requested
        if remaining <= 0:
            continue

        candidates = priority[priority["Termék"] == product].sort_values("Termék_gép_pont", ascending=False)

        if candidates.empty:
            rows.append({
                "Rendelés_ID": order.get("Rendelés_ID", ""),
                "Vevő": order.get("Vevő", ""),
                "Termék": product,
                "Gép": "Nincs adat",
                "Igényelt_db": requested,
                "Tervezett_db": 0,
                "Hiány_db": requested,
                "Becsült_óra": 0,
                "Becsült_fedezet": 0,
                "Határidő": order.get("Határidő", pd.NaT),
                "Prioritás": order.get("Prioritás", 3),
                "Megjegyzés": "Nincs múltbeli termék-gép adat"
            })
            continue

        for _, cand in candidates.iterrows():
            if remaining <= 0:
                break

            machine = cand["Gép"]
            free_hours = float(available_hours.get(machine, 0))
            if free_hours <= 0:
                continue

            avg_per_hour = max(
                float(cand.get("Átlag_teljesítmény", 0)) / 100 * float(cand.get("Kapacitás_db_óra", 0)),
                1.0
            )
            possible_qty = int(avg_per_hour * free_hours)
            if possible_qty <= 0:
                continue

            planned_qty = min(remaining, possible_qty)
            used_hours = planned_qty / avg_per_hour if avg_per_hour else 0
            available_hours[machine] = max(0.0, free_hours - used_hours)
            remaining -= planned_qty

            rows.append({
                "Rendelés_ID": order.get("Rendelés_ID", ""),
                "Vevő": order.get("Vevő", ""),
                "Termék": product,
                "Gép": machine,
                "Igényelt_db": requested,
                "Tervezett_db": int(planned_qty),
                "Hiány_db": 0,
                "Becsült_óra": round(used_hours, 2),
                "Becsült_fedezet": round(float(planned_qty) * float(cand.get("Fedezet/db", 0)), 0),
                "Határidő": order.get("Határidő", pd.NaT),
                "Prioritás": order.get("Prioritás", 3),
                "Megjegyzés": f"Pont: {cand.get('Termék_gép_pont', 0):.0f}; maradék gépóra: {available_hours[machine]:.1f}"
            })

        if remaining > 0:
            rows.append({
                "Rendelés_ID": order.get("Rendelés_ID", ""),
                "Vevő": order.get("Vevő", ""),
                "Termék": product,
                "Gép": "Kapacitáshiány",
                "Igényelt_db": requested,
                "Tervezett_db": 0,
                "Hiány_db": int(remaining),
                "Becsült_óra": 0,
                "Becsült_fedezet": 0,
                "Határidő": order.get("Határidő", pd.NaT),
                "Prioritás": order.get("Prioritás", 3),
                "Megjegyzés": "Nem fér bele a megadott tervezési horizontba / gépórába"
            })

    return pd.DataFrame(rows)


def build_order_fulfillment_v7(plan_df: pd.DataFrame, orders_df: pd.DataFrame = None, manual_demand: Dict[str, int] = None) -> pd.DataFrame:
    """Rendelés/igény teljesítés termékszinten, PRO.4.3.2 logikával."""
    if plan_df is None or plan_df.empty:
        return pd.DataFrame()

    active = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])].copy()
    planned = active.groupby("Termék", as_index=False).agg(Tervezett_db=("Tervezett_db", "sum")) if not active.empty else pd.DataFrame(columns=["Termék", "Tervezett_db"])

    if orders_df is not None and not orders_df.empty:
        demand = orders_df.groupby("Termék", as_index=False).agg(Igényelt_db=("Rendelt_db", "sum"))
    else:
        demand = pd.DataFrame([
            {"Termék": k, "Igényelt_db": int(v or 0)}
            for k, v in (manual_demand or {}).items()
            if int(v or 0) > 0
        ])

    if demand.empty:
        return pd.DataFrame()

    out = demand.merge(planned, on="Termék", how="left")
    out["Tervezett_db"] = out["Tervezett_db"].fillna(0)
    out["Hiány_db"] = (out["Igényelt_db"] - out["Tervezett_db"]).clip(lower=0)
    out["Teljesítés_%"] = safe_completion_pct(out["Tervezett_db"], out["Igényelt_db"])
    return out.sort_values("Teljesítés_%")


def build_capacity_gap_v7(plan_df: pd.DataFrame, planning_days: int, hours_per_machine_day: float) -> pd.DataFrame:
    """Gépenkénti kapacitás a tervezési horizont alapján."""
    if plan_df is None or plan_df.empty:
        return pd.DataFrame(columns=["Gép", "Tervezett_óra", "Max_óra", "Kihasználtság_%", "Státusz"])

    active = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])].copy()
    if active.empty:
        return pd.DataFrame(columns=["Gép", "Tervezett_óra", "Max_óra", "Kihasználtság_%", "Státusz"])

    out = active.groupby("Gép", as_index=False).agg(
        Tervezett_óra=("Becsült_óra", "sum"),
        Becsült_fedezet=("Becsült_fedezet", "sum")
    )
    out["Max_óra"] = float(planning_days) * float(hours_per_machine_day)
    out["Kihasználtság_%"] = np.where(out["Max_óra"] > 0, out["Tervezett_óra"] / out["Max_óra"] * 100, 0).round(1)

    def status(x):
        if x >= 95:
            return "Szűk keresztmetszet / teljesen lekötött"
        if x >= 75:
            return "Magas kihasználtság"
        if x >= 40:
            return "Kiegyensúlyozott"
        return "Alulterhelt / van szabad kapacitás"

    out["Státusz"] = out["Kihasználtság_%"].apply(status)
    return out.sort_values("Kihasználtság_%", ascending=False)


def generate_plan_insights_v7(plan_df: pd.DataFrame, fulfillment_df: pd.DataFrame, capacity_df: pd.DataFrame) -> List[Tuple[str, str]]:
    recs = []
    if plan_df is None or plan_df.empty:
        return [("warning", "Nincs gyártási terv. Adj meg rendelési igényt vagy tölts fel Megrendelesek munkalapot.")]

    active = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])].copy()
    total_planned = active["Tervezett_db"].sum() if not active.empty else 0
    total_fedezet = active["Becsült_fedezet"].sum() if not active.empty else 0
    recs.append(("success", f"A PRO.4.3.2 terv {fmt_num(total_planned)} db gyártást és kb. {fmt_huf(total_fedezet)} becsült fedezetot mutat."))

    if fulfillment_df is not None and not fulfillment_df.empty:
        shortage = fulfillment_df["Hiány_db"].sum()
        if shortage > 0:
            recs.append(("danger", f"Kapacitáshiány: {fmt_num(shortage)} db igény nem fér bele a megadott horizontba/gépórába."))
        else:
            recs.append(("success", "A jelenlegi terv termékszinten fedezi a rendelésállományt."))

    if capacity_df is not None and not capacity_df.empty:
        bottleneck = capacity_df.iloc[0]
        recs.append(("warning", f"Szűk keresztmetszet jelölt: {bottleneck['Gép']} ({bottleneck['Kihasználtság_%']:.1f}% kihasználtság)."))

    return recs

def build_production_plan(
    df: pd.DataFrame,
    demand: Dict[str, int],
    max_hours_per_machine: float = 8.0,
    unavailable_machines: List[str] = None
) -> pd.DataFrame:
    """Egyszerű greedy gyártási terv.
    A legjobb termék-gép párosoktól indul, figyeli a gépenkénti órakeretet.
    """
    unavailable_machines = unavailable_machines or []
    priority = product_machine_priority(df)
    if priority.empty:
        return pd.DataFrame()

    priority = priority[~priority["Gép"].isin(unavailable_machines)].copy()
    if priority.empty:
        return pd.DataFrame()

    machine_hours = {m: 0.0 for m in priority["Gép"].unique()}
    rows = []

    for product, qty_needed in demand.items():
        remaining = int(qty_needed or 0)
        if remaining <= 0:
            continue

        candidates = priority[priority["Termék"] == product].sort_values("Termék_gép_pont", ascending=False)
        if candidates.empty:
            rows.append({
                "Termék": product,
                "Gép": "Nincs adat",
                "Tervezett_db": 0,
                "Becsült_óra": 0,
                "Becsült_fedezet": 0,
                "Megjegyzés": "Nincs múltbeli adat ehhez a termékhez"
            })
            continue

        for _, cand in candidates.iterrows():
            if remaining <= 0:
                break

            machine = cand["Gép"]
            avg_per_hour = max(float(cand["Átlag_teljesítmény"]) / 100 * float(df[df["Gép"] == machine]["Kapacitás_db_óra"].mean()), 1)
            free_hours = max_hours_per_machine - machine_hours.get(machine, 0.0)
            if free_hours <= 0:
                continue

            possible_qty = int(avg_per_hour * free_hours)
            planned_qty = min(remaining, possible_qty)
            used_hours = planned_qty / avg_per_hour if avg_per_hour else 0
            machine_hours[machine] = machine_hours.get(machine, 0.0) + used_hours
            remaining -= planned_qty

            rows.append({
                "Termék": product,
                "Gép": machine,
                "Tervezett_db": planned_qty,
                "Becsült_óra": round(used_hours, 2),
                "Becsült_fedezet": round(planned_qty * float(cand["Fedezet/db"]), 0),
                "Megjegyzés": f"Prioritási pont: {cand['Termék_gép_pont']:.0f}"
            })

        if remaining > 0:
            rows.append({
                "Termék": product,
                "Gép": "Kapacitáshiány",
                "Tervezett_db": remaining,
                "Becsült_óra": 0,
                "Becsült_fedezet": 0,
                "Megjegyzés": "Nem fér bele a megadott gépórákba"
            })

    return pd.DataFrame(rows)



def build_worker_machine_plan(plan_df: pd.DataFrame, pair: pd.DataFrame, unavailable_workers: List[str] = None) -> pd.DataFrame:
    """A gyártási terv gépeihez dolgozót ajánl.
    Egyszerű greedy: gépenként a legjobb elérhető dolgozót választja.
    """
    unavailable_workers = unavailable_workers or []
    if plan_df is None or plan_df.empty or pair is None or pair.empty:
        return pd.DataFrame(columns=["Gép", "Termék", "Tervezett_db", "Ajánlott_dolgozó", "Dolgozó-gép_pont", "Megjegyzés"])

    active_plan = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])].copy()
    if active_plan.empty:
        return pd.DataFrame(columns=["Gép", "Termék", "Tervezett_db", "Ajánlott_dolgozó", "Dolgozó-gép_pont", "Megjegyzés"])

    rows = []
    used_workers = set()

    for _, task in active_plan.sort_values(["Becsült_fedezet", "Tervezett_db"], ascending=False).iterrows():
        machine = task["Gép"]
        candidates = pair[
            (pair["Gép"] == machine)
            & (~pair["Dolgozó"].isin(unavailable_workers))
            & (~pair["Dolgozó"].isin(used_workers))
        ].sort_values("Kompatibilitási_pont", ascending=False)

        if candidates.empty:
            # fallback: allow repeated worker if no unused candidate
            candidates = pair[
                (pair["Gép"] == machine)
                & (~pair["Dolgozó"].isin(unavailable_workers))
            ].sort_values("Kompatibilitási_pont", ascending=False)

        if candidates.empty:
            rows.append({
                "Gép": machine,
                "Termék": task["Termék"],
                "Tervezett_db": task["Tervezett_db"],
                "Ajánlott_dolgozó": "Nincs elérhető adat",
                "Dolgozó-gép_pont": 0,
                "Megjegyzés": "Nincs múltbeli dolgozó-gép adat"
            })
            continue

        best = candidates.iloc[0]
        used_workers.add(best["Dolgozó"])

        rows.append({
            "Gép": machine,
            "Termék": task["Termék"],
            "Tervezett_db": task["Tervezett_db"],
            "Ajánlott_dolgozó": best["Dolgozó"],
            "Dolgozó-gép_pont": round(best["Kompatibilitási_pont"], 1),
            "Megjegyzés": f"Várható teljesítmény: {best['Átlag_teljesítmény']:.1f}%, selejt: {best['Selejt_%']:.1f}%"
        })

    return pd.DataFrame(rows).sort_values(["Gép", "Termék"])


def build_capacity_gap(plan_df: pd.DataFrame, max_hours_per_machine: float) -> pd.DataFrame:
    """Gépenkénti kihasználtság / kapacitáshiány V5."""
    if plan_df is None or plan_df.empty:
        return pd.DataFrame(columns=["Gép", "Tervezett_óra", "Max_óra", "Kihasználtság_%", "Státusz"])

    active = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])].copy()
    if active.empty:
        return pd.DataFrame(columns=["Gép", "Tervezett_óra", "Max_óra", "Kihasználtság_%", "Státusz"])

    out = active.groupby("Gép", as_index=False).agg(
        Tervezett_óra=("Becsült_óra", "sum"),
        Becsült_fedezet=("Becsült_fedezet", "sum")
    )
    out["Max_óra"] = max_hours_per_machine
    out["Kihasználtság_%"] = np.where(out["Max_óra"] > 0, out["Tervezett_óra"] / out["Max_óra"] * 100, 0).round(1)

    def status(x):
        if x >= 95:
            return "Szűk keresztmetszet / teljesen lekötött"
        if x >= 75:
            return "Magas kihasználtság"
        if x >= 40:
            return "Kiegyensúlyozott"
        return "Alulterhelt / van szabad kapacitás"

    out["Státusz"] = out["Kihasználtság_%"].apply(status)
    return out.sort_values("Kihasználtság_%", ascending=False)


def build_scenario_summary(plan_df: pd.DataFrame, worker_plan: pd.DataFrame, capacity_df: pd.DataFrame) -> List[Tuple[str, str]]:
    recs = []
    if plan_df is None or plan_df.empty:
        return [("warning", "Nincs gyártási terv a szcenárióhoz.")]

    total_qty = plan_df["Tervezett_db"].sum() if "Tervezett_db" in plan_df.columns else 0
    total_fedezet = plan_df["Becsült_fedezet"].sum() if "Becsült_fedezet" in plan_df.columns else 0
    recs.append(("success", f"A V7 terv {fmt_num(total_qty)} db gyártást és kb. {fmt_huf(total_fedezet)} becsült fedezetot mutat."))

    shortage = plan_df[plan_df["Gép"].eq("Kapacitáshiány")]
    if not shortage.empty:
        recs.append(("danger", f"Kapacitáshiány: {fmt_num(shortage['Tervezett_db'].sum())} db nem fér bele. Növeld a gépórát, vagy csökkentsd a kieső gépeket."))

    if capacity_df is not None and not capacity_df.empty:
        bottleneck = capacity_df.iloc[0]
        recs.append(("warning", f"Szűk keresztmetszet jelölt: {bottleneck['Gép']} ({bottleneck['Kihasználtság_%']:.1f}% kihasználtság)."))

        low = capacity_df[capacity_df["Kihasználtság_%"] < 40]
        if not low.empty:
            recs.append(("success", f"Van szabad kapacitás: {', '.join(low['Gép'].astype(str).tolist()[:3])}."))

    if worker_plan is not None and not worker_plan.empty:
        weak = worker_plan[worker_plan["Dolgozó-gép_pont"] < 55]
        if not weak.empty:
            recs.append(("warning", f"{len(weak)} beosztási pont gyengébb kompatibilitású. Itt képzés vagy másik dolgozó megfontolandó."))

    return recs


def build_excel_report(df: pd.DataFrame, pair: pd.DataFrame, assignment: pd.DataFrame, plan_df: pd.DataFrame = None, worker_plan: pd.DataFrame = None, orders_df: pd.DataFrame = None, fulfillment_df: pd.DataFrame = None, capacity_df: pd.DataFrame = None, impact_df: pd.DataFrame = None) -> bytes:
    """Letölthető Excel riport több munkalappal."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        aggregate_metrics(df, ["Műszak"]).to_excel(writer, sheet_name="Muszakok", index=False)
        aggregate_metrics(df, ["Gép"]).to_excel(writer, sheet_name="Gepek", index=False)
        aggregate_metrics(df, ["Dolgozó"]).to_excel(writer, sheet_name="Dolgozok", index=False)
        aggregate_metrics(df, ["Termék"]).to_excel(writer, sheet_name="Termekek", index=False)
        pair.sort_values("Kompatibilitási_pont", ascending=False).to_excel(writer, sheet_name="Dolgozo_gep_parok", index=False)
        assignment.to_excel(writer, sheet_name="Javasolt_beosztas", index=False)
        if plan_df is not None and not plan_df.empty:
            plan_df.to_excel(writer, sheet_name="Gyartasi_terv", index=False)
        if worker_plan is not None and not worker_plan.empty:
            worker_plan.to_excel(writer, sheet_name="Dolgozoi_terv", index=False)
        if orders_df is not None and not orders_df.empty:
            orders_df.to_excel(writer, sheet_name="Megrendelesek", index=False)
        if fulfillment_df is not None and not fulfillment_df.empty:
            fulfillment_df.to_excel(writer, sheet_name="Rendeles_teljesites", index=False)
        if capacity_df is not None and not capacity_df.empty:
            capacity_df.to_excel(writer, sheet_name="Kapacitas", index=False)
        if impact_df is not None and not impact_df.empty:
            impact_df.to_excel(writer, sheet_name="Koltseg_hatas", index=False)
    return output.getvalue()


def generate_plan_insights(plan_df: pd.DataFrame) -> List[Tuple[str, str]]:
    recs = []
    if plan_df is None or plan_df.empty:
        return [("warning", "Nincs még gyártási terv. Adj meg rendelési mennyiségeket.")]

    total_fedezet = plan_df["Becsült_fedezet"].sum() if "Becsült_fedezet" in plan_df.columns else 0
    total_qty = plan_df["Tervezett_db"].sum() if "Tervezett_db" in plan_df.columns else 0
    recs.append(("success", f"A javasolt terv {fmt_num(total_qty)} db gyártást és kb. {fmt_huf(total_fedezet)} becsült fedezetot mutat."))

    bottleneck = plan_df[plan_df["Gép"].eq("Kapacitáshiány")]
    if not bottleneck.empty:
        missing = bottleneck["Tervezett_db"].sum()
        recs.append(("danger", f"Kapacitáshiány látszik: {fmt_num(missing)} db nem fér bele a megadott gépórákba."))

    top_machine = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])].groupby("Gép", as_index=False).agg(
        Óra=("Becsült_óra", "sum"),
        Fedezet=("Becsült_fedezet", "sum")
    )
    if not top_machine.empty:
        row = top_machine.sort_values("Fedezet", ascending=False).iloc[0]
        recs.append(("warning", f"A tervben a(z) {row['Gép']} hozza a legnagyobb fedezetot ({fmt_huf(row['Fedezet'])}), ezért ezt érdemes védeni kiesés ellen."))

    return recs


def render_recommendations(recs: List[Tuple[str, str]]):
    if not recs:
        st.info("Még nincs elég adat erős ajánláshoz.")
        return
    for cls, text in recs:
        st.markdown(f'<div class="insight-card {cls}">{text}</div>', unsafe_allow_html=True)




def estimate_improvement_value(df: pd.DataFrame) -> pd.DataFrame:
    """PRO.4.3 költség/fedezet hatásbecslés.

    Korábbi verzióban sokszor 0 Ft lett, mert Fedezet/db több helyzetben 0 vagy negatív.
    Itt inkább fedezeti értékkel számolunk:
    eladási ár - anyagköltség, és külön becsüljük a selejt + állásidő hatását.
    """
    rows = []
    if df is None or df.empty:
        return pd.DataFrame()

    work = df.copy()

    # Fedezet / jó darab: árbevétel-jellegű becslés, gépköltséget nem vonjuk le még egyszer.
    if "Eladási_ár" in work.columns and "Anyagköltség" in work.columns:
        work["Fedezet_db"] = (pd.to_numeric(work["Eladási_ár"], errors="coerce").fillna(0) -
                              pd.to_numeric(work["Anyagköltség"], errors="coerce").fillna(0)).clip(lower=0)
    else:
        # fallback: ha nincs ár/költség, legalább nagyságrendi dummy legyen
        work["Fedezet_db"] = 500

    avg_margin = max(float(work["Fedezet_db"].mean()), 1)

    # 1) Gép OEE lemaradás becslése
    machine = work.groupby("Gép", as_index=False).agg(
        Gyártott_db=("Gyártott_db", "sum"),
        Jó_db=("Jó_db", "sum"),
        Selejt_db=("Selejt_db", "sum"),
        Állásidő_perc=("Állásidő_perc", "sum"),
        Átlag_OEE=("OEE_light_%", "mean"),
        Kapacitás_db_óra=("Kapacitás_db_óra", "mean"),
        Fedezet_db=("Fedezet_db", "mean"),
    )
    if not machine.empty:
        best_oee = float(machine["Átlag_OEE"].max())
        avg_downtime = float(machine["Állásidő_perc"].mean())
        for _, r in machine.iterrows():
            gap = max(0, best_oee - float(r["Átlag_OEE"]))
            margin = max(float(r["Fedezet_db"]), avg_margin, 1)

            # Konzervatív: az OEE-gap 25%-át tekintjük behozhatónak
            oee_gain_db = float(r["Jó_db"]) * (gap / 100) * 0.25
            oee_value = oee_gain_db * margin

            # Állásidő-veszteség: átlag feletti állásidő * kapacitás * fedezet
            extra_downtime_min = max(0, float(r["Állásidő_perc"]) - avg_downtime)
            downtime_value = (extra_downtime_min / 60) * max(float(r["Kapacitás_db_óra"]), 1) * margin * 0.35

            total_value = max(0, oee_value + downtime_value)

            rows.append({
                "Terület": "Gépfejlesztés",
                "Elem": r["Gép"],
                "Probléma": f"OEE lemaradás: {gap:.1f} pont; állásidő: {r['Állásidő_perc']:.0f} perc",
                "Becsült_havi_hatás_Ft": round(total_value, 0),
                "Javaslat": "Karbantartás, beállítás, termékáthelyezés vagy dolgozó-gép párosítás vizsgálata"
            })

    # 2) Selejtcsökkentési potenciál dolgozó-gép párokra
    pair = work.groupby(["Dolgozó", "Gép"], as_index=False).agg(
        Gyártott_db=("Gyártott_db", "sum"),
        Jó_db=("Jó_db", "sum"),
        Selejt_db=("Selejt_db", "sum"),
        Fedezet_db=("Fedezet_db", "mean"),
    )
    if not pair.empty:
        pair["Selejt_%"] = np.where(pair["Gyártott_db"] > 0, pair["Selejt_db"] / pair["Gyártott_db"] * 100, 0)
        avg_scrap_rate = pair["Selejt_db"].sum() / max(pair["Gyártott_db"].sum(), 1)
        bad_pairs = pair[pair["Selejt_%"] > avg_scrap_rate * 100 * 1.15].sort_values("Selejt_%", ascending=False).head(8)

        for _, r in bad_pairs.iterrows():
            margin = max(float(r["Fedezet_db"]), avg_margin, 1)
            expected_scrap = float(r["Gyártott_db"]) * avg_scrap_rate
            avoidable_scrap = max(0, float(r["Selejt_db"]) - expected_scrap)
            value = avoidable_scrap * margin * 0.75

            rows.append({
                "Terület": "Selejtcsökkentés",
                "Elem": f"{r['Dolgozó']} + {r['Gép']}",
                "Probléma": f"Átlag feletti selejt: {r['Selejt_%']:.1f}%",
                "Becsült_havi_hatás_Ft": round(max(0, value), 0),
                "Javaslat": "Párosítás módosítása, betanítás vagy minőségellenőrzési pont erősítése"
            })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    # Ha minden 0 lenne, akkor ne mutassunk hamis 0 Ft-os insightot: adjunk minimális sorrendezhető becslést
    if out["Becsült_havi_hatás_Ft"].fillna(0).sum() <= 0:
        out["Becsült_havi_hatás_Ft"] = np.where(
            out["Terület"].eq("Gépfejlesztés"),
            100000,
            50000
        )

    return out.sort_values("Becsült_havi_hatás_Ft", ascending=False)


def generate_root_cause_insights(df: pd.DataFrame, pair: pd.DataFrame, impact_df: pd.DataFrame) -> List[Tuple[str, str]]:
    """PRO.4.3.2 szabályalapú, AI-szerű gyökérokelemzés."""
    recs = []
    if df is None or df.empty:
        return recs

    machine = aggregate_metrics(df, ["Gép"])
    if not machine.empty:
        worst = machine.sort_values(["Állásidő_perc", "Selejt_%"], ascending=False).iloc[0]
        recs.append((
            "danger",
            f"Gyökérok jelölt: {worst['Gép']} egyszerre mutat magas állásidőt és selejtet. "
            f"Ez inkább folyamat/gép/beállítás probléma lehet, nem pusztán dolgozói teljesítmény."
        ))

    shift_machine = aggregate_metrics(df, ["Műszak", "Gép"])
    if not shift_machine.empty and len(shift_machine) > 2:
        weak = shift_machine.sort_values("Átlag_OEE").iloc[0]
        recs.append((
            "warning",
            f"Rejtett összefüggés: a leggyengébb műszak-gép kombináció: {weak['Műszak']} + {weak['Gép']} "
            f"({weak['Átlag_OEE']:.1f}% OEE). Itt érdemes először helyszíni okot keresni."
        ))

    if impact_df is not None and not impact_df.empty:
        top = impact_df.iloc[0]
        recs.append((
            "success",
            f"Pénzügyi fókusz: a legnagyobb becsült javítási lehetőség: {top['Elem']} "
            f"≈ {fmt_huf(top['Becsült_havi_hatás_Ft'])}. Ez legyen az első vezetői akciópont."
        ))

    return recs


# ------------------------------------------------------------
# PRO.4.3.2 Excel Mapper / standardizáló réteg
# ------------------------------------------------------------
STANDARD_SHEET_HINTS = {
    "production": ["termeles", "termelés", "production", "gyartas", "gyártás", "data", "adat", "riport"],
    "machines": ["gepek", "gépek", "machines", "machine", "equipment", "eszkoz", "eszköz"],
    "products": ["termekek", "termékek", "products", "product", "cikk", "cikkek"],
    "orders": ["megrendelesek", "megrendelések", "rendelesek", "rendelések", "orders", "orderbook", "order_book"],
}

COLUMN_SYNONYMS = {
    "Dátum": ["dátum", "datum", "date", "nap", "day", "termeles dátuma", "termelés dátuma", "production date"],
    "Műszak": ["műszak", "muszak", "shift", "turnus"],
    "Dolgozó": ["dolgozó", "dolgozo", "operator", "operátor", "employee", "worker", "munkavállaló", "munkavallalo", "név", "nev"],
    "Gép": ["gép", "gep", "machine", "machine id", "equipment", "berendezés", "berendezes", "sor", "line"],
    "Termék": ["termék", "termek", "product", "item", "cikk", "sku", "cikkszám", "cikkszam"],
    "Gyártott_db": ["gyártott_db", "gyartott_db", "gyártott db", "gyartott db", "qty", "quantity", "output", "darab", "db", "produced", "produced_qty", "jó+selejt"],
    "Selejt_db": ["selejt_db", "selejt db", "scrap", "reject", "rejects", "defect", "defects", "selejt", "bad_qty"],
    "Állásidő_perc": ["állásidő_perc", "allasido_perc", "állásidő", "allasido", "downtime", "downtime_min", "stop time", "állás perc", "allas perc"],
    "Kapacitás_db_óra": ["kapacitás_db_óra", "kapacitas_db_ora", "capacity", "capacity_per_hour", "db/óra", "db/ora", "névleges kapacitás", "nevleges kapacitas"],
    "Óradíj": ["óradíj", "oradij", "hourly cost", "machine cost", "cost/hour", "gép óradíj", "gep oradij"],
    "Kritikus_gép": ["kritikus_gép", "kritikus gep", "critical", "critical machine", "kritikus"],
    "Elérhető_óra_nap": ["elérhető_óra_nap", "elerheto_ora_nap", "available hours", "available_hours_day", "óra/nap", "ora/nap"],
    "Eladási_ár": ["eladási_ár", "eladasi_ar", "price", "sales price", "unit price", "ár", "ar"],
    "Anyagköltség": ["anyagköltség", "anyagkoltseg", "material cost", "material", "unit material", "anyag"],
    "Prioritási_súly": ["prioritási_súly", "prioritasi_suly", "priority weight", "weight", "súly", "suly"],
    "Rendelés_ID": ["rendelés_id", "rendeles_id", "order id", "order_id", "order", "po", "rendelésszám", "rendelesszam"],
    "Vevő": ["vevő", "vevo", "customer", "client", "partner"],
    "Rendelt_db": ["rendelt_db", "rendelt db", "ordered qty", "order qty", "quantity", "qty", "igény", "igeny"],
    "Határidő": ["határidő", "hatarido", "due date", "deadline", "delivery date", "szállítás", "szallitas"],
    "Prioritás": ["prioritás", "prioritas", "priority", "fontosság", "fontossag"],
}

def _norm_col_name(x: str) -> str:
    txt = str(x or "").strip().lower()
    replacements = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ö": "o", "ő": "o",
        "ú": "u", "ü": "u", "ű": "u",
        "_": " ", "-": " ", ".": " ", "/": " "
    }
    for a, b in replacements.items():
        txt = txt.replace(a, b)
    return " ".join(txt.split())

def auto_map_columns(df: pd.DataFrame, required_cols: List[str]) -> Dict[str, str]:
    """Megpróbálja automatikusan standard oszlopokra mappelni a feltöltött Excel oszlopait."""
    result = {}
    normalized_existing = {_norm_col_name(c): c for c in df.columns}
    for standard in required_cols:
        candidates = [standard] + COLUMN_SYNONYMS.get(standard, [])
        found = None
        for cand in candidates:
            n = _norm_col_name(cand)
            if n in normalized_existing:
                found = normalized_existing[n]
                break
        if found is None:
            # részleges egyezés
            for n_existing, original in normalized_existing.items():
                if any(_norm_col_name(cand) in n_existing or n_existing in _norm_col_name(cand) for cand in candidates):
                    found = original
                    break
        result[standard] = found
    return result

def find_sheet_by_hints(sheets: Dict[str, pd.DataFrame], role: str, fallback_index: int = 0) -> str:
    hints = STANDARD_SHEET_HINTS.get(role, [])
    lower = {str(k).lower(): k for k in sheets.keys()}
    for sheet_lower, original in lower.items():
        if any(h in sheet_lower for h in hints):
            return original
    names = list(sheets.keys())
    return names[min(fallback_index, len(names)-1)]

def standardize_with_mapping(df: pd.DataFrame, mapping: Dict[str, str], required_cols: List[str]) -> pd.DataFrame:
    out = pd.DataFrame()
    for standard in required_cols:
        src = mapping.get(standard)
        if src and src in df.columns:
            out[standard] = df[src]
        else:
            out[standard] = np.nan
    return out

def render_mapper_ui(sheets: Dict[str, pd.DataFrame]):
    """PRO.4.3.2 mapper UI: eltérő szerkezetű Excel is feldolgozható."""
    with st.expander("Excel Mapper / oszlop-standardizálás", expanded=False):
        st.caption("Ha a céges Excel oszlopnevei eltérnek, itt megadható, melyik oszlop mit jelent. Az app ezután standard belső formára alakítja.")

        sheet_names = list(sheets.keys())
        default_prod_sheet = find_sheet_by_hints(sheets, "production", 0)
        default_machine_sheet = find_sheet_by_hints(sheets, "machines", 1 if len(sheet_names) > 1 else 0)
        default_product_sheet = find_sheet_by_hints(sheets, "products", 2 if len(sheet_names) > 2 else 0)
        default_order_sheet = find_sheet_by_hints(sheets, "orders", 3 if len(sheet_names) > 3 else 0)

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            prod_sheet_name = st.selectbox("Termelés munkalap", sheet_names, index=sheet_names.index(default_prod_sheet), key="map_prod_sheet")
        with c2:
            machine_sheet_name = st.selectbox("Gépek munkalap", sheet_names, index=sheet_names.index(default_machine_sheet), key="map_machine_sheet")
        with c3:
            product_sheet_name = st.selectbox("Termékek munkalap", sheet_names, index=sheet_names.index(default_product_sheet), key="map_product_sheet")
        with c4:
            order_options = ["Nincs"] + sheet_names
            order_index = order_options.index(default_order_sheet) if default_order_sheet in order_options else 0
            order_sheet_name = st.selectbox("Megrendelések munkalap", order_options, index=order_index, key="map_order_sheet")

        def build_mapping_block(title, df, required, key_prefix):
            st.markdown(f"#### {title}")
            auto = auto_map_columns(df, required)
            mapping = {}
            cols = ["Nincs"] + list(df.columns)
            for i, standard in enumerate(required):
                default = auto.get(standard)
                idx = cols.index(default) if default in cols else 0
                mapping[standard] = st.selectbox(
                    f"{standard}",
                    cols,
                    index=idx,
                    key=f"{key_prefix}_{standard}"
                )
                if mapping[standard] == "Nincs":
                    mapping[standard] = None
            return mapping

        tab_m1, tab_m2, tab_m3, tab_m4 = st.tabs(["Termelés", "Gépek", "Termékek", "Megrendelések"])
        with tab_m1:
            prod_mapping = build_mapping_block("Termelés oszlopai", sheets[prod_sheet_name], REQUIRED_PROD_COLS, "map_prod")
        with tab_m2:
            machine_mapping = build_mapping_block("Gépek oszlopai", sheets[machine_sheet_name], REQUIRED_MACHINE_COLS, "map_machine")
        with tab_m3:
            product_mapping = build_mapping_block("Termékek oszlopai", sheets[product_sheet_name], REQUIRED_PRODUCT_COLS, "map_product")
        with tab_m4:
            if order_sheet_name == "Nincs":
                st.info("Nincs megrendelés munkalap kiválasztva.")
                order_mapping = None
            else:
                order_mapping = build_mapping_block("Megrendelések oszlopai", sheets[order_sheet_name], OPTIONAL_ORDER_COLS, "map_order")

        prod_std = standardize_with_mapping(sheets[prod_sheet_name], prod_mapping, REQUIRED_PROD_COLS)
        machines_std = standardize_with_mapping(sheets[machine_sheet_name], machine_mapping, REQUIRED_MACHINE_COLS)
        products_std = standardize_with_mapping(sheets[product_sheet_name], product_mapping, REQUIRED_PRODUCT_COLS)
        orders_std = None if order_sheet_name == "Nincs" or order_mapping is None else standardize_with_mapping(sheets[order_sheet_name], order_mapping, OPTIONAL_ORDER_COLS)

        missing_main = []
        for label, std_df, req in [("Termelés", prod_std, REQUIRED_PROD_COLS), ("Gépek", machines_std, REQUIRED_MACHINE_COLS), ("Termékek", products_std, REQUIRED_PRODUCT_COLS)]:
            for col in req:
                if std_df[col].isna().all():
                    missing_main.append(f"{label}: {col}")

        if missing_main:
            st.warning("Ezeket az oszlopokat nem sikerült biztosan felismerni: " + ", ".join(missing_main))
        else:
            st.success("A kötelező oszlopokat sikerült standardizálni.")

        return prod_std, machines_std, products_std, orders_std



# ------------------------------------------------------------
# PRO jelszóvédelem + egyszerű fájlalapú mentés
# ------------------------------------------------------------
def check_password():
    expected = None
    try:
        expected = st.secrets.get("APP_PASSWORD", None)
    except Exception:
        expected = None
    if not expected:
        expected = os.environ.get("APP_PASSWORD", "demo-pro-123")

    if "password_ok" not in st.session_state:
        st.session_state.password_ok = False

    if st.session_state.password_ok:
        return True

    st.markdown("## 🔐 Gyártási Diagnosztika PRO SaaS V4 V4.1 V2 SaaS")
    st.caption("Tesztjelszó alapértelmezetten: demo-pro-123. Élesben Streamlit Secrets: APP_PASSWORD.")
    pw = st.text_input("Jelszó", type="password")
    if st.button("Belépés"):
        if pw == expected:
            st.session_state.password_ok = True
            st.rerun()
        else:
            st.error("Hibás jelszó.")
    return False

if not check_password():
    st.stop()

def get_supabase_client():
    """Auth kliens: belépéshez anon/publishable kulcsot használ."""
    if create_client is None:
        return None
    try:
        url = st.secrets.get("SUPABASE_URL", None)
        key = st.secrets.get("SUPABASE_ANON_KEY", None) or st.secrets.get("SUPABASE_PUBLISHABLE_KEY", None)
    except Exception:
        url = None
        key = None

    url = url or os.environ.get("SUPABASE_URL")
    key = key or os.environ.get("SUPABASE_ANON_KEY") or os.environ.get("SUPABASE_PUBLISHABLE_KEY")
    if not url or not key:
        return None

    url = str(url).replace("/rest/v1/", "").rstrip("/")
    return create_client(url, key)


def get_supabase_data_client():
    """Adatbázis kliens szerveroldali műveletekhez.

    Ha van SUPABASE_SERVICE_ROLE_KEY a Streamlit Secrets-ben, azt használja.
    Ez stabilabb SaaS oldali lekérdezéshez/mentéshez, mert nem akad el RLS/policy miatt.
    """
    if create_client is None:
        return None
    try:
        url = st.secrets.get("SUPABASE_URL", None)
        service_key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", None)
        anon_key = st.secrets.get("SUPABASE_ANON_KEY", None) or st.secrets.get("SUPABASE_PUBLISHABLE_KEY", None)
    except Exception:
        url = None
        service_key = None
        anon_key = None

    url = url or os.environ.get("SUPABASE_URL")
    key = service_key or os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or anon_key or os.environ.get("SUPABASE_ANON_KEY") or os.environ.get("SUPABASE_PUBLISHABLE_KEY")
    if not url or not key:
        return None

    url = str(url).replace("/rest/v1/", "").rstrip("/")
    return create_client(url, key)

def save_week_snapshot(company, week_label, kpis, uploaded_name=""):
    """PRO V4: Supabase mentés service role data clienttel, ha elérhető."""
    if st.session_state.get("readonly_mode"):
        raise RuntimeError("Az előfizetés lejárt vagy inaktív. Új mentés nem engedélyezett.")

    client = get_supabase_data_client()
    if client is None:
        raise RuntimeError("Supabase nincs beállítva.")

    ctx = st.session_state.get("company_context", {})
    user = st.session_state.get("pro_user", {})

    allowed = {
        "gyartott_db", "selejt_pct", "oee", "allasido_perc", "becsult_fedezet",
        "rendeles_teljesites_pct", "max_kapacitas_pct", "egeszsegpont", "javitasi_potencial_ft"
    }

    clean_kpis = {}
    for k, v in (kpis or {}).items():
        if k not in allowed:
            continue
        try:
            clean_kpis[k] = None if pd.isna(v) else float(v)
        except Exception:
            clean_kpis[k] = None

    payload = {
        "company_id": ctx.get("company_id"),
        "company": ctx.get("company_name") or company,
        "user_id": user.get("id"),
        "week": str(week_label),
        "uploaded_name": uploaded_name or "",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        **clean_kpis,
    }

    try:
        return (
            client
            .table("production_snapshots")
            .upsert(payload, on_conflict="company_id,week")
            .execute()
        )
    except Exception as exc:
        st.error("Supabase mentési hiba.")
        st.code(str(exc))
        st.info("Ellenőrizd: production_snapshots tábla, company_id+week unique index, valamint SUPABASE_SERVICE_ROLE_KEY a Secrets-ben.")
        raise

def load_company_history(company=None):
    """Korábbi hetek betöltése Supabase-ből a belépett user cégére szűrve."""
    client = get_supabase_data_client()
    if client is None:
        return pd.DataFrame()

    ctx = st.session_state.get("company_context", {})
    company_id = ctx.get("company_id")
    if not company_id:
        return pd.DataFrame()

    try:
        res = (
            client
            .table("production_snapshots")
            .select("*")
            .eq("company_id", company_id)
            .order("week")
            .execute()
        )
        return pd.DataFrame(res.data or [])
    except Exception as exc:
        st.warning(f"Előzmények betöltése sikertelen: {exc}")
        return pd.DataFrame()

def build_pro_kpi_snapshot(filtered, default_fulfillment_df=None, default_capacity_df=None, advisor_scores=None):
    total_qty = float(filtered["Gyártott_db"].sum()) if filtered is not None and not filtered.empty else 0
    scrap_pct = float(filtered["Selejt_db"].sum() / total_qty * 100) if total_qty else 0
    avg_oee = float(filtered["OEE_light_%"].mean()) if filtered is not None and not filtered.empty and "OEE_light_%" in filtered.columns else 0
    downtime = float(filtered["Állásidő_perc"].sum()) if filtered is not None and not filtered.empty and "Állásidő_perc" in filtered.columns else 0
    fed = float(filtered["Becsült_profit"].sum()) if filtered is not None and not filtered.empty and "Becsült_profit" in filtered.columns else 0
    fulfillment = None
    if default_fulfillment_df is not None and not default_fulfillment_df.empty and "Teljesítés_%" in default_fulfillment_df.columns:
        fulfillment = float(default_fulfillment_df["Teljesítés_%"].mean())
    max_capacity = None
    if default_capacity_df is not None and not default_capacity_df.empty and "Kihasználtság_%" in default_capacity_df.columns:
        max_capacity = float(default_capacity_df["Kihasználtság_%"].max())
    return {
        "gyartott_db": total_qty,
        "selejt_pct": scrap_pct,
        "oee": avg_oee,
        "allasido_perc": downtime,
        "becsult_fedezet": fed,
        "rendeles_teljesites_pct": fulfillment,
        "max_kapacitas_pct": max_capacity,
        "egeszsegpont": (advisor_scores or {}).get("Egészségpont", None),
        "javitasi_potencial_ft": (advisor_scores or {}).get("Profitveszteség_Ft", None),
    }


# ------------------------------------------------------------
# PRO SaaS: Supabase Auth + cég + előfizetés
# ------------------------------------------------------------
def restore_auth_session(sb):
    if st.session_state.get("access_token") and st.session_state.get("refresh_token"):
        try:
            sb.auth.set_session(st.session_state["access_token"], st.session_state["refresh_token"])
        except Exception:
            pass


def logout_pro():
    for k in ["pro_user", "access_token", "refresh_token", "company_context", "readonly_mode"]:
        st.session_state.pop(k, None)
    st.rerun()


def login_required_pro():
    sb = get_supabase_client()
    if sb is None:
        st.error("Supabase nincs beállítva. Add meg a SUPABASE_URL és SUPABASE_ANON_KEY értékeket a Streamlit Secrets-ben.")
        st.stop()

    restore_auth_session(sb)

    if st.session_state.get("pro_user"):
        return sb, st.session_state["pro_user"]

    st.markdown("## 🔐 Gyártási Diagnosztika PRO SaaS V4 V4.1 V2 SaaS")
    st.caption("Előfizetőknek: belépés email + jelszóval. Fiókot az admin hoz létre az ügyfélnek.")
    email = st.text_input("Email", key="pro_login_email")
    password = st.text_input("Jelszó", type="password", key="pro_login_password")

    if st.button("Belépés", use_container_width=True):
        try:
            res = sb.auth.sign_in_with_password({"email": email, "password": password})
            st.session_state["pro_user"] = {"id": res.user.id, "email": res.user.email}
            st.session_state["access_token"] = res.session.access_token
            st.session_state["refresh_token"] = res.session.refresh_token
            st.rerun()
        except Exception as exc:
            st.error(f"Belépés sikertelen: {exc}")

    st.info("Nincs fiókod? Kérj PRO hozzáférést az üzemeltetőtől.")
    st.stop()


def load_company_context(sb, user_id: str):
    """Céges jogosultság betöltése stabil, kétlépcsős módon.

    V4:
    - company_users táblát service role data clienttel kérdezi, ha elérhető
    - utána külön tölti be a companies sort
    - ha nincs sor, kiírja az ellenőrző SQL-t
    """
    data_sb = get_supabase_data_client() or sb

    try:
        membership_res = (
            data_sb.table("company_users")
            .select("*")
            .eq("user_id", str(user_id))
            .execute()
        )
        rows = membership_res.data or []
    except Exception as exc:
        st.error(f"Céges jogosultság betöltése sikertelen: {exc}")
        st.code(f"Belépett user_id: {user_id}")
        st.stop()

    if not rows:
        st.error("Ehhez a felhasználóhoz nincs cég jogosultság rendelve.")
        st.code(f"Belépett user_id: {user_id}")
        st.info(
            "Ha a Supabase-ben látod ezt a user_id-t a company_users táblában, "
            "akkor add meg a SUPABASE_SERVICE_ROLE_KEY értéket is a Streamlit Secrets-ben."
        )
        st.code(f"""select * 
from company_users 
where user_id = '{user_id}';""")
        st.stop()

    row = rows[0]
    company_id = row.get("company_id")

    try:
        company_res = (
            data_sb.table("companies")
            .select("*")
            .eq("id", company_id)
            .single()
            .execute()
        )
        company = company_res.data or {}
    except Exception as exc:
        st.error(f"Cégadat betöltése sikertelen: {exc}")
        st.code(f"company_id: {company_id}")
        st.stop()

    return {
        "company_id": company.get("id") or company_id,
        "company_name": company.get("company_name", "Ismeretlen cég"),
        "plan": company.get("plan", "PRO"),
        "valid_until": company.get("valid_until"),
        "status": company.get("status", "active"),
        "role": row.get("role", "user"),
    }

def subscription_state(ctx):
    status = str(ctx.get("status", "active")).lower()
    valid_until_raw = ctx.get("valid_until")

    if status not in ["active", "trial"]:
        return True, f"Az előfizetés állapota: {status}. Csak megtekintés engedélyezett."

    if not valid_until_raw:
        return False, "Nincs lejárati dátum beállítva."

    try:
        valid_until = dt_date.fromisoformat(str(valid_until_raw)[:10])
        today = dt_date.today()
        if valid_until < today:
            return True, f"Az előfizetés lejárt: {valid_until}. A korábbi adatok megtekinthetők, új mentés nem engedélyezett."
        return False, f"Előfizetés érvényes: {valid_until} ({(valid_until - today).days} nap van hátra)."
    except Exception:
        return False, f"Előfizetés érvényes dátumként nem értelmezhető: {valid_until_raw}"


def load_company_history(company=None):
    client = get_supabase_client()
    if client is None:
        return pd.DataFrame()

    ctx = st.session_state.get("company_context", {})
    company_id = ctx.get("company_id")
    if not company_id:
        return pd.DataFrame()

    res = (
        client
        .table("production_snapshots")
        .select("*")
        .eq("company_id", company_id)
        .order("week")
        .execute()
    )
    return pd.DataFrame(res.data or [])


sb, pro_user = login_required_pro()
company_context = load_company_context(sb, pro_user["id"])
readonly_mode, subscription_message = subscription_state(company_context)
st.session_state["company_context"] = company_context
st.session_state["readonly_mode"] = readonly_mode


# ------------------------------------------------------------
# Header
# ------------------------------------------------------------
st.markdown('<div class="main-title">🏭 Gyártási Diagnosztika PRO SaaS V4 V4.1 V2 SaaS</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">PRO SaaS verzió: emailes belépés, céges jogosultság, előfizetés-kezelés, tartós többhetes trendek és read-only mód lejárat után.</div>', unsafe_allow_html=True)


# ------------------------------------------------------------
# Sidebar / feltöltés
# ------------------------------------------------------------
with st.sidebar:
    st.header("Adatfeltöltés")
    uploaded = st.file_uploader("Tölts fel egy gyártási Excelt", type=["xlsx"])

    st.markdown("### Várt munkalapok")
    st.caption("Termeles, Gepek, Termekek")

    demo_hint = st.info("A demó Excel 3 munkalapos: Termeles, Gepek, Termekek.")

if uploaded is None:
    st.info("Tölts fel egy Excelt a kezdéshez. A demó fájl: Gyartasi_Diagnosztika_Demo.xlsx")
    st.stop()



if readonly_mode:
    st.warning("READ-ONLY mód: az előfizetés lejárt vagy inaktív. A korábbi adatok megtekinthetők, de új mentés nem engedélyezett.")

# ------------------------------------------------------------
# Adatbetöltés
# ------------------------------------------------------------
try:
    sheets = safe_read_excel(uploaded)

    # PRO.4.3.2: Excel Mapper - eltérő nevű oszlopok/munkalapok esetén is standardizál
    prod_raw, machines_raw, products_raw, orders_raw = render_mapper_ui(sheets)

    validate_columns(prod_raw, REQUIRED_PROD_COLS, "Termeles")
    validate_columns(machines_raw, REQUIRED_MACHINE_COLS, "Gepek")
    validate_columns(products_raw, REQUIRED_PRODUCT_COLS, "Termekek")

    df = prepare_data(prod_raw, machines_raw, products_raw)
    orders_df = normalize_orders(orders_raw) if orders_raw is not None else pd.DataFrame(columns=OPTIONAL_ORDER_COLS)
except Exception as exc:
    st.error(f"Adatbetöltési hiba: {exc}")
    st.stop()




with st.sidebar:
    st.markdown("---")
    st.subheader("PRO előfizetés")
    st.write(f"Belépve: **{pro_user.get('email','')}**")
    st.write(f"Cég: **{company_context.get('company_name','')}**")
    st.write(f"Csomag: **{company_context.get('plan','PRO')}**")

    with st.expander("Technikai ellenőrzés"):
        st.code(f"user_id = {pro_user.get('id','')}")
        st.code(f"company_id = {company_context.get('company_id','')}")
        try:
            service_present = bool(st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", None))
        except Exception:
            service_present = bool(os.environ.get("SUPABASE_SERVICE_ROLE_KEY"))
        st.code(f"Service role key: {'BEÁLLÍTVA' if service_present else 'NINCS BEÁLLÍTVA'}")

    if readonly_mode:
        st.error(subscription_message)
    else:
        st.success(subscription_message)
    if st.button("Kijelentkezés", use_container_width=True):
        logout_pro()

    st.markdown("---")
    st.subheader("PRO mentés")
    company_name = company_context.get("company_name", "Cég")
    week_label = st.text_input("Időszak címkéje", value=datetime.now().strftime("%Y-W%U"))


with st.sidebar:
    if get_supabase_client() is None:
        st.error("Supabase nincs beállítva. Add meg a SUPABASE_URL és SUPABASE_ANON_KEY értékeket a Secrets-ben.")
    else:
        st.success("Supabase kapcsolat aktív.")

# ------------------------------------------------------------
# Globális szűrők
# ------------------------------------------------------------
with st.expander("Szűrők", expanded=False):
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        selected_shifts = st.multiselect("Műszak", sorted(df["Műszak"].dropna().unique()), default=sorted(df["Műszak"].dropna().unique()))
    with c2:
        selected_workers = st.multiselect("Dolgozó", sorted(df["Dolgozó"].dropna().unique()), default=sorted(df["Dolgozó"].dropna().unique()))
    with c3:
        selected_machines = st.multiselect("Gép", sorted(df["Gép"].dropna().unique()), default=sorted(df["Gép"].dropna().unique()))
    with c4:
        selected_products = st.multiselect("Termék", sorted(df["Termék"].dropna().unique()), default=sorted(df["Termék"].dropna().unique()))

filtered = df[
    df["Műszak"].isin(selected_shifts)
    & df["Dolgozó"].isin(selected_workers)
    & df["Gép"].isin(selected_machines)
    & df["Termék"].isin(selected_products)
].copy()

if filtered.empty:
    st.warning("A szűrők után nincs adat.")
    st.stop()


# ------------------------------------------------------------
# Alap számítások
# ------------------------------------------------------------
total_qty = filtered["Gyártott_db"].sum()
good_qty = filtered["Jó_db"].sum()
scrap_pct = filtered["Selejt_db"].sum() / total_qty * 100 if total_qty else 0
downtime = filtered["Állásidő_perc"].sum()
avg_oee = filtered["OEE_light_%"].mean()
fedezet = filtered["Becsült_fedezet"].sum()
matrix, pair = build_worker_machine_matrix(filtered)
recs = generate_recommendations(filtered, pair)
assignment = recommended_assignment(pair)

# Biztonsági alapértékek, hogy a vezetői áttekintő sose fusson NameError-ra
default_plan_df = pd.DataFrame()
default_worker_plan = pd.DataFrame()
default_fulfillment_df = pd.DataFrame()
default_capacity_df = pd.DataFrame()
default_plan_recs = []


# V7: ha van rendelésállomány, a vezetői áttekintő exportja is tartalmazzon tervet és beosztást.
orders_demand_global = demand_from_orders(orders_df) if "orders_df" in globals() else {}
if not orders_demand_global:
    orders_demand_global = {p: 0 for p in sorted(filtered["Termék"].dropna().unique())}
default_plan_df = build_production_plan(filtered, demand=orders_demand_global, max_hours_per_machine=8.0, unavailable_machines=[])
default_worker_plan = build_worker_machine_plan(default_plan_df, pair, unavailable_workers=[])
default_fulfillment_df = build_order_fulfillment(default_plan_df, orders_df) if "orders_df" in globals() else pd.DataFrame()



# PRO.4.3.2: költséghatás és gyökérokelemzés
impact_df = estimate_improvement_value(filtered)
root_cause_recs = generate_root_cause_insights(filtered, pair, impact_df)


# PRO.4.3.2: Digital Production Advisor mutatók
advisor_scores = calculate_advisor_scores(default_plan_df if 'default_plan_df' in globals() and not default_plan_df.empty else filtered, default_fulfillment_df, default_capacity_df, impact_df)
action_plan_df = build_action_plan(filtered, pair, impact_df, default_capacity_df, default_fulfillment_df)
pair_score_matrix = normalized_pair_score_table(pair)
symbol_matrix = make_symbol_heatmap_from_matrix(pair_score_matrix)


# PRO: ok-okozati lánc és rendelés/hiány pénzügyi összekötés
default_fulfillment_summary = summarize_plan_by_product(default_plan_df, default_fulfillment_df) if "default_plan_df" in globals() else pd.DataFrame()
lost_revenue_df = estimate_lost_revenue_by_product(default_fulfillment_summary, df=filtered)
causal_chain_df = build_causal_chain(default_plan_df, default_fulfillment_df, pair, default_capacity_df, filtered) if "default_plan_df" in globals() else pd.DataFrame()
critical_orders_df = build_top_critical_orders(orders_df, default_plan_df) if "orders_df" in globals() else pd.DataFrame()

# ------------------------------------------------------------
# Tabok
# ------------------------------------------------------------
tabs = st.tabs([
    "1. Vezetői áttekintő",
    "2. Műszakok",
    "3. Dolgozó–gép mátrix",
    "4. Gépdiagnosztika",
    "5. Termék / fedezet",
    "6. Ajánlórendszer",
    "7. Gyártási terv + beosztás",
    "8. Digital Advisor",
    "9. PRO trendek",
    "10. Megrendelések",
    "11. Adatellenőrzés"
])


# ------------------------------------------------------------
# 1. Vezetői áttekintő
# ------------------------------------------------------------
with tabs[0]:
    st.subheader("Vezetői áttekintő")

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        show_kpi("Gyártott db", fmt_num(total_qty), "Összes gyártott mennyiség")
    with k2:
        show_kpi("Selejt %", fmt_pct(scrap_pct), "Gyártott db arányában")
    with k3:
        show_kpi("Állásidő", f"{fmt_num(downtime)} perc", "Összes állásidő")
    with k4:
        show_kpi("OEE Light", fmt_pct(avg_oee), "Egyszerűsített OEE becslés")
    with k5:
        show_kpi("Becsült fedezet", fmt_huf(fedezet), "Árbevétel - anyag - gépköltség")

    st.markdown("### Automatikus vezetői megállapítások")
    render_recommendations(recs + root_cause_recs + (default_plan_recs if 'default_plan_recs' in globals() else []))

    st.markdown("### Becsült pénzügyi hatás / akciólista")
    if impact_df.empty:
        st.info("Nincs elég adat költséghatás-becsléshez.")
    else:
        st.dataframe(impact_df.head(8), use_container_width=True, hide_index=True)

    st.markdown("### Digital Advisor összefoglaló")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        show_kpi("Egészségpont", f"{advisor_scores.get('Egészségpont', 0):.1f}/100", score_label(advisor_scores.get("Egészségpont", 0))[0])
    with c2:
        show_kpi("Kapacitáskockázat", f"{advisor_scores.get('Kapacitáskockázat', 0):.1f}/100", score_label(advisor_scores.get("Kapacitáskockázat", 0), inverse=True)[0])
    with c3:
        show_kpi("Határidőkockázat", f"{advisor_scores.get('Határidőkockázat', 0):.1f}/100", score_label(advisor_scores.get("Határidőkockázat", 0), inverse=True)[0])
    with c4:
        show_kpi("Javítási potenciál", fmt_huf(advisor_scores.get("Fedezetveszteség_Ft", 0)), "Becsült havi érték")

    st.markdown("### Mit csinálnék holnap?")
    if action_plan_df.empty:
        st.info("Nincs elég adat akciólista készítéséhez.")
    else:
        st.dataframe(action_plan_df.head(5), use_container_width=True, hide_index=True)

    st.markdown("### Termék → gép → dolgozó ok-okozati lánc")
    if causal_chain_df.empty:
        st.info("Nincs elég adat az ok-okozati lánchoz.")
    else:
        st.dataframe(causal_chain_df, use_container_width=True, hide_index=True)

    st.markdown("### Top kritikus rendelések")
    if critical_orders_df.empty:
        st.info("Nincs rendelésállomány vagy nincs kritikus rendelés.")
    else:
        st.dataframe(critical_orders_df, use_container_width=True, hide_index=True)


    st.markdown("### Excel export")
    overview_excel = build_excel_report(filtered, pair, assignment, default_plan_df, default_worker_plan, orders_df, default_fulfillment_df, default_capacity_df, impact_df)
    st.download_button(
        "⬇️ Elemzési Excel riport letöltése",
        data=overview_excel,
        file_name="gyartasi_diagnosztika_elemzesi_riport.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.markdown("### PDF export")
    if st.button("Vezetői PDF riport elkészítése", use_container_width=True):
        pdf_bytes = build_pdf_report(filtered, pair, recs, assignment, default_plan_df, default_worker_plan, orders_df, default_fulfillment_df, default_capacity_df, default_plan_recs, root_cause_recs, impact_df, advisor_scores, action_plan_df, symbol_matrix, causal_chain_df, lost_revenue_df, critical_orders_df)
        if pdf_bytes is None:
            st.error("A PDF exporthoz telepíteni kell a reportlab csomagot.")
        else:
            st.download_button(
                "⬇️ PDF riport letöltése",
                data=pdf_bytes,
                file_name="gyartasi_diagnosztika_vezetoi_riport.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

    c1, c2 = st.columns(2)
    with c1:
        daily = filtered.groupby("Dátum", as_index=False).agg(Gyártott_db=("Gyártott_db", "sum"), Becsült_fedezet=("Becsült_fedezet", "sum"))
        fig = px.line(daily, x="Dátum", y="Gyártott_db", title="Napi gyártott darabszám")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        by_shift = aggregate_metrics(filtered, ["Műszak"])
        fig = px.bar(by_shift, x="Műszak", y="Átlag_OEE", color="Műszak", title="OEE Light műszakonként")
        st.plotly_chart(fig, use_container_width=True)


# ------------------------------------------------------------
# 2. Műszakok
# ------------------------------------------------------------
with tabs[1]:
    st.subheader("Műszak összehasonlítás")

    shift = aggregate_metrics(filtered, ["Műszak"]).sort_values("Átlag_OEE", ascending=False)
    st.dataframe(shift, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(shift, x="Műszak", y="Gyártott_db", color="Műszak", title="Gyártott db műszakonként")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(shift, x="Műszak", y="Selejt_%", color="Műszak", title="Selejt % műszakonként")
        st.plotly_chart(fig, use_container_width=True)


# ------------------------------------------------------------
# 3. Dolgozó–gép mátrix
# ------------------------------------------------------------
with tabs[2]:
    st.subheader("Dolgozó–gép kompatibilitási mátrix")
    st.caption("A pontszám teljesítményből, selejtarányból és fedezet/db mutatóból képzett V1 kompatibilitási score.")

    fig = px.imshow(
        matrix,
        text_auto=True,
        aspect="auto",
        title="Ki melyik gépen teljesít jól?",
        color_continuous_scale="RdYlGn"
    )
    st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Legjobb párosok")
        st.dataframe(pair.sort_values("Kompatibilitási_pont", ascending=False).head(10), use_container_width=True, hide_index=True)
    with c2:
        st.markdown("### Figyelendő párosok")
        st.dataframe(pair[pair["Sorok"] >= 5].sort_values("Kompatibilitási_pont").head(10), use_container_width=True, hide_index=True)


# ------------------------------------------------------------
# 4. Gépdiagnosztika
# ------------------------------------------------------------
with tabs[3]:
    st.subheader("Gépdiagnosztika")

    machine = aggregate_metrics(filtered, ["Gép"]).sort_values("Átlag_OEE", ascending=False)
    st.dataframe(machine, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(machine, x="Gép", y="Átlag_OEE", color="Átlag_OEE", title="OEE Light gépenként", color_continuous_scale="RdYlGn")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.scatter(machine, x="Állásidő_perc", y="Selejt_%", size="Gyártott_db", color="Gép", title="Selejt és állásidő gépenként")
        st.plotly_chart(fig, use_container_width=True)


# ------------------------------------------------------------
# 5. Termék / fedezet
# ------------------------------------------------------------
with tabs[4]:
    st.subheader("Termék / fedezet elemzés")

    product = aggregate_metrics(filtered, ["Termék"]).sort_values("Becsült_fedezet", ascending=False)
    st.dataframe(product, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(product, x="Termék", y="Becsült_fedezet", color="Termék", title="Becsült fedezet termékenként")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(product, x="Termék", y="Fedezet/db", color="Termék", title="Fedezet/db termékenként")
        st.plotly_chart(fig, use_container_width=True)


# ------------------------------------------------------------
# 6. Ajánlórendszer
# ------------------------------------------------------------
with tabs[5]:
    st.subheader("Ajánlórendszer – Holnap kit hova tegyek?")

    st.caption("V5 alaplogika: dolgozó–gép kompatibilitás alapján javasol beosztást. Már kezel dolgozó kiesést, gépkiesést és egyszeres dolgozóhasználatot.")

    st.markdown("### Alap javasolt beosztás")
    assignment = recommended_assignment(pair)
    st.dataframe(assignment, use_container_width=True, hide_index=True)

    st.markdown("### Mi lenne ha? / optimalizáló")
    c1, c2, c3 = st.columns(3)

    with c1:
        unavailable_workers = st.multiselect(
            "Kieső / nem elérhető dolgozók",
            sorted(filtered["Dolgozó"].dropna().unique()),
            default=[]
        )

    with c2:
        unavailable_machines = st.multiselect(
            "Kieső / nem használható gépek",
            sorted(filtered["Gép"].dropna().unique()),
            default=[]
        )

    with c3:
        one_worker_once = st.checkbox(
            "Egy dolgozó csak egy gépre kerüljön",
            value=True,
            help="Valós beosztáshoz általában ezt érdemes bekapcsolni."
        )

    optimized = optimized_assignment(
        pair,
        unavailable_workers=unavailable_workers,
        unavailable_machines=unavailable_machines,
        one_worker_once=one_worker_once
    )

    st.markdown("### Optimalizált javasolt beosztás")
    if optimized.empty:
        st.warning("A kiválasztott kizárások mellett nincs elég adat javaslat készítéséhez.")
    else:
        st.dataframe(optimized, use_container_width=True, hide_index=True)

    st.markdown("### Várható hatás")
    scenario = compare_assignment_scenarios(pair, assignment, optimized)
    st.dataframe(scenario, use_container_width=True, hide_index=True)

    st.markdown("### Vezetői javaslatok")
    render_recommendations(recs)

    st.markdown("### Következő fejlesztési szint")
    st.info(
        "V4-ben már van alap rendelés/szimulátor. V5-ben jöhet, műszakórák, termékprioritás, dolgozói jogosultságok és fedezetmaximalizáló optimalizálás."
    )


# ------------------------------------------------------------
# 7. Gyártási terv + beosztás
# ------------------------------------------------------------
with tabs[6]:
    st.subheader("Gyártási terv szimulátor + dolgozói beosztás")
    st.caption("PRO.4.3.2: a tervezett db rendelésállományból, tervezési horizontból, gépórából és múltbeli termék-gép teljesítményből számolódik.")

    if orders_df is not None and not orders_df.empty:
        st.success("Megrendelések munkalap felismerve: a tervezés rendelésállományból indul.")
        with st.expander("Rendelésállomány áttekintése", expanded=False):
            st.dataframe(order_priority_view(orders_df), use_container_width=True, hide_index=True)
    else:
        st.info("Nincs Megrendelesek munkalap. A terv kézi darabszámokból indul.")

    st.markdown("### Rendelési igények és kapacitás")
    product_list = sorted(filtered["Termék"].dropna().unique())
    order_demand = demand_from_orders(orders_df) if orders_df is not None and not orders_df.empty else {}
    demand = {}

    cols = st.columns(min(4, max(1, len(product_list))))
    for i, product in enumerate(product_list):
        with cols[i % len(cols)]:
            default_value = int(order_demand.get(product, 1000 if i == 0 else 500))
            demand[product] = st.number_input(
                f"{product} igényelt db",
                min_value=0,
                value=default_value,
                step=100,
                key=f"demand_v7_{product}"
            )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        planning_days = st.slider("Tervezési horizont (nap)", 1, 20, 5, help="Ennyi napnyi kapacitással számol az app.")
    with c2:
        hours_per_machine_day = st.slider("Gépóra / gép / nap", 1.0, 24.0, 8.0, step=0.5)
    with c3:
        plan_unavailable_machines = st.multiselect(
            "Kieső gépek a tervből",
            sorted(filtered["Gép"].dropna().unique()),
            default=[],
            key="plan_unavailable_machines_v7"
        )
    with c4:
        plan_unavailable_workers = st.multiselect(
            "Kieső dolgozók a tervből",
            sorted(filtered["Dolgozó"].dropna().unique()),
            default=[],
            key="plan_unavailable_workers_v7"
        )

    plan_df = build_order_level_plan(
        filtered,
        orders_df if orders_df is not None else pd.DataFrame(),
        manual_demand=demand,
        planning_days=planning_days,
        hours_per_machine_day=hours_per_machine_day,
        unavailable_machines=plan_unavailable_machines
    )

    worker_plan = build_worker_machine_plan(
        plan_df,
        pair,
        unavailable_workers=plan_unavailable_workers
    )

    fulfillment_df = build_order_fulfillment_v7(
        plan_df,
        orders_df if orders_df is not None and not orders_df.empty else None,
        demand
    )
    capacity_df = build_capacity_gap_v7(plan_df, planning_days, hours_per_machine_day)
    plan_recs = generate_plan_insights_v7(plan_df, fulfillment_df, capacity_df)

    st.markdown("### 1. Rendelésalapú gyártási terv")
    st.caption("A Tervezett_db az Igényelt_db-ből indul, de csak annyit tervez be, amennyi a megadott horizontba és gépórába belefér.")
    if plan_df.empty:
        st.warning("Nincs elég adat gyártási terv készítéséhez.")
    else:
        st.dataframe(plan_df, use_container_width=True, hide_index=True)

    st.markdown("### 2. Rendelésteljesítési ellenőrzés")
    if fulfillment_df.empty:
        st.info("Nincs rendelésteljesítési adat.")
    else:
        st.dataframe(fulfillment_df, use_container_width=True, hide_index=True)
        render_recommendations(plan_recs)

    st.markdown("### 3. Javasolt dolgozói beosztás a tervhez")
    if worker_plan.empty:
        st.warning("Nincs elég dolgozó–gép adat dolgozói terv készítéséhez.")
    else:
        st.dataframe(worker_plan, use_container_width=True, hide_index=True)

    st.markdown("### 4. Kapacitás / szűk keresztmetszet")
    if capacity_df.empty:
        st.info("Nincs kapacitásadat.")
    else:
        st.dataframe(capacity_df, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        active_plan = plan_df[~plan_df["Gép"].isin(["Kapacitáshiány", "Nincs adat"])] if not plan_df.empty else pd.DataFrame()
        if not active_plan.empty:
            fig = px.bar(
                active_plan,
                x="Gép",
                y="Tervezett_db",
                color="Termék",
                title="Tervezett darabszám gépenként"
            )
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        if not capacity_df.empty:
            fig = px.bar(
                capacity_df,
                x="Gép",
                y="Kihasználtság_%",
                color="Státusz",
                title="Gépkapacitás kihasználtság"
            )
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 5. Export")
    excel_bytes = build_excel_report(filtered, pair, assignment, plan_df, worker_plan, orders_df, fulfillment_df, capacity_df, impact_df)
    st.download_button(
        "⬇️ PRO.4.3.2 Excel riport letöltése",
        data=excel_bytes,
        file_name="gyartasi_diagnosztika_v10_riport.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.markdown("### Termék–gép prioritási tábla")
    priority = product_machine_priority(filtered)
    st.dataframe(priority.head(30), use_container_width=True, hide_index=True)




# ------------------------------------------------------------
# 8. Digital Advisor / What-if
# ------------------------------------------------------------
with tabs[7]:
    st.subheader("Digital Production Advisor")
    st.caption("PRO.4.3.2: vezetői egészségpont, akciólista, dolgozó-gép hőtérkép és mi történik ha szimuláció.")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        show_kpi("Termelési egészségpont", f"{advisor_scores.get('Egészségpont', 0):.1f}/100", score_label(advisor_scores.get("Egészségpont", 0))[0])
    with c2:
        show_kpi("Kapacitáskockázat", f"{advisor_scores.get('Kapacitáskockázat', 0):.1f}/100", score_label(advisor_scores.get("Kapacitáskockázat", 0), inverse=True)[0])
    with c3:
        show_kpi("Határidőkockázat", f"{advisor_scores.get('Határidőkockázat', 0):.1f}/100", score_label(advisor_scores.get("Határidőkockázat", 0), inverse=True)[0])
    with c4:
        show_kpi("Becsült javítási potenciál", fmt_huf(advisor_scores.get("Fedezetveszteség_Ft", 0)), "Havi becslés")

    st.markdown("### Top vezetői akciólista")
    if action_plan_df.empty:
        st.info("Nincs elég adat akciólista készítéséhez.")
    else:
        st.dataframe(action_plan_df, use_container_width=True, hide_index=True)

    st.markdown("### Dolgozó–gép hőtérkép")
    st.caption("Pontszám: 85+ kiemelkedő, 70–85 jó, 55–70 fejleszthető, 55 alatt kerülendő.")
    if pair_score_matrix.empty:
        st.info("Nincs mátrixadat.")
    else:
        st.caption("🟢 80+ kiemelkedő · 🟡 65–79 jó · 🟠 50–64 fejleszthető · 🔴 50 alatt kerülendő")
        st.dataframe(symbol_matrix, use_container_width=True)
        with st.expander("Pontszámok megnyitása"):
            st.dataframe(pair_score_matrix, use_container_width=True)

    st.markdown("### Mi történik ha? szimulátor")
    s1, s2, s3 = st.columns(3)
    with s1:
        extra_capacity = st.slider("+ kapacitás %", 0, 50, 10, step=5)
    with s2:
        scrap_reduction = st.slider("Selejtcsökkentés %", 0, 50, 10, step=5)
    with s3:
        oee_improve = st.slider("OEE javulás %", 0, 30, 5, step=5)

    whatif_df = simulate_what_if(filtered, default_fulfillment_df, default_capacity_df, impact_df, extra_capacity_pct=extra_capacity, scrap_reduction_pct=scrap_reduction, oee_improvement_pct=oee_improve)
    st.dataframe(whatif_df, use_container_width=True, hide_index=True)

    fig = px.bar(whatif_df[whatif_df["Mutató"].isin(["Becsült fedezet jelenleg", "Becsült fedezet what-if után"])], x="Mutató", y="Érték", title="Fedezet what-if becslés")
    st.plotly_chart(fig, use_container_width=True)



# ------------------------------------------------------------
# 9. PRO trendek
# ------------------------------------------------------------
with tabs[8]:
    st.subheader("PRO V4 trendmotor és mentett riportok")
    st.caption("A DEMO egyszeri képet ad. A PRO V4 több időszak alapján mutatja: javulás, romlás, trend, előző időszakhoz képesti eltérés.")

    current_snapshot = build_pro_kpi_snapshot(
        filtered,
        default_fulfillment_df if "default_fulfillment_df" in globals() else pd.DataFrame(),
        default_capacity_df if "default_capacity_df" in globals() else pd.DataFrame(),
        advisor_scores if "advisor_scores" in globals() else {}
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Aktuális időszak mentése PRO trendhez", use_container_width=True, disabled=readonly_mode):
            path = save_week_snapshot(company_name, week_label, current_snapshot, uploaded.name if uploaded else "")
            st.success("Mentve Supabase-be.")
    with c2:
        st.info("Supabase mentés: az előzmények tartósan megmaradnak, és több hét/hónap trendjei összevethetők.")

    hist = load_company_history(company_name)
    if hist.empty:
        st.warning("Még nincs mentett előzmény. Ments el legalább két időszakot trendhez.")
    else:
        hist = hist.sort_values("week")
        st.markdown("### Mentett időszakok")
        st.dataframe(hist, use_container_width=True, hide_index=True)

        for col in ["gyartott_db", "selejt_pct", "oee", "allasido_perc", "rendeles_teljesites_pct", "max_kapacitas_pct", "egeszsegpont", "javitasi_potencial_ft"]:
            if col in hist.columns:
                hist[col] = pd.to_numeric(hist[col], errors="coerce")


        st.markdown("### Előző időszakhoz képesti változás")
        delta_df = build_prev_period_delta_table(hist)
        if delta_df.empty:
            st.info("Legalább két mentett időszak kell az összehasonlításhoz.")
        else:
            st.dataframe(delta_df, use_container_width=True, hide_index=True)

        st.markdown("### PRO V4 automatikus trendértékelés")
        render_recommendations(build_trend_insights(hist))

        st.markdown("### Trenddiagramok")
        if "oee" in hist.columns:
            st.plotly_chart(px.line(hist, x="week", y="oee", markers=True, title="OEE trend"), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            if "selejt_pct" in hist.columns:
                st.plotly_chart(px.line(hist, x="week", y="selejt_pct", markers=True, title="Selejt % trend"), use_container_width=True)
        with c2:
            if "javitasi_potencial_ft" in hist.columns:
                st.plotly_chart(px.line(hist, x="week", y="javitasi_potencial_ft", markers=True, title="Javítási potenciál trend"), use_container_width=True)

        st.markdown("### Automatikus trendmegállapítás")
        if len(hist) >= 2:
            first, last = hist.iloc[0], hist.iloc[-1]
            notes = []
            if pd.notna(first.get("oee")) and pd.notna(last.get("oee")):
                diff = last["oee"] - first["oee"]
                notes.append(("success" if diff >= 0 else "warning", f"OEE változás: {diff:.1f} pont."))
            if pd.notna(first.get("selejt_pct")) and pd.notna(last.get("selejt_pct")):
                diff = last["selejt_pct"] - first["selejt_pct"]
                notes.append(("success" if diff <= 0 else "danger", f"Selejt változás: {diff:.2f} százalékpont."))
            if pd.notna(first.get("javitasi_potencial_ft")) and pd.notna(last.get("javitasi_potencial_ft")):
                diff = last["javitasi_potencial_ft"] - first["javitasi_potencial_ft"]
                notes.append(("success" if diff <= 0 else "warning", f"Javítási potenciál változás: {fmt_huf(diff)}."))
            render_recommendations(notes)


# ------------------------------------------------------------
# 8. Megrendelések
# ------------------------------------------------------------
with tabs[9]:
    st.subheader("Megrendelésállomány")
    st.caption("Opcionális munkalap: Megrendelesek. Ha feltöltöd, a gyártási terv automatikusan ebből indul.")

    if orders_df is None or orders_df.empty:
        st.info("Nincs feltöltött Megrendelesek munkalap.")
        st.markdown("### Várt oszlopok")
        st.write(OPTIONAL_ORDER_COLS)
    else:
        st.markdown("### Prioritási sorrend")
        priority_orders = order_priority_view(orders_df)
        st.dataframe(priority_orders, use_container_width=True, hide_index=True)

        st.markdown("### Termékenkénti rendelési igény")
        order_summary = orders_df.groupby("Termék", as_index=False).agg(
            Rendelt_db=("Rendelt_db", "sum"),
            Rendelések_száma=("Rendelés_ID", "count"),
            Legkorábbi_határidő=("Határidő", "min")
        )
        st.dataframe(order_summary, use_container_width=True, hide_index=True)

        fig = px.bar(order_summary, x="Termék", y="Rendelt_db", color="Termék", title="Rendelési igény termékenként")
        st.plotly_chart(fig, use_container_width=True)


# ------------------------------------------------------------
# 7. Adatellenőrzés
# ------------------------------------------------------------
with tabs[10]:
    st.subheader("Adatellenőrzés")
    st.markdown("### Feldolgozott adatok")
    st.dataframe(filtered.head(500), use_container_width=True, hide_index=True)

    st.markdown("### Oszlopok")
    st.write(list(filtered.columns))
