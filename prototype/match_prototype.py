"""Prototipo di abbinamento fattura Grimaldi <-> imbarchi Intermodale.

Uso:
    python prototype/match_prototype.py FATTURA.xls IMBARCHI.xlsx [--out risultato.xlsx]

- FATTURA.xls: export Grimaldi così come arriva (.xls, prima riga vuota).
- IMBARCHI.xlsx: export Intermodale con i fogli 'ESP' e 'IT'.

Non è il codice definitivo: serve a riprodurre l'analisi fatta sui dati di
settembre 2026 e come base per il motore vero (vedi docs/HANDOFF.md).
"""
import argparse
import re

import numpy as np
import pandas as pd

# Codici porto usati in 'Linea Porti' (Intermodale) -> nomi porto Grimaldi
PORTS = {
    "SAG": "SAGUNTO", "SAL": "SALERNO", "BCN": "BARCELONA", "SAV": "SAVONA",
    "VAL": "VALENCIA", "LIV": "LIVORNO", "CIV": "CIVITAVECCHIA",
    "CAG": "CAGLIARI", "PTS": "PORTO TORRES",
}
WINDOW_DAYS = 7

# Commodity Grimaldi -> commodity Intermodale compatibili (ricavato da settembre 2026,
# da validare con l'ufficio). Serve sia per scegliere tra candidati sia per segnalare
# righe fatturate con una merce diversa (possibile reclamo, es. vuoto fatturato come pieno).
COMMODITY_OK = {
    "Empty Boxes": {"EMPTY BOX"},
    "<EMPTY>": {"VUOTO", "TRAILER VUOTO"},
    "Glass": {"SPECIALE VETRO 10.50"},
    "Hazardous Chemicals": {"HAZARDOUS CHEMICALS"},
    "Industrial": {"MERCE INDUSTRIALE", "CONSUMER GOODS/PROMO"},
    "Industrial Equipment": {"MERCE INDUSTRIALE"},
    "Miscellaneous / Various": {"CONSUMER GOODS/PROMO"},
    "Plastics": {"PLASTICS"},
    "Spare parts excluding vehicle(s)": {"SPARE PARTS"},
    "Tin plate, Rolled steal": {"BANDA STAGNATA"},
}


def commodity_ok(grimaldi, intermodale):
    return intermodale in COMMODITY_OK.get(grimaldi, set())


def norm_plate(value):
    """'XA 693 EP/SAT' -> 'XA693EP'. I suffissi dopo '/' sono tipologie del mezzo."""
    if pd.isna(value):
        return np.nan
    return re.sub(r"[^A-Z0-9]", "", str(value).upper().split("/")[0]) or np.nan


def load_invoice(path):
    # La prima riga del file Grimaldi è vuota: le intestazioni sono nella riga 2.
    inv = pd.read_excel(path, header=1)
    inv = inv.dropna(how="all")
    inv["plate"] = inv["Chassis/Plate/Cntr No."].map(norm_plate)
    # 'Invoice Date' arriva come yy/mm/dd mascherato da dd/mm/yyyy ("26/09/2007" = 07/09/2026)
    d = inv["Invoice Date"].astype(str).str.extract(r"(\d{2})/(\d{2})/\d{2}(\d{2})")
    inv["invoice_date"] = pd.to_datetime("20" + d[0] + "-" + d[1] + "-" + d[2], errors="coerce")
    inv["rid"] = range(len(inv))
    return inv


def load_shipments(path):
    sheets = [pd.read_excel(path, sheet_name=s).assign(sheet=s) for s in ("ESP", "IT")]
    im = pd.concat(sheets, ignore_index=True)
    im["plate"] = im["S/R"].map(norm_plate)
    ports = im["Linea Porti"].str.split()
    im["POL"] = ports.str[0].map(PORTS)
    im["POD"] = ports.str[1].map(PORTS)
    im["imb"] = pd.to_datetime(
        im["Dt.Imbarco"].astype(str).str.replace(r"\s+", " ", regex=True),
        format="%d/%m/%Y %H:%M", errors="coerce",
    )
    im["iid"] = range(len(im))
    return im


def classify(row):
    if row.pol_ok and row.pod_ok and abs(row.dh) <= 24:
        return "A_esatto"
    if row.pol_ok and row.pod_ok:
        return "B_data_spostata"
    if row.pol_ok or row.pod_ok:
        return "C_un_porto_diverso"
    return "D_porti_diversi"


def match_by_plate(inv, im):
    """Abbinamento 1:1 per targa entro +-WINDOW_DAYS, preferendo porti uguali e data vicina."""
    cand = inv[["rid", "plate", "POL", "POD", "Laden on Board Date"]].merge(
        im[["iid", "plate", "POL", "POD", "imb"]], on="plate", suffixes=("", "_im"))
    cand["dh"] = (cand["imb"] - cand["Laden on Board Date"]).dt.total_seconds() / 3600
    cand = cand[cand["dh"].abs() <= WINDOW_DAYS * 24].copy()
    cand["pol_ok"] = cand["POL"] == cand["POL_im"]
    cand["pod_ok"] = cand["POD"] == cand["POD_im"]
    cand["score"] = cand["dh"].abs() / 24 + (~cand["pol_ok"]) * 3 + (~cand["pod_ok"]) * 2
    used_r, used_i, rows = set(), set(), []
    for r in cand.sort_values("score").itertuples():
        if r.rid in used_r or r.iid in used_i:
            continue
        used_r.add(r.rid)
        used_i.add(r.iid)
        rows.append(r)
    res = pd.DataFrame(rows)
    res["level"] = res.apply(classify, axis=1)
    return res[["rid", "iid", "dh", "level"]]


def suggest_without_plate(inv, im, matched):
    """Per le righe senza abbinamento propone imbarchi Intermodale SENZA targa
    sulla stessa tratta e lo stesso giorno (da confermare a mano)."""
    free_inv = inv[~inv["rid"].isin(matched["rid"])]
    free_im = im[~im["iid"].isin(matched["iid"]) & im["plate"].isna()]
    cand = free_inv[["rid", "POL", "POD", "Laden on Board Date", "Commodity"]].merge(
        free_im[["iid", "POL", "POD", "imb", "commodity"]], on=["POL", "POD"])
    cand["dh"] = (cand["imb"] - cand["Laden on Board Date"]).dt.total_seconds() / 3600
    cand = cand[cand["dh"].abs() <= 48].copy()
    same_goods = [commodity_ok(g, i) for g, i in zip(cand["Commodity"], cand["commodity"])]
    cand["score"] = cand["dh"].abs() / 24 + (~pd.Series(same_goods, index=cand.index)) * 2
    used_r, used_i, rows = set(), set(), []
    for r in cand.sort_values("score").itertuples():
        if r.rid in used_r or r.iid in used_i:
            continue
        used_r.add(r.rid)
        used_i.add(r.iid)
        rows.append({"rid": r.rid, "iid": r.iid, "dh": r.dh, "level": "S_proposta_senza_targa"})
    return pd.DataFrame(rows, columns=["rid", "iid", "dh", "level"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("invoice")
    ap.add_argument("shipments")
    ap.add_argument("--out")
    args = ap.parse_args()

    inv = load_invoice(args.invoice)
    im = load_shipments(args.shipments)
    matched = match_by_plate(inv, im)
    suggested = suggest_without_plate(inv, im, matched)
    links = pd.concat([matched, suggested], ignore_index=True)

    out = inv.merge(links, on="rid", how="left").merge(
        im[["iid", "Prot.", "sheet", "Linea Porti", "Dt.Imbarco", "Dt. Carico", "Tipo", "commodity", "Cliente"]],
        on="iid", how="left")
    out["level"] = out["level"].fillna("X_non_abbinata")
    linked = out["iid"].notna()
    out["commodity_diversa"] = linked & ~pd.Series(
        [commodity_ok(g, i) for g, i in zip(out["Commodity"], out["commodity"])], index=out.index)

    print(f"Righe fattura: {len(inv)}  | righe Intermodale: {len(im)} (senza targa: {im['plate'].isna().sum()})")
    print(out["level"].value_counts().sort_index().to_string())
    print(f"Righe abbinate con commodity diversa da Intermodale: {out['commodity_diversa'].sum()}")
    unused = im[~im["iid"].isin(links["iid"])]
    print(f"Imbarchi Intermodale non usati: {len(unused)}")

    if args.out:
        with pd.ExcelWriter(args.out) as xw:
            out.drop(columns=["rid", "iid"]).to_excel(xw, sheet_name="Fattura_abbinata", index=False)
            unused.drop(columns=["iid"]).to_excel(xw, sheet_name="Intermodale_non_usati", index=False)
        print("Scritto", args.out)


if __name__ == "__main__":
    main()
