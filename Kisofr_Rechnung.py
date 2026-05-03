import io
import math
import re
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


ROLLI_TEILER = 16
TKT_TEILER = 12.85

TKT_SUFFIXE = {
    "2221",
    "2222",
    "2223",
    "4444",
    "7773",
    "7778",
    "7779",
}

TKT_EXAKT = {
    "1058",
    "2058",
    "3058",
    "4058",
    "5058",
    "6030",
    "12221",
    "12222",
    "12223",
    "14444",
    "17773",
    "17778",
    "17779",
    "27779",
}


def norm_text(value) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\ufeff", "")
    text = text.strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_number_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0") or text.endswith(",0"):
        text = text[:-2]
    return text


def normalize_tour(value) -> str:
    return re.sub(r"\D", "", normalize_number_text(value))


def is_tkt_tour(value) -> bool:
    tour = normalize_tour(value)
    if tour in TKT_EXAKT:
        return True
    if len(tour) >= 5 and tour[-4:] in TKT_SUFFIXE:
        return True
    return False


def to_number(value) -> float:
    if pd.isna(value):
        return 0.0

    text = str(value).strip()

    if text == "":
        return 0.0

    if "," in text:
        text = text.replace(".", "").replace(",", ".")

    text = re.sub(r"[^0-9.\-]", "", text)

    if text in ("", "-", ".", "-."):
        return 0.0

    try:
        return float(text)
    except ValueError:
        return 0.0


def ceil_positive(value: float) -> int:
    if value <= 0:
        return 0
    return int(math.ceil(value))


def calculate_quantity(tournummer, e2, e1, karton) -> int:
    if is_tkt_tour(tournummer):
        return ceil_positive((e2 + e1 + karton) / TKT_TEILER)

    return ceil_positive((e2 + (0.5 * e1) + (0.5 * karton)) / ROLLI_TEILER)


def find_column(df: pd.DataFrame, candidates: list[str], required_name: str) -> str:
    normalized_columns = {norm_text(col): col for col in df.columns}

    for candidate in candidates:
        key = norm_text(candidate)
        if key in normalized_columns:
            return normalized_columns[key]

    available = ", ".join(str(col) for col in df.columns)
    raise ValueError(
        f"Spalte nicht gefunden: {required_name}\n\n"
        f"Gesuchte Varianten: {', '.join(candidates)}\n\n"
        f"Vorhandene Spalten: {available}"
    )


def read_csv_file(uploaded_file) -> pd.DataFrame:
    raw = uploaded_file.getvalue()

    for encoding in ["utf-8-sig", "utf-8", "cp1252", "latin1"]:
        try:
            return pd.read_csv(
                io.BytesIO(raw),
                sep=";",
                encoding=encoding,
                dtype=str,
                keep_default_na=False,
            )
        except Exception:
            pass

    raise ValueError("CSV konnte nicht gelesen werden.")


def prepare_data(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    col_tour = find_column(df, ["CSB Tournummer", "CSB Tour Nummer", "Tournummer CSB"], "CSB Tournummer")
    col_csb_customer = find_column(df, ["CSB Kundennummer", "CSB Kunden Nummer", "CSB Kunden-Nr.", "CSB Kundennr"], "CSB Kundennummer")
    col_name = find_column(df, ["Kundenname", "Name", "Kunde"], "Kundenname")
    col_city = find_column(df, ["Stadt", "Ort"], "Stadt")
    col_street = find_column(df, ["Straße", "Strasse", "Str.", "Adresse"], "Straße")
    col_e2 = find_column(df, ["Anzahl gepl. E2", "Anzahl gepl E2", "E2"], "Anzahl gepl. E2")
    col_e1 = find_column(df, ["Anzahl gepl. E1", "Anzahl gepl E1", "E1"], "Anzahl gepl. E1")
    col_karton = find_column(df, ["Anzahl gepl. KARTON", "Anzahl gepl KARTON", "Anzahl gepl. KT", "Anzahl gepl KT", "KARTON", "KT"], "Anzahl gepl. KARTON")

    work = pd.DataFrame(
        {
            "CSB Tournummer": df[col_tour].map(normalize_number_text),
            "CSB Kundennummer": df[col_csb_customer].map(normalize_number_text),
            "Kundenname": df[col_name].astype(str).str.strip(),
            "Stadt": df[col_city].astype(str).str.strip(),
            "Straße": df[col_street].astype(str).str.strip(),
            "Anzahl gepl. E2": df[col_e2].map(to_number),
            "Anzahl gepl. E1": df[col_e1].map(to_number),
            "Anzahl gepl. KARTON": df[col_karton].map(to_number),
        }
    )

    customer_rows = (
        work.groupby(
            ["CSB Tournummer", "CSB Kundennummer", "Kundenname", "Stadt", "Straße"],
            dropna=False,
            as_index=False,
        )[["Anzahl gepl. E2", "Anzahl gepl. E1", "Anzahl gepl. KARTON"]]
        .sum()
    )

    customer_rows["Einheit"] = customer_rows["CSB Tournummer"].apply(lambda tour: "TKT" if is_tkt_tour(tour) else "Rolli")

    # Pro Kundenzeile berechnet. Diese Spalte ist für den Kundenfilter gedacht.
    customer_rows["Anzahl Rolli/TKT je Kunde"] = customer_rows.apply(
        lambda row: calculate_quantity(
            row["CSB Tournummer"],
            row["Anzahl gepl. E2"],
            row["Anzahl gepl. E1"],
            row["Anzahl gepl. KARTON"],
        ),
        axis=1,
    )

    customer_rows = customer_rows[
        [
            "CSB Tournummer",
            "CSB Kundennummer",
            "Kundenname",
            "Stadt",
            "Straße",
            "Anzahl gepl. E2",
            "Anzahl gepl. E1",
            "Anzahl gepl. KARTON",
            "Einheit",
            "Anzahl Rolli/TKT je Kunde",
        ]
    ].sort_values(
        by=["CSB Tournummer", "CSB Kundennummer", "Kundenname"],
        kind="stable",
    ).reset_index(drop=True)

    # Übersicht:
    # Erst Tour summieren, dann genau einmal berechnen und aufrunden.
    # Dadurch ergibt dein Beispiel 183 + 66 + 71 = 320 TKT.
    overview = (
        customer_rows.groupby(["CSB Tournummer", "Einheit"], dropna=False, as_index=False)
        .agg(
            Kunden=("CSB Kundennummer", "nunique"),
            **{
                "Anzahl gepl. E2": ("Anzahl gepl. E2", "sum"),
                "Anzahl gepl. E1": ("Anzahl gepl. E1", "sum"),
                "Anzahl gepl. KARTON": ("Anzahl gepl. KARTON", "sum"),
            },
        )
    )

    overview["Anzahl Rolli/TKT je Tour"] = overview.apply(
        lambda row: calculate_quantity(
            row["CSB Tournummer"],
            row["Anzahl gepl. E2"],
            row["Anzahl gepl. E1"],
            row["Anzahl gepl. KARTON"],
        ),
        axis=1,
    )

    overview = overview[
        [
            "CSB Tournummer",
            "Einheit",
            "Kunden",
            "Anzahl gepl. E2",
            "Anzahl gepl. E1",
            "Anzahl gepl. KARTON",
            "Anzahl Rolli/TKT je Tour",
        ]
    ].sort_values(
        by=["CSB Tournummer", "Einheit"],
        kind="stable",
    ).reset_index(drop=True)

    return customer_rows, overview


def format_sheet(worksheet):
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True, color="000000")
    thin_side = Side(style="thin", color="D0D0D0")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    worksheet.freeze_panes = "A2"

    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    for col_idx in range(1, worksheet.max_column + 1):
        column_letter = get_column_letter(col_idx)
        header = str(worksheet.cell(row=1, column=col_idx).value or "")

        max_length = len(header)
        for cell in worksheet[column_letter]:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))

        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 35)

        if header in {
            "Anzahl gepl. E2",
            "Anzahl gepl. E1",
            "Anzahl gepl. KARTON",
            "Anzahl Rolli/TKT je Kunde",
            "Anzahl Rolli/TKT je Tour",
            "Kunden",
        }:
            for cell in worksheet[column_letter][1:]:
                cell.number_format = "0"


def build_excel(customer_rows: pd.DataFrame, overview: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        customer_rows.to_excel(writer, index=False, sheet_name="Kunden je CSB Tour")
        overview.to_excel(writer, index=False, sheet_name="Übersicht je CSB Tour")

        for worksheet in writer.book.worksheets:
            format_sheet(worksheet)

    output.seek(0)
    return output.getvalue()


st.set_page_config(page_title="Auftragspool zu Excel", layout="centered")

st.title("Auftragspool zu Excel")

uploaded_file = st.file_uploader("CSV hochladen", type=["csv"])

if uploaded_file is not None:
    try:
        df_raw = read_csv_file(uploaded_file)
        customer_rows, overview = prepare_data(df_raw)
        excel_bytes = build_excel(customer_rows, overview)

        filename = f"Auftragspool_CSB_Tournummer_Kunden_Mengen_{date.today():%Y_%m_%d}.xlsx"

        st.download_button(
            label="Excel herunterladen",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as error:
        st.error(str(error))
