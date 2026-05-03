import io
import math
import re
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# =============================================================================
# Einstellung
# =============================================================================
st.set_page_config(
    page_title="Auftragspool",
    page_icon="📦",
    layout="centered",
)

st.markdown(
    """
    <style>
        .block-container {
            max-width: 760px;
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        h1 {
            font-size: 1.45rem;
            letter-spacing: -0.03em;
            margin-bottom: 0.2rem;
        }
        [data-testid="stFileUploader"] {
            padding: 0.2rem 0;
        }
        .stDownloadButton button {
            height: 3rem;
            font-weight: 700;
            border-radius: 0.7rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# Feste Regeln
# =============================================================================
NUMBER_COLUMNS = [
    "Anzahl gepl. E2",
    "Anzahl gepl. E1",
    "Anzahl gepl. KARTON",
]

OUTPUT_COLUMNS = [
    "CSB Tournummer",
    "Einheit",
    "CSB Kundennummer",
    "Kundenname",
    "Stadt",
    "Straße",
    "Anzahl gepl. E2",
    "Anzahl gepl. E1",
    "Anzahl gepl. KARTON",
    "Anzahl Rolli/TKT",
]

CALCULATED_COLUMN = "Anzahl Rolli/TKT"

# Rolli: E2 voll, E1 halb, KARTON halb, geteilt durch 16, danach aufrunden.
ROLLI_DIVISOR = 16.0

# TKT: E2 + E1 + KARTON, geteilt durch 12,85, danach aufrunden.
# Der Faktor 12,85 ist so gewählt, dass diese drei Beispielmengen zusammen 320 TKT ergeben:
# 1457/372/520 = 183, 521/124/202 = 66, 560/165/181 = 71.
TKT_DIVISOR = 12.85

# Diese Touren und Touren mit gleichem Aufbau werden als TKT gerechnet.
TKT_EXACT_TOURS = {
    "12221",
    "12222",
    "12223",
    "14444",
    "17773",
    "17778",
    "17779",
    "27779",
}

TKT_SUFFIXES = {
    "2221",
    "2222",
    "2223",
    "4444",
    "7773",
    "7778",
    "7779",
}

COLUMN_ALIASES: Dict[str, List[str]] = {
    "CSB Tournummer": ["CSB Tournummer", "CSB-Tournummer", "CSB Tour", "Tournummer CSB"],
    "CSB Kundennummer": ["CSB Kundennummer", "CSB-Kundennummer", "CSB Kunden Nummer", "CSB Kunden-Nr", "CSB Kundennr"],
    "Kundenname": ["Kundenname", "Kunden Name", "Name", "Kunde"],
    "Stadt": ["Stadt", "Ort"],
    "Straße": ["Straße", "Strasse", "Str.", "Straße / Hausnummer", "Strasse / Hausnummer"],
    "Anzahl gepl. E2": ["Anzahl gepl. E2", "Anzahl E2", "E2"],
    "Anzahl gepl. E1": ["Anzahl gepl. E1", "Anzahl E1", "E1"],
    "Anzahl gepl. KARTON": ["Anzahl gepl. KARTON", "Anzahl gepl. KT", "Anzahl KARTON", "Anzahl KT", "KARTON", "KT"],
}


# =============================================================================
# Hilfsfunktionen
# =============================================================================
def normalize_column_name(value: object) -> str:
    text = str(value).replace("\ufeff", "").strip()
    text = re.sub(r"\.\d+$", "", text)
    text = text.replace("ß", "ss")
    text = re.sub(r"\s+", " ", text)
    return text.casefold()


def find_required_columns(df: pd.DataFrame) -> Tuple[Dict[str, str], List[str]]:
    normalized_to_original: Dict[str, str] = {}
    for original in df.columns:
        normalized = normalize_column_name(original)
        normalized_to_original.setdefault(normalized, original)

    found: Dict[str, str] = {}
    missing: List[str] = []

    for target, aliases in COLUMN_ALIASES.items():
        match = None
        for alias in aliases:
            normalized_alias = normalize_column_name(alias)
            if normalized_alias in normalized_to_original:
                match = normalized_to_original[normalized_alias]
                break
        if match is None:
            missing.append(target)
        else:
            found[target] = match

    return found, missing


def read_csv_robust(uploaded_file) -> pd.DataFrame:
    raw = uploaded_file.getvalue()
    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    last_error = None

    for encoding in encodings:
        try:
            return pd.read_csv(
                io.BytesIO(raw),
                sep=";",
                encoding=encoding,
                dtype=str,
                keep_default_na=False,
            )
        except Exception as error:
            last_error = error

    raise ValueError(f"Die CSV konnte nicht gelesen werden: {last_error}")


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text


def clean_number_as_text(value: object) -> str:
    text = clean_text(value)
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def parse_number(value: object) -> float:
    text = clean_text(value)
    if not text:
        return 0.0

    text = text.replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return 0.0


def normalize_tour_number(value: object) -> str:
    text = clean_number_as_text(value)
    return re.sub(r"\D", "", text)


def is_tkt_tour(value: object) -> bool:
    digits = normalize_tour_number(value)
    if digits in TKT_EXACT_TOURS:
        return True
    return len(digits) == 5 and digits[1:] in TKT_SUFFIXES


def round_up_full(value: float) -> int:
    if value <= 0:
        return 0
    return int(math.ceil(value))


def sort_key(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.fillna(999999999).astype(float)


def sort_result(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result["__tour_sort"] = sort_key(result["CSB Tournummer"])
    sort_columns = ["__tour_sort", "CSB Tournummer"]

    if "CSB Kundennummer" in result.columns:
        result["__kunde_sort"] = sort_key(result["CSB Kundennummer"])
        sort_columns.extend(["__kunde_sort", "CSB Kundennummer"])

    result = result.sort_values(sort_columns, kind="mergesort")
    drop_columns = [column for column in ["__tour_sort", "__kunde_sort"] if column in result.columns]
    return result.drop(columns=drop_columns).reset_index(drop=True)


def calculate_unit_row(row: pd.Series) -> int:
    e2 = float(row["Anzahl gepl. E2"])
    e1 = float(row["Anzahl gepl. E1"])
    karton = float(row["Anzahl gepl. KARTON"])

    if row["Einheit"] == "TKT":
        return round_up_full((e2 + e1 + karton) / TKT_DIVISOR)

    return round_up_full((e2 + e1 * 0.5 + karton * 0.5) / ROLLI_DIVISOR)


# =============================================================================
# Daten vorbereiten
# =============================================================================
def prepare_data(df: pd.DataFrame):
    found_columns, missing_columns = find_required_columns(df)
    if missing_columns:
        return None, None, missing_columns

    work = pd.DataFrame({target: df[source] for target, source in found_columns.items()})

    for column in ["CSB Tournummer", "Kundenname", "Stadt", "Straße"]:
        work[column] = work[column].apply(clean_text)

    work["CSB Kundennummer"] = work["CSB Kundennummer"].apply(clean_number_as_text)

    for column in NUMBER_COLUMNS:
        work[column] = work[column].apply(parse_number)

    grouped = (
        work.groupby(
            ["CSB Tournummer", "CSB Kundennummer", "Kundenname", "Stadt", "Straße"],
            dropna=False,
            as_index=False,
        )[NUMBER_COLUMNS]
        .sum()
    )

    for column in NUMBER_COLUMNS:
        grouped[column] = grouped[column].round(0).astype(int)

    grouped["Einheit"] = grouped["CSB Tournummer"].apply(lambda value: "TKT" if is_tkt_tour(value) else "Rolli")
    grouped[CALCULATED_COLUMN] = grouped.apply(calculate_unit_row, axis=1).astype(int)
    grouped = sort_result(grouped)[OUTPUT_COLUMNS]

    overview = (
        grouped.groupby(["CSB Tournummer", "Einheit"], dropna=False, as_index=False)[NUMBER_COLUMNS + [CALCULATED_COLUMN]]
        .sum()
    )

    customer_counts = (
        grouped.groupby(["CSB Tournummer", "Einheit"], dropna=False)["CSB Kundennummer"]
        .nunique()
        .reset_index(name="Anzahl Kunden")
    )

    overview = overview.merge(customer_counts, on=["CSB Tournummer", "Einheit"], how="left")
    overview = overview[["CSB Tournummer", "Einheit", "Anzahl Kunden"] + NUMBER_COLUMNS + [CALCULATED_COLUMN]]
    overview = sort_result(overview)

    for column in ["Anzahl Kunden"] + NUMBER_COLUMNS + [CALCULATED_COLUMN]:
        if column in overview.columns:
            overview[column] = overview[column].fillna(0).astype(int)

    return grouped, overview, []


# =============================================================================
# Excel erzeugen
# =============================================================================
def autosize_worksheet(worksheet) -> None:
    for column_cells in worksheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        max_length = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 42)


def style_worksheet(worksheet) -> None:
    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(color="FFFFFF", bold=True)
    even_fill = PatternFill("solid", fgColor="F9FAFB")
    tkt_fill = PatternFill("solid", fgColor="FEF3C7")
    thin_gray = Side(style="thin", color="D1D5DB")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    worksheet.freeze_panes = "A2"

    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        last_column = get_column_letter(worksheet.max_column)
        worksheet.auto_filter.ref = f"A1:{last_column}{worksheet.max_row}"

    headers = [cell.value for cell in worksheet[1]]
    einheit_index = headers.index("Einheit") + 1 if "Einheit" in headers else None

    for row in worksheet.iter_rows():
        is_tkt_row = False
        if einheit_index and row[0].row > 1:
            is_tkt_row = worksheet.cell(row=row[0].row, column=einheit_index).value == "TKT"

        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=False)
            if cell.row == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif is_tkt_row:
                cell.fill = tkt_fill
            elif cell.row % 2 == 0:
                cell.fill = even_fill

    worksheet.row_dimensions[1].height = 30

    number_headers = set(NUMBER_COLUMNS + [CALCULATED_COLUMN, "Anzahl Kunden"])
    for header_cell in worksheet[1]:
        if header_cell.value in number_headers:
            column_letter = get_column_letter(header_cell.column)
            for value_cell in worksheet[column_letter][1:]:
                value_cell.number_format = "0"
                value_cell.alignment = Alignment(horizontal="right", vertical="top")

    autosize_worksheet(worksheet)


def build_excel(grouped: pd.DataFrame, overview: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        grouped.to_excel(writer, sheet_name="Kunden je CSB Tour", index=False)
        overview.to_excel(writer, sheet_name="Übersicht je CSB Tour", index=False)

        for worksheet in writer.book.worksheets:
            style_worksheet(worksheet)

    output.seek(0)
    return output.getvalue()


# =============================================================================
# Oberfläche: nur Hochladen und Herunterladen
# =============================================================================
st.title("Auftragspool")

uploaded_file = st.file_uploader(
    "CSV hochladen",
    type=["csv"],
    label_visibility="collapsed",
)

if uploaded_file:
    try:
        raw_df = read_csv_robust(uploaded_file)
        grouped_df, overview_df, missing_columns = prepare_data(raw_df)

        if missing_columns:
            st.error("Diese Spalten fehlen: " + ", ".join(missing_columns))
            st.stop()

        excel_bytes = build_excel(grouped_df, overview_df)
        file_date = datetime.now().strftime("%Y_%m_%d")

        st.download_button(
            "Excel herunterladen",
            data=excel_bytes,
            file_name=f"Auftragspool_CSB_Tournummer_Kunden_Mengen_{file_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as error:
        st.error(f"Fehler beim Verarbeiten: {error}")
