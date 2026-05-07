from __future__ import annotations

import re
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st


DEFAULT_RANKING_DATE = date(2026, 1, 30)
DEFAULT_EXCLUDED_PRODUCTS = ("D400", "D460")

REQUIRED_FACTURACION_COLUMNS = (
    "PRODUCTO",
    "POLIALTA",
    "POLIZA",
    "MEDIADOR",
    "PRIMA NETA",
)

REQUIRED_ANULACIONES_COLUMNS = (
    "PRODUCTO",
    "FECHA EMISION",
    "POLIZA",
    "MEDIADOR",
    "PRIMA NETA",
    "CAUSA",
)

AGENCY_NAME_COLUMNS = (
    "NOMBRE AGENCIA",
    "NOMBRE_AGENCIA",
    "AGENCIA",
    "NOMBRE MEDIADOR",
    "NOM MEDIADOR",
    "MEDIADOR NOMBRE",
)

DEPENDENCY_COLUMNS = (
    "SECTOCOB",
    "SECTOR",
    "DEPENDENCIA",
)


def parse_spanish_number(value: object) -> float:
    if pd.isna(value):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"[^\d,.\-]", "", text)

    if not text or text in {"-", ",", "."}:
        return 0.0

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return 0.0


def normalize_product(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def normalize_text(value: object, default: str = "") -> str:
    if pd.isna(value) or str(value).strip() == "":
        return default

    text = str(value).strip()

    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]

    return text


def normalize_reason_text(value: object) -> str:
    if pd.isna(value):
        return ""

    text = str(value).strip().upper()
    text = text.replace("Á", "A")
    text = text.replace("É", "E")
    text = text.replace("Í", "I")
    text = text.replace("Ó", "O")
    text = text.replace("Ú", "U")

    return text


def normalize_agent(value: object) -> str:
    return normalize_text(value, "Sin mediador")


def first_existing_column(columns: list[str] | pd.Index, options: tuple[str, ...]) -> str | None:
    normalized = {str(column).strip().upper(): str(column).strip() for column in columns}

    for option in options:
        column = normalized.get(option.upper())
        if column is not None:
            return column

    return None


def first_non_empty(values: pd.Series) -> str:
    for value in values:
        if not pd.isna(value) and str(value).strip() != "":
            return str(value).strip()

    return ""


def read_excel_all_sheets(uploaded_file) -> pd.DataFrame:
    sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)
    frames = []

    for sheet_name, sheet_df in sheets.items():
        sheet_df = sheet_df.copy()
        sheet_df.columns = [str(column).strip() for column in sheet_df.columns]
        sheet_df["HOJA_ORIGEN"] = sheet_name
        frames.append(sheet_df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def validate_columns(df: pd.DataFrame, required_columns: tuple[str, ...]) -> list[str]:
    return [column for column in required_columns if column not in df.columns]


def prepare_decesos_data(
    df: pd.DataFrame,
    date_column: str,
    movement: str,
) -> pd.DataFrame:
    work = df.copy()

    agency_name_column = first_existing_column(work.columns, AGENCY_NAME_COLUMNS)
    dependency_column = first_existing_column(work.columns, DEPENDENCY_COLUMNS)

    work["MOVIMIENTO"] = movement
    work["PRODUCTO_NORMALIZADO"] = work["PRODUCTO"].apply(normalize_product)
    work["AGENTE"] = work["MEDIADOR"].apply(normalize_agent)

    if agency_name_column:
        agency_names = work[agency_name_column].apply(lambda value: normalize_text(value, ""))
        work["NOMBRE_AGENCIA"] = [
            agency_name if agency_name else agent
            for agency_name, agent in zip(agency_names, work["AGENTE"])
        ]
    else:
        work["NOMBRE_AGENCIA"] = work["AGENTE"]

    if dependency_column:
        work["DEPENDENCIA"] = work[dependency_column].apply(
            lambda value: normalize_text(value, "Sin dependencia")
        )
    else:
        work["DEPENDENCIA"] = "Sin dependencia"

    work["FECHA_MOVIMIENTO"] = pd.to_datetime(
        work[date_column],
        dayfirst=True,
        errors="coerce",
    )
    work["ANIO_MOVIMIENTO"] = work["FECHA_MOVIMIENTO"].dt.year
    work["MES_MOVIMIENTO"] = work["FECHA_MOVIMIENTO"].dt.month
    work["PRIMA_NETA_VALOR"] = work["PRIMA NETA"].apply(parse_spanish_number)

    return work


def filter_movements(
    df: pd.DataFrame,
    ranking_date: date,
    excluded_products: list[str],
    date_column: str,
    movement: str,
) -> pd.DataFrame:
    work = prepare_decesos_data(df, date_column, movement)
    excluded = {normalize_product(product) for product in excluded_products}

    mask = (
        work["FECHA_MOVIMIENTO"].notna()
        & work["ANIO_MOVIMIENTO"].eq(ranking_date.year)
        & work["MES_MOVIMIENTO"].le(ranking_date.month)
        & ~work["PRODUCTO_NORMALIZADO"].isin(excluded)
    )

    if movement == "ANULACION":
    motivo = work["MOTIVO"].apply(normalize_reason_text) if "MOTIVO" in work.columns else ""
    causa = work["CAUSA"].apply(normalize_reason_text) if "CAUSA" in work.columns else ""

    excluded_by_motivo = (
        motivo.str.contains("DEFUNCION DEL ULTIMO O UNICO ASEGURADO", na=False)
        | motivo.str.contains("DEFUNCION \\(QUEDAN MAS ASEGURADOS PERO NO LA QUIEREN\\)", na=False)
        | motivo.str.contains("SINIESTRO TOTAL", na=False)
        | motivo.str.contains("DEFUNCION", na=False)
    )

    excluded_by_causa = causa.str.contains("INDIVIDUAL POR SINIESTRO", na=False)

    mask = mask & ~excluded_by_motivo & ~excluded_by_causa

def aggregate_movements(
    detail: pd.DataFrame,
    amount_column: str,
    count_column: str,
) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame(
            columns=[
                "AGENTE",
                "DEPENDENCIA",
                "NOMBRE_AGENCIA",
                amount_column,
                count_column,
            ]
        )

    return (
        detail.groupby(["AGENTE", "DEPENDENCIA"], dropna=False)
        .agg(
            NOMBRE_AGENCIA=("NOMBRE_AGENCIA", first_non_empty),
            **{
                amount_column: ("PRIMA_NETA_VALOR", "sum"),
                count_column: ("POLIZA", "count"),
            },
        )
        .reset_index()
    )


def calculate_facturacion_neta(
    facturacion_df: pd.DataFrame,
    anulaciones_df: pd.DataFrame,
    ranking_date: date,
    excluded_products: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    altas_detail = filter_movements(
        facturacion_df,
        ranking_date,
        excluded_products,
        "POLIALTA",
        "ALTA",
    )

    anulaciones_detail = filter_movements(
        anulaciones_df,
        ranking_date,
        excluded_products,
        "FECHA EMISION",
        "ANULACION",
    )

    altas = aggregate_movements(
        altas_detail,
        "FACTURACION_ALTAS_BRUTAS",
        "POLIZAS_ALTAS",
    )

    anulaciones = aggregate_movements(
        anulaciones_detail,
        "FACTURACION_ANULACIONES",
        "POLIZAS_ANULADAS",
    )

    ranking = pd.merge(
        altas,
        anulaciones,
        on=["AGENTE", "DEPENDENCIA"],
        how="outer",
        suffixes=("_ALTAS", "_ANULACIONES"),
    )

    if ranking.empty:
        empty = pd.DataFrame(
            columns=[
                "RANKING",
                "AGENTE",
                "NOMBRE_AGENCIA",
                "DEPENDENCIA",
                "FACTURACION_ALTAS_BRUTAS",
                "FACTURACION_ANULACIONES",
                "FACTURACION_NETA",
                "POLIZAS_ALTAS",
                "POLIZAS_ANULADAS",
                "POLIZAS_NETAS",
                "PRIMA_MEDIA_NETA",
            ]
        )
        return empty, altas_detail, anulaciones_detail

    ranking["NOMBRE_AGENCIA"] = ranking["NOMBRE_AGENCIA_ALTAS"].combine_first(
        ranking["NOMBRE_AGENCIA_ANULACIONES"]
    )

    ranking["NOMBRE_AGENCIA"] = [
        name if not pd.isna(name) and str(name).strip() != "" else agent
        for name, agent in zip(ranking["NOMBRE_AGENCIA"], ranking["AGENTE"])
    ]

    numeric_columns = [
        "FACTURACION_ALTAS_BRUTAS",
        "FACTURACION_ANULACIONES",
        "POLIZAS_ALTAS",
        "POLIZAS_ANULADAS",
    ]

    for column in numeric_columns:
        ranking[column] = ranking[column].fillna(0)

    ranking["FACTURACION_NETA"] = (
        ranking["FACTURACION_ALTAS_BRUTAS"] - ranking["FACTURACION_ANULACIONES"]
    )
    ranking["POLIZAS_NETAS"] = ranking["POLIZAS_ALTAS"] - ranking["POLIZAS_ANULADAS"]
    ranking["PRIMA_MEDIA_NETA"] = [
        facturacion / polizas if polizas else 0.0
        for facturacion, polizas in zip(
            ranking["FACTURACION_NETA"],
            ranking["POLIZAS_NETAS"],
        )
    ]

    ranking = ranking[
        [
            "AGENTE",
            "NOMBRE_AGENCIA",
            "DEPENDENCIA",
            "FACTURACION_ALTAS_BRUTAS",
            "FACTURACION_ANULACIONES",
            "FACTURACION_NETA",
            "POLIZAS_ALTAS",
            "POLIZAS_ANULADAS",
            "POLIZAS_NETAS",
            "PRIMA_MEDIA_NETA",
        ]
    ].sort_values(
        ["FACTURACION_NETA", "FACTURACION_ALTAS_BRUTAS", "AGENTE", "DEPENDENCIA"],
        ascending=[False, False, True, True],
    )

    ranking.insert(0, "RANKING", range(1, len(ranking) + 1))

    return ranking, altas_detail, anulaciones_detail


def build_sheet_summary(df: pd.DataFrame, file_name: str) -> pd.DataFrame:
    if "HOJA_ORIGEN" not in df.columns:
        return pd.DataFrame(columns=["ARCHIVO", "HOJA", "FILAS_LEIDAS"])

    summary = (
        df.groupby("HOJA_ORIGEN", dropna=False)
        .size()
        .reset_index(name="FILAS_LEIDAS")
        .rename(columns={"HOJA_ORIGEN": "HOJA"})
    )

    summary.insert(0, "ARCHIVO", file_name)

    return summary


def format_euro(value: float) -> str:
    return f"{value:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def detail_columns_for_display(detail: pd.DataFrame) -> list[str]:
    columns = [
        "HOJA_ORIGEN",
        "MOVIMIENTO",
        "PRODUCTO",
        "POLIALTA",
        "FECHA BAJA",
        "FECHA EMISION",
        "FECHA GRABACION",
        "CAUSA",
        "MOTIVO",
        "POLIZA",
        "MEDIADOR",
        "SECTOCOB",
        "SECTOR",
        "NOMBRE_AGENCIA",
        "DEPENDENCIA",
        "PRIMA NETA",
        "PRIMA_NETA_VALOR",
        "ANIO_MOVIMIENTO",
        "MES_MOVIMIENTO",
    ]

    return [column for column in columns if column in detail.columns]


def dataframe_to_excel(
    ranking: pd.DataFrame,
    altas_detail: pd.DataFrame,
    anulaciones_detail: pd.DataFrame,
    sheet_summary: pd.DataFrame,
    ranking_date: date,
    excluded_products: list[str],
) -> bytes:
    output = BytesIO()

    parametros = pd.DataFrame(
        [
            {"CAMPO": "FECHA_RANKING", "VALOR": ranking_date.strftime("%d/%m/%Y")},
            {"CAMPO": "ANIO", "VALOR": ranking_date.year},
            {"CAMPO": "MES_HASTA", "VALOR": ranking_date.month},
            {"CAMPO": "PRODUCTOS_EXCLUIDOS", "VALOR": ", ".join(excluded_products)},
            {"CAMPO": "FECHA_ANULACIONES", "VALOR": "FECHA EMISION"},
            {"CAMPO": "ANULACIONES_EXCLUIDAS_CAUSA", "VALOR": "DEFUNCION, DEFUNCIÓN"},
        ]
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        parametros.to_excel(writer, index=False, sheet_name="PARAMETROS")
        sheet_summary.to_excel(writer, index=False, sheet_name="COMPROBACION_HOJAS")
        ranking.to_excel(writer, index=False, sheet_name="RANKING_NETO")
        altas_detail[detail_columns_for_display(altas_detail)].to_excel(
            writer,
            index=False,
            sheet_name="DETALLE_ALTAS",
        )
        anulaciones_detail[detail_columns_for_display(anulaciones_detail)].to_excel(
            writer,
            index=False,
            sheet_name="DETALLE_ANULACIONES",
        )

    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="Ranking agentes", layout="wide")

    st.title("Ranking agentes - DECESOS")

    ranking_date = st.date_input(
        "Para que fecha quieres calcular el ranking?",
        value=DEFAULT_RANKING_DATE,
        format="DD/MM/YYYY",
    )

    col_upload_1, col_upload_2 = st.columns(2)

    with col_upload_1:
        uploaded_facturacion = st.file_uploader(
            "Sube FACTURACION_DECESOS.xls",
            type=["xls", "xlsx"],
            key="facturacion",
        )

    with col_upload_2:
        uploaded_anulaciones = st.file_uploader(
            "Sube FACTURACION_ANULACIONES_DECESOS.xls",
            type=["xls", "xlsx"],
            key="anulaciones",
        )

    if uploaded_facturacion is None or uploaded_anulaciones is None:
        st.info("Sube los dos archivos para calcular altas, anulaciones y facturacion neta.")
        st.stop()

    try:
        raw_facturacion_df = read_excel_all_sheets(uploaded_facturacion)
        raw_anulaciones_df = read_excel_all_sheets(uploaded_anulaciones)
    except Exception as error:
        st.error(f"No he podido leer los archivos: {error}")
        st.stop()

    if raw_facturacion_df.empty or raw_anulaciones_df.empty:
        st.warning("Alguno de los archivos esta vacio.")
        st.stop()

    missing_facturacion = validate_columns(
        raw_facturacion_df,
        REQUIRED_FACTURACION_COLUMNS,
    )
    missing_anulaciones = validate_columns(
        raw_anulaciones_df,
        REQUIRED_ANULACIONES_COLUMNS,
    )

    if missing_facturacion or missing_anulaciones:
        messages = []

        if missing_facturacion:
            messages.append(f"Facturacion: {', '.join(missing_facturacion)}")

        if missing_anulaciones:
            messages.append(f"Anulaciones: {', '.join(missing_anulaciones)}")

        st.error("Faltan columnas obligatorias. " + " | ".join(messages))
        st.stop()

    sheet_summary = pd.concat(
        [
            build_sheet_summary(raw_facturacion_df, "FACTURACION_DECESOS"),
            build_sheet_summary(raw_anulaciones_df, "FACTURACION_ANULACIONES_DECESOS"),
        ],
        ignore_index=True,
    )

    with st.expander("Comprobacion de hojas leidas", expanded=True):
        st.dataframe(sheet_summary, use_container_width=True, hide_index=True)

    product_options = sorted(
        pd.concat(
            [
                raw_facturacion_df["PRODUCTO"],
                raw_anulaciones_df["PRODUCTO"],
            ],
            ignore_index=True,
        )
        .dropna()
        .map(normalize_product)
        .unique()
        .tolist()
    )

    default_excluded = [
        product for product in DEFAULT_EXCLUDED_PRODUCTS if product in product_options
    ]

    excluded_products = st.multiselect(
        "Productos excluidos",
        product_options,
        default=default_excluded,
    )

    ranking, altas_detail, anulaciones_detail = calculate_facturacion_neta(
        raw_facturacion_df,
        raw_anulaciones_df,
        ranking_date,
        excluded_products,
    )

    total_altas = (
        float(ranking["FACTURACION_ALTAS_BRUTAS"].sum()) if not ranking.empty else 0.0
    )
    total_anulaciones = (
        float(ranking["FACTURACION_ANULACIONES"].sum()) if not ranking.empty else 0.0
    )
    total_neta = float(ranking["FACTURACION_NETA"].sum()) if not ranking.empty else 0.0
    total_polizas_netas = int(ranking["POLIZAS_NETAS"].sum()) if not ranking.empty else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Facturacion altas brutas", format_euro(total_altas))
    col2.metric("Facturacion anulaciones", format_euro(total_anulaciones))
    col3.metric("Facturacion neta", format_euro(total_neta))
    col4.metric("Polizas netas", f"{total_polizas_netas:,}".replace(",", "."))

    st.subheader("Ranking facturacion neta")

    if ranking.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        st.dataframe(
            ranking.style.format(
                {
                    "FACTURACION_ALTAS_BRUTAS": lambda value: format_euro(float(value)),
                    "FACTURACION_ANULACIONES": lambda value: format_euro(float(value)),
                    "FACTURACION_NETA": lambda value: format_euro(float(value)),
                    "PRIMA_MEDIA_NETA": lambda value: format_euro(float(value)),
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("Detalle de altas incluidas"):
        st.dataframe(
            altas_detail[detail_columns_for_display(altas_detail)].style.format(
                {"PRIMA_NETA_VALOR": lambda value: format_euro(float(value))}
            ),
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("Detalle de anulaciones incluidas"):
        st.dataframe(
            anulaciones_detail[detail_columns_for_display(anulaciones_detail)].style.format(
                {"PRIMA_NETA_VALOR": lambda value: format_euro(float(value))}
            ),
            use_container_width=True,
            hide_index=True,
        )

    excel_bytes = dataframe_to_excel(
        ranking,
        altas_detail,
        anulaciones_detail,
        sheet_summary,
        ranking_date,
        excluded_products,
    )

    st.download_button(
        "Descargar ranking neto en Excel",
        data=excel_bytes,
        file_name=f"ranking_decesos_facturacion_neta_{ranking_date:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
