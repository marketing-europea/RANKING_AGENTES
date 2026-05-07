from __future__ import annotations

import re
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st


DEFAULT_RANKING_DATE = date(2026, 1, 30)
DEFAULT_EXCLUDED_PRODUCTS = ("D400", "D460")

REQUIRED_COLUMNS = (
    "PRODUCTO",
    "POLIALTA",
    "POLIZA",
    "MEDIADOR",
    "PRIMA NETA",
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


def normalize_agent(value: object) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return "Sin mediador"

    text = str(value).strip()

    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]

    return text


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


def validate_columns(df: pd.DataFrame) -> list[str]:
    return [column for column in REQUIRED_COLUMNS if column not in df.columns]


def prepare_decesos_data(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    work["PRODUCTO_NORMALIZADO"] = work["PRODUCTO"].apply(normalize_product)
    work["AGENTE"] = work["MEDIADOR"].apply(normalize_agent)
    work["FECHA_ALTA"] = pd.to_datetime(work["POLIALTA"], dayfirst=True, errors="coerce")
    work["ANIO_ALTA"] = work["FECHA_ALTA"].dt.year
    work["MES_ALTA"] = work["FECHA_ALTA"].dt.month
    work["PRIMA_NETA_VALOR"] = work["PRIMA NETA"].apply(parse_spanish_number)

    return work


def calculate_facturacion_altas_brutas(
    df: pd.DataFrame,
    ranking_date: date,
    excluded_products: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    work = prepare_decesos_data(df)
    excluded = {normalize_product(product) for product in excluded_products}

    mask = (
        work["FECHA_ALTA"].notna()
        & work["ANIO_ALTA"].eq(ranking_date.year)
        & work["MES_ALTA"].le(ranking_date.month)
        & ~work["PRODUCTO_NORMALIZADO"].isin(excluded)
    )

    detail = work.loc[mask].copy()

    ranking = (
        detail.groupby("AGENTE", dropna=False)
        .agg(
            FACTURACION_ALTAS_BRUTAS=("PRIMA_NETA_VALOR", "sum"),
            POLIZAS_ALTAS=("POLIZA", "count"),
            PRIMA_MEDIA=("PRIMA_NETA_VALOR", "mean"),
        )
        .reset_index()
        .sort_values(
            ["FACTURACION_ALTAS_BRUTAS", "POLIZAS_ALTAS", "AGENTE"],
            ascending=[False, False, True],
        )
    )

    ranking.insert(0, "RANKING", range(1, len(ranking) + 1))

    return ranking, detail


def format_euro(value: float) -> str:
    return f"{value:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def dataframe_to_excel(
    ranking: pd.DataFrame,
    detail: pd.DataFrame,
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
        ]
    )

    detail_columns = [
        "HOJA_ORIGEN",
        "PRODUCTO",
        "POLIALTA",
        "POLIZA",
        "MEDIADOR",
        "PRIMA NETA",
        "PRIMA_NETA_VALOR",
        "ANIO_ALTA",
        "MES_ALTA",
    ]
    detail_columns = [column for column in detail_columns if column in detail.columns]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        parametros.to_excel(writer, index=False, sheet_name="PARAMETROS")
        ranking.to_excel(writer, index=False, sheet_name="RANKING")
        detail[detail_columns].to_excel(writer, index=False, sheet_name="DETALLE_ALTAS")

    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="Ranking agentes", layout="wide")

    st.title("Ranking agentes - DECESOS")

    ranking_date = st.date_input(
        "Para que fecha quieres calcular el ranking?",
        value=DEFAULT_RANKING_DATE,
        format="DD/MM/YYYY",
    )

    uploaded_file = st.file_uploader(
        "Sube FACTURACION_DECESOS.xls",
        type=["xls", "xlsx"],
    )

    if uploaded_file is None:
        st.info("Sube el archivo FACTURACION_DECESOS.xls para calcular el ranking.")
        st.stop()

    try:
        raw_df = read_excel_all_sheets(uploaded_file)
    except Exception as error:
        st.error(f"No he podido leer el archivo: {error}")
        st.stop()

    if raw_df.empty:
        st.warning("El archivo esta vacio.")
        st.stop()

    missing_columns = validate_columns(raw_df)

    if missing_columns:
        st.error(f"Faltan columnas obligatorias: {', '.join(missing_columns)}")
        st.stop()

    product_options = sorted(
        raw_df["PRODUCTO"].dropna().map(normalize_product).unique().tolist()
    )

    default_excluded = [
        product for product in DEFAULT_EXCLUDED_PRODUCTS if product in product_options
    ]

    excluded_products = st.multiselect(
        "Productos excluidos",
        product_options,
        default=default_excluded,
    )

    ranking, detail = calculate_facturacion_altas_brutas(
        raw_df,
        ranking_date,
        excluded_products,
    )

    total_facturacion = (
        float(ranking["FACTURACION_ALTAS_BRUTAS"].sum()) if not ranking.empty else 0.0
    )
    total_polizas = int(ranking["POLIZAS_ALTAS"].sum()) if not ranking.empty else 0
    total_agentes = int(ranking["AGENTE"].nunique()) if not ranking.empty else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Facturacion altas brutas", format_euro(total_facturacion))
    col2.metric("Polizas altas", f"{total_polizas:,}".replace(",", "."))
    col3.metric("Agentes", f"{total_agentes:,}".replace(",", "."))
    col4.metric("Fecha ranking", ranking_date.strftime("%d/%m/%Y"))

    st.subheader("Ranking facturacion altas brutas")

    if ranking.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        st.dataframe(
            ranking.style.format(
                {
                    "FACTURACION_ALTAS_BRUTAS": lambda value: format_euro(float(value)),
                    "PRIMA_MEDIA": lambda value: format_euro(float(value)),
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("Detalle de polizas incluidas"):
        detail_columns = [
            "HOJA_ORIGEN",
            "PRODUCTO",
            "POLIALTA",
            "POLIZA",
            "MEDIADOR",
            "PRIMA NETA",
            "PRIMA_NETA_VALOR",
        ]
        detail_columns = [column for column in detail_columns if column in detail.columns]

        st.dataframe(
            detail[detail_columns].style.format(
                {"PRIMA_NETA_VALOR": lambda value: format_euro(float(value))}
            ),
            use_container_width=True,
            hide_index=True,
        )

    excel_bytes = dataframe_to_excel(
        ranking,
        detail,
        ranking_date,
        excluded_products,
    )

    st.download_button(
        "Descargar ranking en Excel",
        data=excel_bytes,
        file_name=f"ranking_decesos_altas_brutas_{ranking_date:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
if __name__ == "__main__":
    main()
