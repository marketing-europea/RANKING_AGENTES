from __future__ import annotations

import re
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st


DEFAULT_RANKING_DATE = date(2026, 1, 30)
DEFAULT_EXCLUDED_PRODUCTS = ("D600", "D460")
SINIESTRALIDAD_MAXIMA = 0.25

DECESOS_LIGA_PRO_MINIMA = 30000.0
DECESOS_LIGA_ELITE_MINIMA = 60000.0
SALUD_LIGA_PRO_MINIMA = 25000.0
SALUD_LIGA_PRO_DECESOS_MINIMA = 12000.0
SALUD_LIGA_ELITE_MINIMA = 80000.0
SALUD_LIGA_ELITE_DECESOS_MINIMA = 4000.0

COLOR_DECESOS = "#f32735"
COLOR_SALUD = "#5271ff"
COLOR_VIDA = "#ffb4ab"

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

# Siniestros: fecha por FECHOCUR e importe por RESERACT + PAGOSPDT.
REQUIRED_SINIESTROS_COLUMNS = (
    "PRODUCTO",
    "CODIMEDI",
    "FECHOCUR",
    "RESERACT",
    "PAGOSPDT",
)

REQUIRED_PRIMAS_COLUMNS = (
    "MEDIADOR",
    "POLIPNET",
)

REQUIRED_FACTURACION_SALUD_COLUMNS = (
    "PRODUCTO",
    "IDPOLIZA",
    "MEDIADOR",
    "POLIEFEC",
    "PRIMA NETA",
)

REQUIRED_BAJAS_SALUD_COLUMNS = (
    "POLIZA",
    "DES_PRODUCTO",
    "FEC_EFECTO_BAJA",
    "FEC_EFECTO_REACTIV",
)

REQUIRED_MAPEO_COLUMNS = (
    "CODIMEDI",
    "NOMBCOME",
    "Responsable",
)

REQUIRED_FACTURACION_VIDA_COLUMNS = (
    "NUMERO",
    "CODIMEDI",
    "FECHALTA",
    "PRIMATOTAL",
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
        sheet_df["ARCHIVO_ORIGEN"] = getattr(uploaded_file, "name", "archivo")
        sheet_df["HOJA_ORIGEN"] = sheet_name
        frames.append(sheet_df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def read_excel_many_files(uploaded_files) -> pd.DataFrame:
    frames = []

    for uploaded_file in uploaded_files or []:
        file_df = read_excel_all_sheets(uploaded_file)
        if not file_df.empty:
            frames.append(file_df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def validate_columns(df: pd.DataFrame, required_columns: tuple[str, ...]) -> list[str]:
    return [column for column in required_columns if column not in df.columns]


# =========================
# MAPEO MEDIADORES
# =========================

def prepare_mapeo_data(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    work["AGENTE"] = work["CODIMEDI"].apply(normalize_agent)
    work["CODIMEDI"] = work["CODIMEDI"].apply(normalize_agent)
    work["NOMBRE_AGENCIA_MAPEO"] = work["NOMBCOME"].apply(lambda value: normalize_text(value, ""))
    work["RESPONSABLE"] = work["Responsable"].apply(lambda value: normalize_text(value, "Sin responsable"))

    return (
        work[["AGENTE", "CODIMEDI", "NOMBRE_AGENCIA_MAPEO", "RESPONSABLE"]]
        .drop_duplicates("AGENTE")
        .copy()
    )


def add_mapeo_to_ranking(ranking: pd.DataFrame, mapeo: pd.DataFrame) -> pd.DataFrame:
    if ranking.empty:
        return ranking.copy()

    result = ranking.copy()
    result["AGENTE"] = result["AGENTE"].apply(normalize_agent)

    if not mapeo.empty:
        result = pd.merge(result, mapeo, on="AGENTE", how="left")
    else:
        result["CODIMEDI"] = result["AGENTE"]
        result["NOMBRE_AGENCIA_MAPEO"] = ""
        result["RESPONSABLE"] = "Sin responsable"

    result["CODIMEDI"] = result["CODIMEDI"].fillna(result["AGENTE"])
    result["NOMBRE_AGENCIA"] = [
        mapped if not pd.isna(mapped) and str(mapped).strip() != "" else original
        for mapped, original in zip(result["NOMBRE_AGENCIA_MAPEO"], result["NOMBRE_AGENCIA"])
    ]
    result["RESPONSABLE"] = result["RESPONSABLE"].fillna("Sin responsable")

    return result.drop(columns=["NOMBRE_AGENCIA_MAPEO"], errors="ignore")


def add_mapeo_to_simple_ranking(ranking: pd.DataFrame, mapeo: pd.DataFrame) -> pd.DataFrame:
    if ranking.empty:
        return ranking.copy()

    result = ranking.copy()
    result["AGENTE"] = result["AGENTE"].apply(normalize_agent)

    if not mapeo.empty:
        result = pd.merge(result, mapeo, on="AGENTE", how="left")
    else:
        result["CODIMEDI"] = result["AGENTE"]
        result["NOMBRE_AGENCIA_MAPEO"] = ""
        result["RESPONSABLE"] = "Sin responsable"

    result["CODIMEDI"] = result["CODIMEDI"].fillna(result["AGENTE"])
    result["NOMBRE_AGENCIA"] = [
        mapped if not pd.isna(mapped) and str(mapped).strip() != "" else agent
        for mapped, agent in zip(result["NOMBRE_AGENCIA_MAPEO"], result["AGENTE"])
    ]
    result["RESPONSABLE"] = result["RESPONSABLE"].fillna("Sin responsable")

    return result.drop(columns=["NOMBRE_AGENCIA_MAPEO"], errors="ignore")


# =========================
# DECESOS
# =========================

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

    if movement == "ANULACION":
        # Regla de imputacion Decesos:
        # las bajas con FECHA EMISION entre el 1 y el 25 de enero
        # se imputan al ejercicio anterior.
        enero_hasta_25 = (
            work["FECHA_MOVIMIENTO"].notna()
            & work["FECHA_MOVIMIENTO"].dt.month.eq(1)
            & work["FECHA_MOVIMIENTO"].dt.day.lt(25)
        )
        work.loc[enero_hasta_25, "ANIO_MOVIMIENTO"] = (
            work.loc[enero_hasta_25, "ANIO_MOVIMIENTO"] - 1
        )

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
        motivo = (
            work["MOTIVO"].apply(normalize_reason_text)
            if "MOTIVO" in work.columns
            else pd.Series("", index=work.index)
        )

        causa = (
            work["CAUSA"].apply(normalize_reason_text)
            if "CAUSA" in work.columns
            else pd.Series("", index=work.index)
        )

        excluded_by_motivo = (
            motivo.str.startswith("DEFUNCION DEL ULTIMO O UNICO ASEGURADO", na=False)
            | motivo.str.startswith("DEFUNCION (QUEDAN MAS ASEGURADOS PERO NO LA QUIEREN)", na=False)
            | motivo.str.startswith("SINIESTRO TOTAL", na=False)
            | motivo.str.startswith("DEFUNCION", na=False)
        )

        excluded_by_causa = (
            causa.str.startswith("DEFUNCION", na=False)
            | causa.str.startswith("INDIVIDUAL POR SINIESTRO", na=False)
        )

        mask = mask & ~excluded_by_motivo & ~excluded_by_causa

    return work[mask].copy()


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


def prepare_siniestros_data(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    work["PRODUCTO_NORMALIZADO"] = work["PRODUCTO"].apply(normalize_product)
    work["AGENTE"] = work["CODIMEDI"].apply(normalize_agent)

    # Fecha para imputar el siniestro: fecha de ocurrencia.
    work["FECHA_SINIESTRO"] = pd.to_datetime(
        work["FECHOCUR"],
        dayfirst=True,
        errors="coerce",
    )
    work["ANIO_SINIESTRO"] = work["FECHA_SINIESTRO"].dt.year
    work["MES_SINIESTRO"] = work["FECHA_SINIESTRO"].dt.month

    # Importe para siniestralidad: reserva actual + pagos pendientes.
    # Ejemplo: 53,86 + 4.846,14 = 4.900,00
    # Ejemplo: 400,00 + 3.600,00 = 4.000,00
    work["RESERACT_VALOR"] = work["RESERACT"].apply(parse_spanish_number)
    work["PAGOSPDT_VALOR"] = work["PAGOSPDT"].apply(parse_spanish_number)

    # Columnas auxiliares para comprobar en detalle, no se usan en el importe.
    work["EXPECACT_VALOR"] = (
        work["EXPECACT"].apply(parse_spanish_number)
        if "EXPECACT" in work.columns
        else 0.0
    )
    work["PAGOSRZD_VALOR"] = (
        work["PAGOSRZD"].apply(parse_spanish_number)
        if "PAGOSRZD" in work.columns
        else 0.0
    )
    work["COSTESIN_VALOR"] = (
        work["COSTESIN"].apply(parse_spanish_number)
        if "COSTESIN" in work.columns
        else 0.0
    )

    work["IMPORTE_SINIESTRO"] = work["RESERACT_VALOR"] + work["PAGOSPDT_VALOR"]

    return work


def filter_siniestros(
    df: pd.DataFrame,
    ranking_date: date,
    excluded_products: list[str],
) -> pd.DataFrame:
    work = prepare_siniestros_data(df)
    excluded = {normalize_product(product) for product in excluded_products}

    mask = (
        work["FECHA_SINIESTRO"].notna()
        & work["ANIO_SINIESTRO"].eq(ranking_date.year)
        & work["MES_SINIESTRO"].le(ranking_date.month)
        & ~work["PRODUCTO_NORMALIZADO"].isin(excluded)
    )

    return work[mask].copy()


def aggregate_siniestros(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame(columns=["AGENTE", "IMPORTE_SINIESTROS", "NUM_SINIESTROS"])

    count_source = "NUMESINI" if "NUMESINI" in detail.columns else "IMPORTE_SINIESTRO"

    return (
        detail.groupby("AGENTE", dropna=False)
        .agg(
            IMPORTE_SINIESTROS=("IMPORTE_SINIESTRO", "sum"),
            NUM_SINIESTROS=(count_source, "count"),
        )
        .reset_index()
    )


def prepare_primas_data(df: pd.DataFrame, movement: str) -> pd.DataFrame:
    work = df.copy()

    work["MOVIMIENTO_PRIMA"] = movement
    work["AGENTE"] = work["MEDIADOR"].apply(normalize_agent)
    work["POLIPNET_VALOR"] = work["POLIPNET"].apply(parse_spanish_number)

    return work


def aggregate_primas(
    df: pd.DataFrame,
    movement: str,
    amount_column: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    detail = prepare_primas_data(df, movement)

    if detail.empty:
        return pd.DataFrame(columns=["AGENTE", amount_column]), detail

    aggregated = (
        detail.groupby("AGENTE", dropna=False)
        .agg(**{amount_column: ("POLIPNET_VALOR", "sum")})
        .reset_index()
    )

    return aggregated, detail


def calculate_ranking(
    facturacion_df: pd.DataFrame,
    anulaciones_df: pd.DataFrame,
    siniestros_df: pd.DataFrame,
    primas_emitidas_df: pd.DataFrame,
    primas_anuladas_df: pd.DataFrame,
    ranking_date: date,
    excluded_products: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
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

    siniestros_detail = filter_siniestros(
        siniestros_df,
        ranking_date,
        excluded_products,
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

    siniestros = aggregate_siniestros(siniestros_detail)

    primas_emitidas, primas_emitidas_detail = aggregate_primas(
        primas_emitidas_df,
        "EMITIDA",
        "PRIMAS_EMITIDAS",
    )

    primas_anuladas, primas_anuladas_detail = aggregate_primas(
        primas_anuladas_df,
        "ANULADA",
        "PRIMAS_ANULADAS",
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
                "PRIMAS_EMITIDAS",
                "PRIMAS_ANULADAS",
                "PRIMAS_NETAS",
                "IMPORTE_SINIESTROS",
                "SINIESTRALIDAD",
                "CUMPLE_SINIESTRALIDAD",
                "POLIZAS_ALTAS",
                "POLIZAS_ANULADAS",
                "POLIZAS_NETAS",
                "NUM_SINIESTROS",
                "PRIMA_MEDIA_NETA",
            ]
        )
        return (
            empty,
            altas_detail,
            anulaciones_detail,
            siniestros_detail,
            primas_emitidas_detail,
            primas_anuladas_detail,
        )

    ranking["NOMBRE_AGENCIA"] = ranking["NOMBRE_AGENCIA_ALTAS"].combine_first(
        ranking["NOMBRE_AGENCIA_ANULACIONES"]
    )

    ranking["NOMBRE_AGENCIA"] = [
        name if not pd.isna(name) and str(name).strip() != "" else agent
        for name, agent in zip(ranking["NOMBRE_AGENCIA"], ranking["AGENTE"])
    ]

    ranking = pd.merge(ranking, siniestros, on="AGENTE", how="left")
    ranking = pd.merge(ranking, primas_emitidas, on="AGENTE", how="left")
    ranking = pd.merge(ranking, primas_anuladas, on="AGENTE", how="left")

    numeric_columns = [
        "FACTURACION_ALTAS_BRUTAS",
        "FACTURACION_ANULACIONES",
        "POLIZAS_ALTAS",
        "POLIZAS_ANULADAS",
        "IMPORTE_SINIESTROS",
        "NUM_SINIESTROS",
        "PRIMAS_EMITIDAS",
        "PRIMAS_ANULADAS",
    ]

    for column in numeric_columns:
        if column in ranking.columns:
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

    ranking["PRIMAS_NETAS"] = ranking["PRIMAS_EMITIDAS"] - ranking["PRIMAS_ANULADAS"]

    ranking["SINIESTRALIDAD"] = [
        importe_siniestros / primas_netas if primas_netas > 0 else 0.0
        for importe_siniestros, primas_netas in zip(
            ranking["IMPORTE_SINIESTROS"],
            ranking["PRIMAS_NETAS"],
        )
    ]

    ranking["CUMPLE_SINIESTRALIDAD"] = ranking["SINIESTRALIDAD"] < SINIESTRALIDAD_MAXIMA

    ranking = ranking[
        [
            "AGENTE",
            "NOMBRE_AGENCIA",
            "DEPENDENCIA",
            "FACTURACION_ALTAS_BRUTAS",
            "FACTURACION_ANULACIONES",
            "FACTURACION_NETA",
            "PRIMAS_EMITIDAS",
            "PRIMAS_ANULADAS",
            "PRIMAS_NETAS",
            "IMPORTE_SINIESTROS",
            "SINIESTRALIDAD",
            "CUMPLE_SINIESTRALIDAD",
            "POLIZAS_ALTAS",
            "POLIZAS_ANULADAS",
            "POLIZAS_NETAS",
            "NUM_SINIESTROS",
            "PRIMA_MEDIA_NETA",
        ]
    ].sort_values(
        ["FACTURACION_NETA", "SINIESTRALIDAD", "AGENTE", "DEPENDENCIA"],
        ascending=[False, True, True, True],
    )

    ranking.insert(0, "RANKING", range(1, len(ranking) + 1))

    return (
        ranking,
        altas_detail,
        anulaciones_detail,
        siniestros_detail,
        primas_emitidas_detail,
        primas_anuladas_detail,
    )


# =========================
# SALUD
# =========================

def prepare_facturacion_salud_data(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    work["AGENTE"] = work["MEDIADOR"].apply(normalize_agent)
    work["POLIZA_NORMALIZADA"] = work["IDPOLIZA"].apply(lambda value: normalize_text(value, ""))
    work["FECHA_EFECTO_SALUD"] = pd.to_datetime(
        work["POLIEFEC"],
        dayfirst=True,
        errors="coerce",
    )
    work["ANIO_SALUD"] = work["FECHA_EFECTO_SALUD"].dt.year
    work["MES_SALUD"] = work["FECHA_EFECTO_SALUD"].dt.month
    work["PRIMA_NETA_SALUD_VALOR"] = work["PRIMA NETA"].apply(parse_spanish_number)

    return work


def prepare_bajas_salud_data(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    work["POLIZA_NORMALIZADA"] = work["POLIZA"].apply(lambda value: normalize_text(value, ""))
    work["DES_PRODUCTO_NORMALIZADO"] = work["DES_PRODUCTO"].apply(normalize_product)

    work["FECHA_BAJA_SALUD"] = pd.to_datetime(
        work["FEC_EFECTO_BAJA"],
        dayfirst=True,
        errors="coerce",
    )

    work["FECHA_REACTIVACION_SALUD"] = pd.to_datetime(
        work["FEC_EFECTO_REACTIV"],
        dayfirst=True,
        errors="coerce",
    )

    work["ANIO_BAJA_SALUD"] = work["FECHA_BAJA_SALUD"].dt.year
    work["MES_BAJA_SALUD"] = work["FECHA_BAJA_SALUD"].dt.month

    return work


def calculate_facturacion_salud(
    facturacion_salud_df: pd.DataFrame,
    bajas_salud_df: pd.DataFrame,
    ranking_date: date,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    facturacion_detail = prepare_facturacion_salud_data(facturacion_salud_df)
    bajas_detail = prepare_bajas_salud_data(bajas_salud_df)

    bruta_mask = (
        facturacion_detail["FECHA_EFECTO_SALUD"].notna()
        & facturacion_detail["ANIO_SALUD"].eq(ranking_date.year)
        & facturacion_detail["MES_SALUD"].le(ranking_date.month)
    )

    salud_bruta_detail = facturacion_detail[bruta_mask].copy()

    bajas_mask = (
        bajas_detail["FECHA_BAJA_SALUD"].notna()
        & bajas_detail["ANIO_BAJA_SALUD"].eq(ranking_date.year)
        & bajas_detail["MES_BAJA_SALUD"].le(ranking_date.month)
        & bajas_detail["FECHA_REACTIVACION_SALUD"].isna()
        & ~bajas_detail["DES_PRODUCTO_NORMALIZADO"].eq("ASISA VIDA RIESGO")
    )

    bajas_validas = bajas_detail[bajas_mask].copy()
    bajas_validas = bajas_validas.drop_duplicates("POLIZA_NORMALIZADA")

    salud_anulaciones_detail = pd.merge(
        facturacion_detail,
        bajas_validas,
        on="POLIZA_NORMALIZADA",
        how="inner",
        suffixes=("", "_BAJA"),
    )

    salud_bruta = (
        salud_bruta_detail.groupby("AGENTE", dropna=False)
        .agg(
            FACTURACION_SALUD_BRUTA=("PRIMA_NETA_SALUD_VALOR", "sum"),
            POLIZAS_SALUD_BRUTAS=("POLIZA_NORMALIZADA", "nunique"),
        )
        .reset_index()
    )

    salud_anulaciones = (
        salud_anulaciones_detail.groupby("AGENTE", dropna=False)
        .agg(
            FACTURACION_SALUD_ANULACIONES=("PRIMA_NETA_SALUD_VALOR", "sum"),
            POLIZAS_SALUD_ANULADAS=("POLIZA_NORMALIZADA", "nunique"),
        )
        .reset_index()
    )

    ranking_salud = pd.merge(
        salud_bruta,
        salud_anulaciones,
        on="AGENTE",
        how="outer",
    )

    if ranking_salud.empty:
        return (
            pd.DataFrame(
                columns=[
                    "RANKING_SALUD",
                    "AGENTE",
                    "FACTURACION_SALUD_BRUTA",
                    "FACTURACION_SALUD_ANULACIONES",
                    "FACTURACION_SALUD_NETA",
                    "POLIZAS_SALUD_BRUTAS",
                    "POLIZAS_SALUD_ANULADAS",
                    "POLIZAS_SALUD_NETAS",
                ]
            ),
            salud_bruta_detail,
            salud_anulaciones_detail,
        )

    for column in [
        "FACTURACION_SALUD_BRUTA",
        "FACTURACION_SALUD_ANULACIONES",
        "POLIZAS_SALUD_BRUTAS",
        "POLIZAS_SALUD_ANULADAS",
    ]:
        ranking_salud[column] = ranking_salud[column].fillna(0)

    ranking_salud["FACTURACION_SALUD_NETA"] = (
        ranking_salud["FACTURACION_SALUD_BRUTA"]
        - ranking_salud["FACTURACION_SALUD_ANULACIONES"]
    )

    ranking_salud["POLIZAS_SALUD_NETAS"] = (
        ranking_salud["POLIZAS_SALUD_BRUTAS"]
        - ranking_salud["POLIZAS_SALUD_ANULADAS"]
    )

    ranking_salud = ranking_salud.sort_values(
        ["FACTURACION_SALUD_NETA", "AGENTE"],
        ascending=[False, True],
    )

    ranking_salud.insert(0, "RANKING_SALUD", range(1, len(ranking_salud) + 1))

    return ranking_salud, salud_bruta_detail, salud_anulaciones_detail


# =========================
# VIDA
# =========================

def prepare_facturacion_vida_data(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    work["AGENTE"] = work["CODIMEDI"].apply(normalize_agent)
    work["POLIZA_NORMALIZADA"] = work["NUMERO"].apply(lambda value: normalize_text(value, ""))
    work["FECHA_ALTA_VIDA"] = pd.to_datetime(
        work["FECHALTA"],
        dayfirst=True,
        errors="coerce",
    )
    work["ANIO_VIDA"] = work["FECHA_ALTA_VIDA"].dt.year
    work["MES_VIDA"] = work["FECHA_ALTA_VIDA"].dt.month
    work["PRIMA_VIDA_VALOR"] = work["PRIMATOTAL"].apply(parse_spanish_number)

    return work


def calculate_facturacion_vida(
    facturacion_vida_df: pd.DataFrame,
    bajas_salud_df: pd.DataFrame,
    ranking_date: date,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    facturacion_detail = prepare_facturacion_vida_data(facturacion_vida_df)
    bajas_detail = prepare_bajas_salud_data(bajas_salud_df)

    bruta_mask = (
        facturacion_detail["FECHA_ALTA_VIDA"].notna()
        & facturacion_detail["ANIO_VIDA"].eq(ranking_date.year)
        & facturacion_detail["MES_VIDA"].le(ranking_date.month)
    )

    vida_bruta_detail = facturacion_detail[bruta_mask].copy()

    bajas_mask = (
        bajas_detail["FECHA_BAJA_SALUD"].notna()
        & bajas_detail["ANIO_BAJA_SALUD"].eq(ranking_date.year)
        & bajas_detail["MES_BAJA_SALUD"].le(ranking_date.month)
        & bajas_detail["FECHA_REACTIVACION_SALUD"].isna()
        & bajas_detail["DES_PRODUCTO_NORMALIZADO"].eq("ASISA VIDA RIESGO")
    )

    bajas_validas = bajas_detail[bajas_mask].copy()
    bajas_validas = bajas_validas.drop_duplicates("POLIZA_NORMALIZADA")

    vida_anulaciones_detail = pd.merge(
        facturacion_detail,
        bajas_validas,
        on="POLIZA_NORMALIZADA",
        how="inner",
        suffixes=("", "_BAJA"),
    )

    vida_bruta = (
        vida_bruta_detail.groupby("AGENTE", dropna=False)
        .agg(
            FACTURACION_VIDA_BRUTA=("PRIMA_VIDA_VALOR", "sum"),
            POLIZAS_VIDA_BRUTAS=("POLIZA_NORMALIZADA", "nunique"),
        )
        .reset_index()
    )

    vida_anulaciones = (
        vida_anulaciones_detail.groupby("AGENTE", dropna=False)
        .agg(
            FACTURACION_VIDA_ANULACIONES=("PRIMA_VIDA_VALOR", "sum"),
            POLIZAS_VIDA_ANULADAS=("POLIZA_NORMALIZADA", "nunique"),
        )
        .reset_index()
    )

    ranking_vida = pd.merge(
        vida_bruta,
        vida_anulaciones,
        on="AGENTE",
        how="outer",
    )

    if ranking_vida.empty:
        return (
            pd.DataFrame(
                columns=[
                    "RANKING_VIDA",
                    "AGENTE",
                    "FACTURACION_VIDA_BRUTA",
                    "FACTURACION_VIDA_ANULACIONES",
                    "FACTURACION_VIDA_NETA",
                    "POLIZAS_VIDA_BRUTAS",
                    "POLIZAS_VIDA_ANULADAS",
                    "POLIZAS_VIDA_NETAS",
                ]
            ),
            vida_bruta_detail,
            vida_anulaciones_detail,
        )

    for column in [
        "FACTURACION_VIDA_BRUTA",
        "FACTURACION_VIDA_ANULACIONES",
        "POLIZAS_VIDA_BRUTAS",
        "POLIZAS_VIDA_ANULADAS",
    ]:
        ranking_vida[column] = ranking_vida[column].fillna(0)

    ranking_vida["FACTURACION_VIDA_NETA"] = (
        ranking_vida["FACTURACION_VIDA_BRUTA"]
        - ranking_vida["FACTURACION_VIDA_ANULACIONES"]
    )

    ranking_vida["POLIZAS_VIDA_NETAS"] = (
        ranking_vida["POLIZAS_VIDA_BRUTAS"]
        - ranking_vida["POLIZAS_VIDA_ANULADAS"]
    )

    ranking_vida = ranking_vida.sort_values(
        ["FACTURACION_VIDA_NETA", "AGENTE"],
        ascending=[False, True],
    )

    ranking_vida.insert(0, "RANKING_VIDA", range(1, len(ranking_vida) + 1))

    return ranking_vida, vida_bruta_detail, vida_anulaciones_detail


# =========================
# LIGAS
# =========================

def classify_liga_decesos(facturacion_decesos: float, cumple_siniestralidad: bool) -> str:
    if not cumple_siniestralidad:
        return "No cumple siniestralidad"

    if facturacion_decesos >= DECESOS_LIGA_ELITE_MINIMA:
        return "LIGA ELITE"

    if facturacion_decesos >= DECESOS_LIGA_PRO_MINIMA:
        return "LIGA PRO"

    return "No clasifica"


def classify_liga_salud(facturacion_salud: float, facturacion_decesos: float) -> str:
    if (
        facturacion_salud >= SALUD_LIGA_ELITE_MINIMA
        and facturacion_decesos >= SALUD_LIGA_ELITE_DECESOS_MINIMA
    ):
        return "LIGA ELITE"

    if (
        facturacion_salud >= SALUD_LIGA_PRO_MINIMA
        and facturacion_decesos >= SALUD_LIGA_PRO_DECESOS_MINIMA
    ):
        return "LIGA PRO"

    return "No clasifica"


def build_ranking_decesos_top10(ranking: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "PUESTO_DECESOS",
        "LIGA_DECESOS",
        "CODIMEDI",
        "NOMBRE_AGENCIA",
        "RESPONSABLE",
        "FACTURACION_NETA",
        "SINIESTRALIDAD",
        "CUMPLE_SINIESTRALIDAD",
        "IMPORTE_SINIESTROS",
        "PRIMAS_NETAS",
    ]

    if ranking.empty:
        return pd.DataFrame(columns=columns)

    result = ranking.copy()
    result["LIGA_DECESOS"] = [
        classify_liga_decesos(facturacion, cumple)
        for facturacion, cumple in zip(result["FACTURACION_NETA"], result["CUMPLE_SINIESTRALIDAD"])
    ]

    result = result.sort_values(
        ["FACTURACION_NETA", "SINIESTRALIDAD", "CODIMEDI"],
        ascending=[False, True, True],
    ).head(10)

    result.insert(0, "PUESTO_DECESOS", range(1, len(result) + 1))

    return result[columns]


def build_ranking_salud_top10(ranking_salud: pd.DataFrame, ranking_decesos: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "PUESTO_SALUD",
        "LIGA_SALUD",
        "CODIMEDI",
        "NOMBRE_AGENCIA",
        "RESPONSABLE",
        "FACTURACION_SALUD_NETA",
        "FACTURACION_DECESOS_NETA",
    ]

    if ranking_salud.empty:
        return pd.DataFrame(columns=columns)

    salud_columns = ["AGENTE", "CODIMEDI", "NOMBRE_AGENCIA", "RESPONSABLE", "FACTURACION_SALUD_NETA"]
    salud = ranking_salud[[column for column in salud_columns if column in ranking_salud.columns]].copy()

    if "CODIMEDI" not in salud.columns:
        salud["CODIMEDI"] = salud["AGENTE"]
    if "NOMBRE_AGENCIA" not in salud.columns:
        salud["NOMBRE_AGENCIA"] = salud["AGENTE"]
    if "RESPONSABLE" not in salud.columns:
        salud["RESPONSABLE"] = "Sin responsable"

    if ranking_decesos.empty:
        decesos = pd.DataFrame(columns=["AGENTE", "FACTURACION_DECESOS_NETA"])
    else:
        decesos = ranking_decesos[["AGENTE", "FACTURACION_NETA"]].rename(
            columns={"FACTURACION_NETA": "FACTURACION_DECESOS_NETA"}
        )

    result = pd.merge(salud, decesos, on="AGENTE", how="left")
    result["FACTURACION_DECESOS_NETA"] = result["FACTURACION_DECESOS_NETA"].fillna(0)

    result["LIGA_SALUD"] = [
        classify_liga_salud(facturacion_salud, facturacion_decesos)
        for facturacion_salud, facturacion_decesos in zip(
            result["FACTURACION_SALUD_NETA"],
            result["FACTURACION_DECESOS_NETA"],
        )
    ]

    result = result.sort_values(
        ["FACTURACION_SALUD_NETA", "FACTURACION_DECESOS_NETA", "CODIMEDI"],
        ascending=[False, False, True],
    ).head(10)

    result.insert(0, "PUESTO_SALUD", range(1, len(result) + 1))

    return result[columns]


def build_ranking_vida_top10(ranking_vida: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "PUESTO_VIDA",
        "CODIMEDI",
        "NOMBRE_AGENCIA",
        "RESPONSABLE",
        "FACTURACION_VIDA_NETA",
        "FACTURACION_VIDA_BRUTA",
        "FACTURACION_VIDA_ANULACIONES",
        "POLIZAS_VIDA_NETAS",
    ]

    if ranking_vida.empty:
        return pd.DataFrame(columns=columns)

    result = ranking_vida.copy()

    if "CODIMEDI" not in result.columns:
        result["CODIMEDI"] = result["AGENTE"]
    if "NOMBRE_AGENCIA" not in result.columns:
        result["NOMBRE_AGENCIA"] = result["AGENTE"]
    if "RESPONSABLE" not in result.columns:
        result["RESPONSABLE"] = "Sin responsable"

    result = result.sort_values(
        ["FACTURACION_VIDA_NETA", "CODIMEDI"],
        ascending=[False, True],
    ).head(10)

    result.insert(0, "PUESTO_VIDA", range(1, len(result) + 1))

    return result[columns]


# =========================
# OUTPUT
# =========================

def build_sheet_summary(df: pd.DataFrame, file_name: str) -> pd.DataFrame:
    if "HOJA_ORIGEN" not in df.columns:
        return pd.DataFrame(columns=["TIPO_ARCHIVO", "ARCHIVO", "HOJA", "FILAS_LEIDAS"])

    group_columns = ["HOJA_ORIGEN"]
    if "ARCHIVO_ORIGEN" in df.columns:
        group_columns = ["ARCHIVO_ORIGEN", "HOJA_ORIGEN"]

    summary = (
        df.groupby(group_columns, dropna=False)
        .size()
        .reset_index(name="FILAS_LEIDAS")
        .rename(columns={"ARCHIVO_ORIGEN": "ARCHIVO", "HOJA_ORIGEN": "HOJA"})
    )

    if "ARCHIVO" not in summary.columns:
        summary.insert(0, "ARCHIVO", file_name)

    summary.insert(0, "TIPO_ARCHIVO", file_name)

    return summary


def format_euro(value: float) -> str:
    return f"{value:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def format_percent(value: float) -> str:
    return f"{value:.2%}".replace(".", ",")


def style_liga_decesos(df: pd.DataFrame):
    return df.style.format(
        {
            "FACTURACION_NETA": lambda value: format_euro(float(value)),
            "SINIESTRALIDAD": lambda value: format_percent(float(value)),
            "IMPORTE_SINIESTROS": lambda value: format_euro(float(value)),
            "PRIMAS_NETAS": lambda value: format_euro(float(value)),
        }
    ).set_properties(
        subset=["LIGA_DECESOS"],
        **{"background-color": COLOR_DECESOS, "color": "white", "font-weight": "bold"},
    )


def style_liga_salud(df: pd.DataFrame):
    return df.style.format(
        {
            "FACTURACION_SALUD_NETA": lambda value: format_euro(float(value)),
            "FACTURACION_DECESOS_NETA": lambda value: format_euro(float(value)),
        }
    ).set_properties(
        subset=["LIGA_SALUD"],
        **{"background-color": COLOR_SALUD, "color": "white", "font-weight": "bold"},
    )


def style_ranking_vida(df: pd.DataFrame):
    return df.style.format(
        {
            "FACTURACION_VIDA_NETA": lambda value: format_euro(float(value)),
            "FACTURACION_VIDA_BRUTA": lambda value: format_euro(float(value)),
            "FACTURACION_VIDA_ANULACIONES": lambda value: format_euro(float(value)),
        }
    ).set_properties(
        subset=["PUESTO_VIDA"],
        **{"background-color": COLOR_VIDA, "color": "black", "font-weight": "bold"},
    )


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


def siniestros_columns_for_display(detail: pd.DataFrame) -> list[str]:
    columns = [
        "HOJA_ORIGEN",
        "PRODUCTO",
        "CODIMEDI",
        "AGENTE",
        "FECHOCUR",
        "FECHA_SINIESTRO",
        "ANIO_SINIESTRO",
        "MES_SINIESTRO",
        "NUMESINI",
        "POLIZSEC",
        "ESTADO",
        "MOTIVO",
        "COBERTURA",
        "NATURALEZA",
        "RESERACT",
        "RESERACT_VALOR",
        "EXPECACT",
        "EXPECACT_VALOR",
        "PAGOSPDT",
        "PAGOSPDT_VALOR",
        "PAGOSRZD",
        "PAGOSRZD_VALOR",
        "COSTESIN",
        "COSTESIN_VALOR",
        "IMPORTE_SINIESTRO",
    ]

    return [column for column in columns if column in detail.columns]


def primas_columns_for_display(detail: pd.DataFrame) -> list[str]:
    columns = [
        "ARCHIVO_ORIGEN",
        "HOJA_ORIGEN",
        "MOVIMIENTO_PRIMA",
        "MEDIADOR",
        "AGENTE",
        "GARANTIA",
        "POLIPTOT",
        "POLIPNET",
        "POLIPNET_VALOR",
        "CONSORCIO",
        "CLEA",
        "IPS",
        "RECARGO",
    ]

    return [column for column in columns if column in detail.columns]


def salud_columns_for_display(detail: pd.DataFrame) -> list[str]:
    columns = [
        "ARCHIVO_ORIGEN",
        "HOJA_ORIGEN",
        "PRODUCTO",
        "IDPOLIZA",
        "POLIZA_NORMALIZADA",
        "MEDIADOR",
        "AGENTE",
        "POLIEFEC",
        "FECHA_EFECTO_SALUD",
        "PRIMA NETA",
        "PRIMA_NETA_SALUD_VALOR",
        "DES_PRODUCTO",
        "FEC_EFECTO_BAJA",
        "FEC_EFECTO_REACTIV",
        "MOTIVO_BAJA",
        "TOMADOR",
        "NIF",
    ]

    return [column for column in columns if column in detail.columns]


def vida_columns_for_display(detail: pd.DataFrame) -> list[str]:
    columns = [
        "ARCHIVO_ORIGEN",
        "HOJA_ORIGEN",
        "NUMERO",
        "POLIZA_NORMALIZADA",
        "CODIMEDI",
        "AGENTE",
        "ESTADO",
        "FECHALTA",
        "FECHA_ALTA_VIDA",
        "GARANTIA",
        "PRIMATOTAL",
        "PRIMA_VIDA_VALOR",
        "NOMBRE",
        "APE1",
        "APE2",
        "NUMEDOCU",
        "DES_PRODUCTO",
        "FEC_EFECTO_BAJA",
        "FEC_EFECTO_REACTIV",
        "MOTIVO_BAJA",
    ]

    return [column for column in columns if column in detail.columns]


def dataframe_to_excel(
    ranking: pd.DataFrame,
    ranking_salud: pd.DataFrame,
    ranking_vida: pd.DataFrame,
    ranking_decesos_top10: pd.DataFrame,
    ranking_salud_top10: pd.DataFrame,
    ranking_vida_top10: pd.DataFrame,
    altas_detail: pd.DataFrame,
    anulaciones_detail: pd.DataFrame,
    siniestros_detail: pd.DataFrame,
    primas_emitidas_detail: pd.DataFrame,
    primas_anuladas_detail: pd.DataFrame,
    salud_bruta_detail: pd.DataFrame,
    salud_anulaciones_detail: pd.DataFrame,
    vida_bruta_detail: pd.DataFrame,
    vida_anulaciones_detail: pd.DataFrame,
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
            {"CAMPO": "REGLA_ANULACIONES_DECESOS", "VALOR": "FECHA EMISION del 1 al 25 de enero se imputa al año anterior"},
            {"CAMPO": "SINIESTROS_FECHA", "VALOR": "FECHOCUR"},
            {"CAMPO": "SINIESTROS_IMPORTE", "VALOR": "RESERACT + PAGOSPDT"},
            {"CAMPO": "DECESOS_LIGA_PRO", "VALOR": "Facturacion neta >= 30.000 y siniestralidad < 25%"},
            {"CAMPO": "DECESOS_LIGA_ELITE", "VALOR": "Facturacion neta >= 60.000 y siniestralidad < 25%"},
            {"CAMPO": "SALUD_LIGA_PRO", "VALOR": "Facturacion Salud >= 25.000 y Facturacion Decesos >= 12.000"},
            {"CAMPO": "SALUD_LIGA_ELITE", "VALOR": "Facturacion Salud >= 80.000 y Facturacion Decesos >= 4.000"},
            {"CAMPO": "SALUD_BRUTA", "VALOR": "FACTURACION_SALUD con POLIEFEC dentro del periodo"},
            {"CAMPO": "SALUD_ANULACIONES", "VALOR": "INFORME_BAJAS_SALUD con FEC_EFECTO_BAJA dentro del periodo, sin FEC_EFECTO_REACTIV y excluyendo ASISA VIDA RIESGO"},
            {"CAMPO": "VIDA_BRUTA", "VALOR": "FACTURACION_VIDA con FECHALTA dentro del periodo y PRIMATOTAL como importe"},
            {"CAMPO": "VIDA_ANULACIONES", "VALOR": "INFORME_BAJAS_SALUD filtrando DES_PRODUCTO = ASISA VIDA RIESGO"},
            {"CAMPO": "TOP_RANKINGS", "VALOR": "Se muestran hasta 10 puestos por facturacion neta aunque no entren en liga"},
        ]
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        parametros.to_excel(writer, index=False, sheet_name="PARAMETROS")
        sheet_summary.to_excel(writer, index=False, sheet_name="COMPROBACION_HOJAS")
        ranking_decesos_top10.to_excel(writer, index=False, sheet_name="RANKING_DECESOS")
        ranking_salud_top10.to_excel(writer, index=False, sheet_name="RANKING_SALUD")
        ranking_vida_top10.to_excel(writer, index=False, sheet_name="RANKING_VIDA")
        ranking.to_excel(writer, index=False, sheet_name="RANKING_NETO_DECESOS")
        ranking_salud.to_excel(writer, index=False, sheet_name="FACTURACION_SALUD")
        ranking_vida.to_excel(writer, index=False, sheet_name="FACTURACION_VIDA")
        altas_detail[detail_columns_for_display(altas_detail)].to_excel(writer, index=False, sheet_name="DETALLE_ALTAS_DECESOS")
        anulaciones_detail[detail_columns_for_display(anulaciones_detail)].to_excel(writer, index=False, sheet_name="DETALLE_ANUL_DECESOS")
        siniestros_detail[siniestros_columns_for_display(siniestros_detail)].to_excel(writer, index=False, sheet_name="DETALLE_SINIESTROS")
        primas_emitidas_detail[primas_columns_for_display(primas_emitidas_detail)].to_excel(writer, index=False, sheet_name="DETALLE_PRIMAS_EMIT")
        primas_anuladas_detail[primas_columns_for_display(primas_anuladas_detail)].to_excel(writer, index=False, sheet_name="DETALLE_PRIMAS_ANUL")
        salud_bruta_detail[salud_columns_for_display(salud_bruta_detail)].to_excel(writer, index=False, sheet_name="DETALLE_SALUD_BRUTA")
        salud_anulaciones_detail[salud_columns_for_display(salud_anulaciones_detail)].to_excel(writer, index=False, sheet_name="DETALLE_SALUD_ANUL")
        vida_bruta_detail[vida_columns_for_display(vida_bruta_detail)].to_excel(writer, index=False, sheet_name="DETALLE_VIDA_BRUTA")
        vida_anulaciones_detail[vida_columns_for_display(vida_anulaciones_detail)].to_excel(writer, index=False, sheet_name="DETALLE_VIDA_ANUL")

    return output.getvalue()


# =========================
# APP
# =========================

def main() -> None:
    st.set_page_config(page_title="Ranking agentes", layout="wide")

    st.title("Ranking agentes - DECESOS + SALUD + VIDA")

    ranking_date = st.date_input(
        "Para que fecha quieres calcular el ranking?",
        value=DEFAULT_RANKING_DATE,
        format="DD/MM/YYYY",
    )

    st.header("Archivos Decesos")

    col_upload_1, col_upload_2, col_upload_3 = st.columns(3)
    col_upload_4, col_upload_5 = st.columns(2)

    with col_upload_1:
        uploaded_facturacion = st.file_uploader("Sube FACTURACION_DECESOS.xls", type=["xls", "xlsx"], key="facturacion")

    with col_upload_2:
        uploaded_anulaciones = st.file_uploader("Sube FACTURACION_ANULACIONES_DECESOS.xls", type=["xls", "xlsx"], key="anulaciones")

    with col_upload_3:
        uploaded_siniestros = st.file_uploader("Sube SINIESTROS_DECESOS.xls", type=["xls", "xlsx"], key="siniestros")

    with col_upload_4:
        uploaded_primas_emitidas = st.file_uploader(
            "Sube PRIMAS_EMITIDAS mensuales (.xls/.xlsx)",
            type=["xls", "xlsx"],
            key="primas_emitidas",
            accept_multiple_files=True,
        )

    with col_upload_5:
        uploaded_primas_anuladas = st.file_uploader(
            "Sube PRIMAS_ANULADAS mensuales (.xls/.xlsx)",
            type=["xls", "xlsx"],
            key="primas_anuladas",
            accept_multiple_files=True,
        )

    st.header("Archivos Salud, Vida y Mapeo")

    col_upload_6, col_upload_7, col_upload_8, col_upload_9 = st.columns(4)

    with col_upload_6:
        uploaded_facturacion_salud = st.file_uploader(
            "Sube FACTURACION_SALUD.xls",
            type=["xls", "xlsx"],
            key="facturacion_salud",
        )

    with col_upload_7:
        uploaded_bajas_salud = st.file_uploader(
            "Sube INFORME_BAJAS_SALUD.xlsx",
            type=["xls", "xlsx"],
            key="bajas_salud",
        )

    with col_upload_8:
        uploaded_facturacion_vida = st.file_uploader(
            "Sube FACTURACION_VIDA.xls",
            type=["xls", "xlsx"],
            key="facturacion_vida",
        )

    with col_upload_9:
        uploaded_mapeo = st.file_uploader(
            "Sube MAPEO_MEDIADORES.xlsx",
            type=["xls", "xlsx"],
            key="mapeo_mediadores",
        )

    if (
        uploaded_facturacion is None
        or uploaded_anulaciones is None
        or uploaded_siniestros is None
        or not uploaded_primas_emitidas
        or not uploaded_primas_anuladas
        or uploaded_facturacion_salud is None
        or uploaded_bajas_salud is None
        or uploaded_facturacion_vida is None
        or uploaded_mapeo is None
    ):
        st.info("Sube todos los archivos para calcular Decesos, Salud, Vida, mapeo y rankings.")
        st.stop()

    try:
        raw_facturacion_df = read_excel_all_sheets(uploaded_facturacion)
        raw_anulaciones_df = read_excel_all_sheets(uploaded_anulaciones)
        raw_siniestros_df = read_excel_all_sheets(uploaded_siniestros)
        raw_primas_emitidas_df = read_excel_many_files(uploaded_primas_emitidas)
        raw_primas_anuladas_df = read_excel_many_files(uploaded_primas_anuladas)
        raw_facturacion_salud_df = read_excel_all_sheets(uploaded_facturacion_salud)
        raw_bajas_salud_df = read_excel_all_sheets(uploaded_bajas_salud)
        raw_facturacion_vida_df = read_excel_all_sheets(uploaded_facturacion_vida)
        raw_mapeo_df = read_excel_all_sheets(uploaded_mapeo)
    except Exception as error:
        st.error(f"No he podido leer los archivos: {error}")
        st.stop()

    if (
        raw_facturacion_df.empty
        or raw_anulaciones_df.empty
        or raw_siniestros_df.empty
        or raw_primas_emitidas_df.empty
        or raw_primas_anuladas_df.empty
        or raw_facturacion_salud_df.empty
        or raw_bajas_salud_df.empty
        or raw_facturacion_vida_df.empty
        or raw_mapeo_df.empty
    ):
        st.warning("Alguno de los archivos esta vacio.")
        st.stop()

    missing_facturacion = validate_columns(raw_facturacion_df, REQUIRED_FACTURACION_COLUMNS)
    missing_anulaciones = validate_columns(raw_anulaciones_df, REQUIRED_ANULACIONES_COLUMNS)
    missing_siniestros = validate_columns(raw_siniestros_df, REQUIRED_SINIESTROS_COLUMNS)
    missing_primas_emitidas = validate_columns(raw_primas_emitidas_df, REQUIRED_PRIMAS_COLUMNS)
    missing_primas_anuladas = validate_columns(raw_primas_anuladas_df, REQUIRED_PRIMAS_COLUMNS)
    missing_facturacion_salud = validate_columns(raw_facturacion_salud_df, REQUIRED_FACTURACION_SALUD_COLUMNS)
    missing_bajas_salud = validate_columns(raw_bajas_salud_df, REQUIRED_BAJAS_SALUD_COLUMNS)
    missing_facturacion_vida = validate_columns(raw_facturacion_vida_df, REQUIRED_FACTURACION_VIDA_COLUMNS)
    missing_mapeo = validate_columns(raw_mapeo_df, REQUIRED_MAPEO_COLUMNS)

    if (
        missing_facturacion
        or missing_anulaciones
        or missing_siniestros
        or missing_primas_emitidas
        or missing_primas_anuladas
        or missing_facturacion_salud
        or missing_bajas_salud
        or missing_facturacion_vida
        or missing_mapeo
    ):
        messages = []
        if missing_facturacion:
            messages.append(f"Facturacion Decesos: {', '.join(missing_facturacion)}")
        if missing_anulaciones:
            messages.append(f"Anulaciones Decesos: {', '.join(missing_anulaciones)}")
        if missing_siniestros:
            messages.append(f"Siniestros Decesos: {', '.join(missing_siniestros)}")
        if missing_primas_emitidas:
            messages.append(f"Primas emitidas: {', '.join(missing_primas_emitidas)}")
        if missing_primas_anuladas:
            messages.append(f"Primas anuladas: {', '.join(missing_primas_anuladas)}")
        if missing_facturacion_salud:
            messages.append(f"Facturacion Salud: {', '.join(missing_facturacion_salud)}")
        if missing_bajas_salud:
            messages.append(f"Bajas Salud: {', '.join(missing_bajas_salud)}")
        if missing_facturacion_vida:
            messages.append(f"Facturacion Vida: {', '.join(missing_facturacion_vida)}")
        if missing_mapeo:
            messages.append(f"Mapeo mediadores: {', '.join(missing_mapeo)}")

        st.error("Faltan columnas obligatorias. " + " | ".join(messages))
        st.stop()

    mapeo = prepare_mapeo_data(raw_mapeo_df)

    sheet_summary = pd.concat(
        [
            build_sheet_summary(raw_facturacion_df, "FACTURACION_DECESOS"),
            build_sheet_summary(raw_anulaciones_df, "FACTURACION_ANULACIONES_DECESOS"),
            build_sheet_summary(raw_siniestros_df, "SINIESTROS_DECESOS"),
            build_sheet_summary(raw_primas_emitidas_df, "PRIMAS_EMITIDAS"),
            build_sheet_summary(raw_primas_anuladas_df, "PRIMAS_ANULADAS"),
            build_sheet_summary(raw_facturacion_salud_df, "FACTURACION_SALUD"),
            build_sheet_summary(raw_bajas_salud_df, "BAJAS_SALUD"),
            build_sheet_summary(raw_facturacion_vida_df, "FACTURACION_VIDA"),
            build_sheet_summary(raw_mapeo_df, "MAPEO_MEDIADORES"),
        ],
        ignore_index=True,
    )

    with st.expander("Comprobacion de hojas leidas", expanded=True):
        st.dataframe(sheet_summary, use_container_width=True, hide_index=True)

    product_options = sorted(
        pd.concat(
            [raw_facturacion_df["PRODUCTO"], raw_anulaciones_df["PRODUCTO"], raw_siniestros_df["PRODUCTO"]],
            ignore_index=True,
        )
        .dropna()
        .map(normalize_product)
        .unique()
        .tolist()
    )

    default_excluded = [product for product in DEFAULT_EXCLUDED_PRODUCTS if product in product_options]

    excluded_products = st.multiselect(
        "Productos excluidos Decesos",
        product_options,
        default=default_excluded,
    )

    (
        ranking,
        altas_detail,
        anulaciones_detail,
        siniestros_detail,
        primas_emitidas_detail,
        primas_anuladas_detail,
    ) = calculate_ranking(
        raw_facturacion_df,
        raw_anulaciones_df,
        raw_siniestros_df,
        raw_primas_emitidas_df,
        raw_primas_anuladas_df,
        ranking_date,
        excluded_products,
    )

    ranking = add_mapeo_to_ranking(ranking, mapeo)

    ranking_salud, salud_bruta_detail, salud_anulaciones_detail = calculate_facturacion_salud(
        raw_facturacion_salud_df,
        raw_bajas_salud_df,
        ranking_date,
    )
    ranking_salud = add_mapeo_to_simple_ranking(ranking_salud, mapeo)

    ranking_vida, vida_bruta_detail, vida_anulaciones_detail = calculate_facturacion_vida(
        raw_facturacion_vida_df,
        raw_bajas_salud_df,
        ranking_date,
    )
    ranking_vida = add_mapeo_to_simple_ranking(ranking_vida, mapeo)

    ranking_decesos_top10 = build_ranking_decesos_top10(ranking)
    ranking_salud_top10 = build_ranking_salud_top10(ranking_salud, ranking)
    ranking_vida_top10 = build_ranking_vida_top10(ranking_vida)

    total_neta = float(ranking["FACTURACION_NETA"].sum()) if not ranking.empty else 0.0
    total_primas_netas = float(ranking["PRIMAS_NETAS"].sum()) if not ranking.empty else 0.0
    total_siniestros = float(ranking["IMPORTE_SINIESTROS"].sum()) if not ranking.empty else 0.0
    siniestralidad_total = total_siniestros / total_primas_netas if total_primas_netas > 0 else 0.0

    total_salud_bruta = float(ranking_salud["FACTURACION_SALUD_BRUTA"].sum()) if not ranking_salud.empty else 0.0
    total_salud_anulaciones = float(ranking_salud["FACTURACION_SALUD_ANULACIONES"].sum()) if not ranking_salud.empty else 0.0
    total_salud_neta = float(ranking_salud["FACTURACION_SALUD_NETA"].sum()) if not ranking_salud.empty else 0.0
    total_vida_bruta = float(ranking_vida["FACTURACION_VIDA_BRUTA"].sum()) if not ranking_vida.empty else 0.0
    total_vida_anulaciones = float(ranking_vida["FACTURACION_VIDA_ANULACIONES"].sum()) if not ranking_vida.empty else 0.0
    total_vida_neta = float(ranking_vida["FACTURACION_VIDA_NETA"].sum()) if not ranking_vida.empty else 0.0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Facturacion neta Decesos", format_euro(total_neta))
    col2.metric("Siniestralidad Decesos", format_percent(siniestralidad_total))
    col3.metric("Facturacion neta Salud", format_euro(total_salud_neta))
    col4.metric("Facturacion neta Vida", format_euro(total_vida_neta))

    col5, col6, col7, col8 = st.columns(4)
    col5.metric("Facturacion Salud bruta", format_euro(total_salud_bruta))
    col6.metric("Anulaciones Salud", format_euro(total_salud_anulaciones))
    col7.metric("Facturacion Vida bruta", format_euro(total_vida_bruta))
    col8.metric("Anulaciones Vida", format_euro(total_vida_anulaciones))

    st.sidebar.title("Rankings")
    vista = st.sidebar.radio(
        "Selecciona vista",
        [
            "Ranking Decesos",
            "Ranking Salud",
            "Ranking Vida",
            "Facturacion Salud",
            "Facturacion Vida",
            "Ranking completo Decesos",
            "Detalles",
        ],
    )

    if vista == "Ranking Decesos":
        st.markdown(f"<h2 style='color:{COLOR_DECESOS};'>Ranking Decesos</h2>", unsafe_allow_html=True)
        st.caption("Top 10 por facturacion neta. LIGA PRO: desde 30.000 €. LIGA ELITE: desde 60.000 €. En ambos casos debe cumplir siniestralidad inferior al 25%.")

        if ranking_decesos_top10.empty:
            st.info("No hay datos en Ranking Decesos.")
        else:
            st.dataframe(style_liga_decesos(ranking_decesos_top10), use_container_width=True, hide_index=True)

    elif vista == "Ranking Salud":
        st.markdown(f"<h2 style='color:{COLOR_SALUD};'>Ranking Salud</h2>", unsafe_allow_html=True)
        st.caption("Top 10 por facturacion neta Salud. LIGA PRO: Salud >= 25.000 € y Decesos >= 12.000 €. LIGA ELITE: Salud >= 80.000 € y Decesos >= 4.000 €.")

        if ranking_salud_top10.empty:
            st.info("No hay datos en Ranking Salud.")
        else:
            st.dataframe(style_liga_salud(ranking_salud_top10), use_container_width=True, hide_index=True)

    elif vista == "Ranking Vida":
        st.markdown(f"<h2 style='color:{COLOR_VIDA};'>Ranking Vida</h2>", unsafe_allow_html=True)
        st.caption("Top 10 por facturacion neta Vida. La facturacion bruta usa FECHALTA y PRIMATOTAL; las anulaciones salen de bajas con ASISA VIDA RIESGO.")

        if ranking_vida_top10.empty:
            st.info("No hay datos en Ranking Vida.")
        else:
            st.dataframe(style_ranking_vida(ranking_vida_top10), use_container_width=True, hide_index=True)

    elif vista == "Facturacion Salud":
        st.subheader("Facturacion Salud")
        if ranking_salud.empty:
            st.info("No hay datos de Salud para los filtros seleccionados.")
        else:
            st.dataframe(
                ranking_salud.style.format(
                    {
                        "FACTURACION_SALUD_BRUTA": lambda value: format_euro(float(value)),
                        "FACTURACION_SALUD_ANULACIONES": lambda value: format_euro(float(value)),
                        "FACTURACION_SALUD_NETA": lambda value: format_euro(float(value)),
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    elif vista == "Facturacion Vida":
        st.subheader("Facturacion Vida")
        if ranking_vida.empty:
            st.info("No hay datos de Vida para los filtros seleccionados.")
        else:
            st.dataframe(
                ranking_vida.style.format(
                    {
                        "FACTURACION_VIDA_BRUTA": lambda value: format_euro(float(value)),
                        "FACTURACION_VIDA_ANULACIONES": lambda value: format_euro(float(value)),
                        "FACTURACION_VIDA_NETA": lambda value: format_euro(float(value)),
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    elif vista == "Ranking completo Decesos":
        st.subheader("Ranking facturacion neta Decesos con primas netas y siniestralidad")
        if ranking.empty:
            st.info("No hay datos para los filtros seleccionados.")
        else:
            st.dataframe(
                ranking.style.format(
                    {
                        "FACTURACION_ALTAS_BRUTAS": lambda value: format_euro(float(value)),
                        "FACTURACION_ANULACIONES": lambda value: format_euro(float(value)),
                        "FACTURACION_NETA": lambda value: format_euro(float(value)),
                        "PRIMAS_EMITIDAS": lambda value: format_euro(float(value)),
                        "PRIMAS_ANULADAS": lambda value: format_euro(float(value)),
                        "PRIMAS_NETAS": lambda value: format_euro(float(value)),
                        "IMPORTE_SINIESTROS": lambda value: format_euro(float(value)),
                        "SINIESTRALIDAD": lambda value: format_percent(float(value)),
                        "PRIMA_MEDIA_NETA": lambda value: format_euro(float(value)),
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    else:
        st.subheader("Detalles")

        with st.expander("Detalle de altas Decesos incluidas"):
            st.dataframe(
                altas_detail[detail_columns_for_display(altas_detail)].style.format(
                    {"PRIMA_NETA_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle de anulaciones Decesos incluidas"):
            st.dataframe(
                anulaciones_detail[detail_columns_for_display(anulaciones_detail)].style.format(
                    {"PRIMA_NETA_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle de siniestros Decesos incluidos"):
            st.caption("El importe usado para siniestralidad es RESERACT_VALOR + PAGOSPDT_VALOR. La fecha usada para año/mes es FECHOCUR.")
            st.dataframe(
                siniestros_detail[siniestros_columns_for_display(siniestros_detail)].style.format(
                    {
                        "PAGOSPDT_VALOR": lambda value: format_euro(float(value)),
                        "PAGOSRZD_VALOR": lambda value: format_euro(float(value)),
                        "COSTESIN_VALOR": lambda value: format_euro(float(value)),
                        "IMPORTE_SINIESTRO": lambda value: format_euro(float(value)),
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle de primas emitidas Decesos"):
            st.dataframe(
                primas_emitidas_detail[primas_columns_for_display(primas_emitidas_detail)].style.format(
                    {"POLIPNET_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle de primas anuladas Decesos"):
            st.dataframe(
                primas_anuladas_detail[primas_columns_for_display(primas_anuladas_detail)].style.format(
                    {"POLIPNET_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Salud - facturacion bruta"):
            st.dataframe(
                salud_bruta_detail[salud_columns_for_display(salud_bruta_detail)].style.format(
                    {"PRIMA_NETA_SALUD_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Salud - anulaciones"):
            st.dataframe(
                salud_anulaciones_detail[salud_columns_for_display(salud_anulaciones_detail)].style.format(
                    {"PRIMA_NETA_SALUD_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Vida - facturacion bruta"):
            st.dataframe(
                vida_bruta_detail[vida_columns_for_display(vida_bruta_detail)].style.format(
                    {"PRIMA_VIDA_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Vida - anulaciones"):
            st.dataframe(
                vida_anulaciones_detail[vida_columns_for_display(vida_anulaciones_detail)].style.format(
                    {"PRIMA_VIDA_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

    excel_bytes = dataframe_to_excel(
        ranking,
        ranking_salud,
        ranking_vida,
        ranking_decesos_top10,
        ranking_salud_top10,
        ranking_vida_top10,
        altas_detail,
        anulaciones_detail,
        siniestros_detail,
        primas_emitidas_detail,
        primas_anuladas_detail,
        salud_bruta_detail,
        salud_anulaciones_detail,
        vida_bruta_detail,
        vida_anulaciones_detail,
        sheet_summary,
        ranking_date,
        excluded_products,
    )

    st.download_button(
        "Descargar ranking neto con Salud y Vida en Excel",
        data=excel_bytes,
        file_name=f"ranking_decesos_salud_vida_{ranking_date:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
