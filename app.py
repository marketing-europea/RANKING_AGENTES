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

REQUIRED_SINIESTROS_COLUMNS = (
    "PRODUCTO",
    "CODIMEDI",
    "FECHDECL",
    "PAGOSRZD",
    "COSTESIN",
)

REQUIRED_PRIMAS_COLUMNS = (
    "MEDIADOR",
    "POLIPNET",
)

REQUIRED_CARTERA_SALUD_COLUMNS = (
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

    work["FECHA_SINIESTRO"] = pd.to_datetime(
        work["FECHDECL"],
        dayfirst=True,
        errors="coerce",
    )

    work["ANIO_SINIESTRO"] = work["FECHA_SINIESTRO"].dt.year
    work["MES_SINIESTRO"] = work["FECHA_SINIESTRO"].dt.month

    work["PAGOSRZD_VALOR"] = work["PAGOSRZD"].apply(parse_spanish_number)
    work["COSTESIN_VALOR"] = work["COSTESIN"].apply(parse_spanish_number)
    work["IMPORTE_SINIESTRO"] = work["PAGOSRZD_VALOR"] + work["COSTESIN_VALOR"]

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


def classify_liga_decesos(facturacion_decesos: float, cumple_siniestralidad: bool) -> str:
    if not cumple_siniestralidad:
        return "No cumple siniestralidad"

    if facturacion_decesos >= DECESOS_LIGA_ELITE_MINIMA:
        return "LIGA ELITE"

    if facturacion_decesos >= DECESOS_LIGA_PRO_MINIMA:
        return "LIGA PRO"

    return "No clasifica"


def calculate_decesos_ranking(
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
                "LIGA_DECESOS",
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

    ranking["CUMPLE_SINIESTRALIDAD"] = ranking["SINIESTRALIDAD"] <= SINIESTRALIDAD_MAXIMA

    ranking["LIGA_DECESOS"] = [
        classify_liga_decesos(facturacion, cumple)
        for facturacion, cumple in zip(
            ranking["FACTURACION_NETA"],
            ranking["CUMPLE_SINIESTRALIDAD"],
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
            "PRIMAS_EMITIDAS",
            "PRIMAS_ANULADAS",
            "PRIMAS_NETAS",
            "IMPORTE_SINIESTROS",
            "SINIESTRALIDAD",
            "CUMPLE_SINIESTRALIDAD",
            "LIGA_DECESOS",
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


def build_ranking_ligas_decesos(ranking_decesos: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "RANKING_LIGA_DECESOS",
        "LIGA_DECESOS",
        "AGENTE",
        "NOMBRE_AGENCIA",
        "FACTURACION_DECESOS_NETA",
        "SINIESTRALIDAD",
        "CUMPLE_SINIESTRALIDAD",
        "IMPORTE_SINIESTROS",
        "PRIMAS_NETAS",
    ]

    if ranking_decesos.empty:
        return pd.DataFrame(columns=columns)

    ranking_ligas = ranking_decesos[
        ranking_decesos["LIGA_DECESOS"].isin(["LIGA ELITE", "LIGA PRO"])
    ].copy()

    if ranking_ligas.empty:
        return pd.DataFrame(columns=columns)

    ranking_ligas = ranking_ligas.rename(
        columns={"FACTURACION_NETA": "FACTURACION_DECESOS_NETA"}
    )

    liga_order = {"LIGA ELITE": 1, "LIGA PRO": 2}
    ranking_ligas["ORDEN_LIGA"] = ranking_ligas["LIGA_DECESOS"].map(liga_order)

    ranking_ligas = ranking_ligas.sort_values(
        ["ORDEN_LIGA", "FACTURACION_DECESOS_NETA", "SINIESTRALIDAD", "AGENTE"],
        ascending=[True, False, True, True],
    )

    ranking_ligas.insert(0, "RANKING_LIGA_DECESOS", range(1, len(ranking_ligas) + 1))

    return ranking_ligas[columns]


# =========================
# SALUD
# =========================

def prepare_cartera_salud_data(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    work["RAMO"] = "SALUD"
    work["AGENTE"] = work["MEDIADOR"].apply(normalize_agent)
    work["POLIZA_NORMALIZADA"] = work["IDPOLIZA"].apply(lambda value: normalize_text(value, ""))
    work["FECHA_EFECTO"] = pd.to_datetime(
        work["POLIEFEC"],
        dayfirst=True,
        errors="coerce",
    )
    work["ANIO_EFECTO"] = work["FECHA_EFECTO"].dt.year
    work["MES_EFECTO"] = work["FECHA_EFECTO"].dt.month
    work["PRIMA_NETA_VALOR"] = work["PRIMA NETA"].apply(parse_spanish_number)

    return work


def prepare_bajas_salud_data(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    work["POLIZA_NORMALIZADA"] = work["POLIZA"].apply(lambda value: normalize_text(value, ""))
    work["PRODUCTO_BAJA_NORMALIZADO"] = work["DES_PRODUCTO"].apply(normalize_product)
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


def calculate_salud_ranking(
    cartera_salud_df: pd.DataFrame,
    bajas_salud_df: pd.DataFrame,
    ranking_decesos: pd.DataFrame,
    ranking_date: date,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    cartera_detail = prepare_cartera_salud_data(cartera_salud_df)
    bajas_detail = prepare_bajas_salud_data(bajas_salud_df)

    altas_mask = (
        cartera_detail["FECHA_EFECTO"].notna()
        & cartera_detail["ANIO_EFECTO"].eq(ranking_date.year)
        & cartera_detail["MES_EFECTO"].le(ranking_date.month)
    )

    salud_altas_detail = cartera_detail[altas_mask].copy()

    bajas_mask = (
        bajas_detail["FECHA_BAJA_SALUD"].notna()
        & bajas_detail["ANIO_BAJA_SALUD"].eq(ranking_date.year)
        & bajas_detail["MES_BAJA_SALUD"].le(ranking_date.month)
        & bajas_detail["FECHA_REACTIVACION_SALUD"].isna()
        & ~bajas_detail["PRODUCTO_BAJA_NORMALIZADO"].eq("ASISA VIDA RIESGO")
    )

    bajas_validas = bajas_detail[bajas_mask].copy()
    bajas_validas = bajas_validas.drop_duplicates("POLIZA_NORMALIZADA")

    salud_bajas_detail = pd.merge(
        cartera_detail,
        bajas_validas,
        on="POLIZA_NORMALIZADA",
        how="inner",
        suffixes=("", "_BAJA"),
    )

    altas = (
        salud_altas_detail.groupby("AGENTE", dropna=False)
        .agg(
            FACTURACION_SALUD_BRUTA=("PRIMA_NETA_VALOR", "sum"),
            POLIZAS_SALUD_ALTAS=("POLIZA_NORMALIZADA", "nunique"),
        )
        .reset_index()
    )

    bajas = (
        salud_bajas_detail.groupby("AGENTE", dropna=False)
        .agg(
            FACTURACION_SALUD_ANULACIONES=("PRIMA_NETA_VALOR", "sum"),
            POLIZAS_SALUD_ANULADAS=("POLIZA_NORMALIZADA", "nunique"),
        )
        .reset_index()
    )

    ranking_salud = pd.merge(altas, bajas, on="AGENTE", how="outer")

    if ranking_salud.empty:
        empty = pd.DataFrame(
            columns=[
                "RANKING_SALUD",
                "AGENTE",
                "NOMBRE_AGENCIA",
                "FACTURACION_SALUD_BRUTA",
                "FACTURACION_SALUD_ANULACIONES",
                "FACTURACION_SALUD_NETA",
                "FACTURACION_DECESOS_NETA",
                "LIGA_SALUD",
                "POLIZAS_SALUD_ALTAS",
                "POLIZAS_SALUD_ANULADAS",
                "POLIZAS_SALUD_NETAS",
            ]
        )
        return empty, salud_altas_detail, salud_bajas_detail

    for column in [
        "FACTURACION_SALUD_BRUTA",
        "FACTURACION_SALUD_ANULACIONES",
        "POLIZAS_SALUD_ALTAS",
        "POLIZAS_SALUD_ANULADAS",
    ]:
        ranking_salud[column] = ranking_salud[column].fillna(0)

    ranking_salud["FACTURACION_SALUD_NETA"] = (
        ranking_salud["FACTURACION_SALUD_BRUTA"]
        - ranking_salud["FACTURACION_SALUD_ANULACIONES"]
    )
    ranking_salud["POLIZAS_SALUD_NETAS"] = (
        ranking_salud["POLIZAS_SALUD_ALTAS"]
        - ranking_salud["POLIZAS_SALUD_ANULADAS"]
    )

    if ranking_decesos.empty:
        decesos_base = pd.DataFrame(columns=["AGENTE", "NOMBRE_AGENCIA", "FACTURACION_DECESOS_NETA"])
    else:
        decesos_base = ranking_decesos[
            ["AGENTE", "NOMBRE_AGENCIA", "FACTURACION_NETA"]
        ].rename(columns={"FACTURACION_NETA": "FACTURACION_DECESOS_NETA"})

    ranking_salud = pd.merge(ranking_salud, decesos_base, on="AGENTE", how="left")
    ranking_salud["FACTURACION_DECESOS_NETA"] = ranking_salud["FACTURACION_DECESOS_NETA"].fillna(0)
    ranking_salud["NOMBRE_AGENCIA"] = [
        name if not pd.isna(name) and str(name).strip() != "" else agent
        for name, agent in zip(ranking_salud["NOMBRE_AGENCIA"], ranking_salud["AGENTE"])
    ]

    ranking_salud["LIGA_SALUD"] = [
        classify_liga_salud(facturacion_salud, facturacion_decesos)
        for facturacion_salud, facturacion_decesos in zip(
            ranking_salud["FACTURACION_SALUD_NETA"],
            ranking_salud["FACTURACION_DECESOS_NETA"],
        )
    ]

    ranking_salud = ranking_salud[
        [
            "AGENTE",
            "NOMBRE_AGENCIA",
            "FACTURACION_SALUD_BRUTA",
            "FACTURACION_SALUD_ANULACIONES",
            "FACTURACION_SALUD_NETA",
            "FACTURACION_DECESOS_NETA",
            "LIGA_SALUD",
            "POLIZAS_SALUD_ALTAS",
            "POLIZAS_SALUD_ANULADAS",
            "POLIZAS_SALUD_NETAS",
        ]
    ].sort_values(
        ["FACTURACION_SALUD_NETA", "FACTURACION_DECESOS_NETA", "AGENTE"],
        ascending=[False, False, True],
    )

    ranking_salud.insert(0, "RANKING_SALUD", range(1, len(ranking_salud) + 1))

    return ranking_salud, salud_altas_detail, salud_bajas_detail


def build_ranking_ligas_salud(ranking_salud: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "RANKING_LIGA_SALUD",
        "LIGA_SALUD",
        "AGENTE",
        "NOMBRE_AGENCIA",
        "FACTURACION_SALUD_NETA",
        "FACTURACION_DECESOS_NETA",
    ]

    if ranking_salud.empty:
        return pd.DataFrame(columns=columns)

    ranking_ligas = ranking_salud[
        ranking_salud["LIGA_SALUD"].isin(["LIGA ELITE", "LIGA PRO"])
    ].copy()

    if ranking_ligas.empty:
        return pd.DataFrame(columns=columns)

    liga_order = {"LIGA ELITE": 1, "LIGA PRO": 2}
    ranking_ligas["ORDEN_LIGA"] = ranking_ligas["LIGA_SALUD"].map(liga_order)

    ranking_ligas = ranking_ligas.sort_values(
        ["ORDEN_LIGA", "FACTURACION_SALUD_NETA", "FACTURACION_DECESOS_NETA", "AGENTE"],
        ascending=[True, False, False, True],
    )

    ranking_ligas.insert(0, "RANKING_LIGA_SALUD", range(1, len(ranking_ligas) + 1))

    return ranking_ligas[columns]


# =========================
# OUTPUT / DISPLAY
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
        "FECHDECL",
        "FECHA_SINIESTRO",
        "ANIO_SINIESTRO",
        "MES_SINIESTRO",
        "NUMESINI",
        "POLIZSEC",
        "ESTADO",
        "MOTIVO",
        "COBERTURA",
        "NATURALEZA",
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


def salud_altas_columns_for_display(detail: pd.DataFrame) -> list[str]:
    columns = [
        "ARCHIVO_ORIGEN",
        "HOJA_ORIGEN",
        "PRODUCTO",
        "IDPOLIZA",
        "POLIZA_NORMALIZADA",
        "MEDIADOR",
        "AGENTE",
        "POLIEFEC",
        "FECHA_EFECTO",
        "PRIMA NETA",
        "PRIMA_NETA_VALOR",
        "ESTADO",
        "TOMADOR",
        "NIF",
    ]

    return [column for column in columns if column in detail.columns]


def salud_bajas_columns_for_display(detail: pd.DataFrame) -> list[str]:
    columns = [
        "ARCHIVO_ORIGEN",
        "HOJA_ORIGEN",
        "PRODUCTO",
        "IDPOLIZA",
        "POLIZA_NORMALIZADA",
        "MEDIADOR",
        "AGENTE",
        "POLIEFEC",
        "FECHA_EFECTO",
        "PRIMA NETA",
        "PRIMA_NETA_VALOR",
        "DES_PRODUCTO",
        "FEC_EFECTO_BAJA",
        "FEC_EFECTO_REACTIV",
        "MOTIVO_BAJA",
    ]

    return [column for column in columns if column in detail.columns]


def dataframe_to_excel(
    ranking_decesos: pd.DataFrame,
    ranking_salud: pd.DataFrame,
    altas_detail: pd.DataFrame,
    anulaciones_detail: pd.DataFrame,
    siniestros_detail: pd.DataFrame,
    primas_emitidas_detail: pd.DataFrame,
    primas_anuladas_detail: pd.DataFrame,
    salud_altas_detail: pd.DataFrame,
    salud_bajas_detail: pd.DataFrame,
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
            {"CAMPO": "PRODUCTOS_EXCLUIDOS_DECESOS", "VALOR": ", ".join(excluded_products)},
            {"CAMPO": "DECESOS_LIGA_PRO", "VALOR": "Facturacion Decesos >= 30.000 y siniestralidad <= 25%"},
            {"CAMPO": "DECESOS_LIGA_ELITE", "VALOR": "Facturacion Decesos >= 60.000 y siniestralidad <= 25%"},
            {"CAMPO": "SALUD_LIGA_PRO", "VALOR": "Facturacion Salud >= 25.000 y Facturacion Decesos >= 12.000"},
            {"CAMPO": "SALUD_LIGA_ELITE", "VALOR": "Facturacion Salud >= 80.000 y Facturacion Decesos >= 4.000"},
            {"CAMPO": "SALUD_ALTAS", "VALOR": "Cartera Salud por POLIEFEC"},
            {"CAMPO": "SALUD_ANULACIONES", "VALOR": "Bajas Salud por FEC_EFECTO_BAJA, sin FEC_EFECTO_REACTIV, excluyendo ASISA VIDA RIESGO"},
        ]
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        parametros.to_excel(writer, index=False, sheet_name="PARAMETROS")
        sheet_summary.to_excel(writer, index=False, sheet_name="COMPROBACION_HOJAS")

        build_ranking_ligas_decesos(ranking_decesos).to_excel(
            writer,
            index=False,
            sheet_name="LIGAS_DECESOS",
        )
        build_ranking_ligas_salud(ranking_salud).to_excel(
            writer,
            index=False,
            sheet_name="LIGAS_SALUD",
        )
        ranking_decesos.to_excel(writer, index=False, sheet_name="RANKING_DECESOS")
        ranking_salud.to_excel(writer, index=False, sheet_name="RANKING_SALUD")

        altas_detail[detail_columns_for_display(altas_detail)].to_excel(
            writer,
            index=False,
            sheet_name="DETALLE_ALTAS_DECESOS",
        )
        anulaciones_detail[detail_columns_for_display(anulaciones_detail)].to_excel(
            writer,
            index=False,
            sheet_name="DETALLE_ANUL_DECESOS",
        )
        siniestros_detail[siniestros_columns_for_display(siniestros_detail)].to_excel(
            writer,
            index=False,
            sheet_name="DETALLE_SINIESTROS",
        )
        primas_emitidas_detail[primas_columns_for_display(primas_emitidas_detail)].to_excel(
            writer,
            index=False,
            sheet_name="DETALLE_PRIMAS_EMIT",
        )
        primas_anuladas_detail[primas_columns_for_display(primas_anuladas_detail)].to_excel(
            writer,
            index=False,
            sheet_name="DETALLE_PRIMAS_ANUL",
        )
        salud_altas_detail[salud_altas_columns_for_display(salud_altas_detail)].to_excel(
            writer,
            index=False,
            sheet_name="DETALLE_SALUD_ALTAS",
        )
        salud_bajas_detail[salud_bajas_columns_for_display(salud_bajas_detail)].to_excel(
            writer,
            index=False,
            sheet_name="DETALLE_SALUD_BAJAS",
        )

    return output.getvalue()


# =========================
# STREAMLIT APP
# =========================

def main() -> None:
    st.set_page_config(page_title="Ranking agentes", layout="wide")

    st.title("Ranking agentes - Decesos y Salud")

    ranking_date = st.date_input(
        "Para que fecha quieres calcular el ranking?",
        value=DEFAULT_RANKING_DATE,
        format="DD/MM/YYYY",
    )

    st.sidebar.title("Vista")

    st.header("Archivos Decesos")
    col_upload_1, col_upload_2, col_upload_3 = st.columns(3)
    col_upload_4, col_upload_5 = st.columns(2)

    with col_upload_1:
        uploaded_facturacion = st.file_uploader(
            "Sube FACTURACION_DECESOS.xls",
            type=["xls", "xlsx"],
            key="facturacion_decesos",
        )

    with col_upload_2:
        uploaded_anulaciones = st.file_uploader(
            "Sube FACTURACION_ANULACIONES_DECESOS.xls",
            type=["xls", "xlsx"],
            key="anulaciones_decesos",
        )

    with col_upload_3:
        uploaded_siniestros = st.file_uploader(
            "Sube SINIESTROS_DECESOS.xls",
            type=["xls", "xlsx"],
            key="siniestros_decesos",
        )

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

    st.header("Archivos Salud")
    col_upload_6, col_upload_7 = st.columns(2)

    with col_upload_6:
        uploaded_cartera_salud = st.file_uploader(
            "Sube CARTERA_SALUD.xls",
            type=["xls", "xlsx"],
            key="cartera_salud",
        )

    with col_upload_7:
        uploaded_bajas_salud = st.file_uploader(
            "Sube BAJAS_SALUD.xls",
            type=["xls", "xlsx"],
            key="bajas_salud",
        )

    if (
        uploaded_facturacion is None
        or uploaded_anulaciones is None
        or uploaded_siniestros is None
        or not uploaded_primas_emitidas
        or not uploaded_primas_anuladas
        or uploaded_cartera_salud is None
        or uploaded_bajas_salud is None
    ):
        st.info("Sube todos los archivos de Decesos y Salud para calcular los rankings.")
        st.stop()

    try:
        raw_facturacion_df = read_excel_all_sheets(uploaded_facturacion)
        raw_anulaciones_df = read_excel_all_sheets(uploaded_anulaciones)
        raw_siniestros_df = read_excel_all_sheets(uploaded_siniestros)
        raw_primas_emitidas_df = read_excel_many_files(uploaded_primas_emitidas)
        raw_primas_anuladas_df = read_excel_many_files(uploaded_primas_anuladas)
        raw_cartera_salud_df = read_excel_all_sheets(uploaded_cartera_salud)
        raw_bajas_salud_df = read_excel_all_sheets(uploaded_bajas_salud)
    except Exception as error:
        st.error(f"No he podido leer los archivos: {error}")
        st.stop()

    if (
        raw_facturacion_df.empty
        or raw_anulaciones_df.empty
        or raw_siniestros_df.empty
        or raw_primas_emitidas_df.empty
        or raw_primas_anuladas_df.empty
        or raw_cartera_salud_df.empty
        or raw_bajas_salud_df.empty
    ):
        st.warning("Alguno de los archivos esta vacio.")
        st.stop()

    missing_facturacion = validate_columns(raw_facturacion_df, REQUIRED_FACTURACION_COLUMNS)
    missing_anulaciones = validate_columns(raw_anulaciones_df, REQUIRED_ANULACIONES_COLUMNS)
    missing_siniestros = validate_columns(raw_siniestros_df, REQUIRED_SINIESTROS_COLUMNS)
    missing_primas_emitidas = validate_columns(raw_primas_emitidas_df, REQUIRED_PRIMAS_COLUMNS)
    missing_primas_anuladas = validate_columns(raw_primas_anuladas_df, REQUIRED_PRIMAS_COLUMNS)
    missing_cartera_salud = validate_columns(raw_cartera_salud_df, REQUIRED_CARTERA_SALUD_COLUMNS)
    missing_bajas_salud = validate_columns(raw_bajas_salud_df, REQUIRED_BAJAS_SALUD_COLUMNS)

    if (
        missing_facturacion
        or missing_anulaciones
        or missing_siniestros
        or missing_primas_emitidas
        or missing_primas_anuladas
        or missing_cartera_salud
        or missing_bajas_salud
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
        if missing_cartera_salud:
            messages.append(f"Cartera Salud: {', '.join(missing_cartera_salud)}")
        if missing_bajas_salud:
            messages.append(f"Bajas Salud: {', '.join(missing_bajas_salud)}")

        st.error("Faltan columnas obligatorias. " + " | ".join(messages))
        st.stop()

    sheet_summary = pd.concat(
        [
            build_sheet_summary(raw_facturacion_df, "FACTURACION_DECESOS"),
            build_sheet_summary(raw_anulaciones_df, "FACTURACION_ANULACIONES_DECESOS"),
            build_sheet_summary(raw_siniestros_df, "SINIESTROS_DECESOS"),
            build_sheet_summary(raw_primas_emitidas_df, "PRIMAS_EMITIDAS"),
            build_sheet_summary(raw_primas_anuladas_df, "PRIMAS_ANULADAS"),
            build_sheet_summary(raw_cartera_salud_df, "CARTERA_SALUD"),
            build_sheet_summary(raw_bajas_salud_df, "BAJAS_SALUD"),
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
                raw_siniestros_df["PRODUCTO"],
            ],
            ignore_index=True,
        )
        .dropna()
        .map(normalize_product)
        .unique()
        .tolist()
    )

    default_excluded = [product for product in DEFAULT_EXCLUDED_PRODUCTS if product in product_options]

    excluded_products = st.multiselect(
        "Productos excluidos en Decesos",
        product_options,
        default=default_excluded,
    )

    (
        ranking_decesos,
        altas_detail,
        anulaciones_detail,
        siniestros_detail,
        primas_emitidas_detail,
        primas_anuladas_detail,
    ) = calculate_decesos_ranking(
        raw_facturacion_df,
        raw_anulaciones_df,
        raw_siniestros_df,
        raw_primas_emitidas_df,
        raw_primas_anuladas_df,
        ranking_date,
        excluded_products,
    )

    ranking_salud, salud_altas_detail, salud_bajas_detail = calculate_salud_ranking(
        raw_cartera_salud_df,
        raw_bajas_salud_df,
        ranking_decesos,
        ranking_date,
    )

    ligas_decesos = build_ranking_ligas_decesos(ranking_decesos)
    ligas_salud = build_ranking_ligas_salud(ranking_salud)

    total_decesos_neta = float(ranking_decesos["FACTURACION_NETA"].sum()) if not ranking_decesos.empty else 0.0
    total_salud_neta = float(ranking_salud["FACTURACION_SALUD_NETA"].sum()) if not ranking_salud.empty else 0.0
    total_salud_bruta = float(ranking_salud["FACTURACION_SALUD_BRUTA"].sum()) if not ranking_salud.empty else 0.0
    total_salud_anulaciones = float(ranking_salud["FACTURACION_SALUD_ANULACIONES"].sum()) if not ranking_salud.empty else 0.0
    total_primas_netas = float(ranking_decesos["PRIMAS_NETAS"].sum()) if not ranking_decesos.empty else 0.0
    total_siniestros = float(ranking_decesos["IMPORTE_SINIESTROS"].sum()) if not ranking_decesos.empty else 0.0
    siniestralidad_total = total_siniestros / total_primas_netas if total_primas_netas > 0 else 0.0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Facturacion neta Decesos", format_euro(total_decesos_neta))
    col2.metric("Facturacion bruta Salud", format_euro(total_salud_bruta))
    col3.metric("Anulaciones Salud", format_euro(total_salud_anulaciones))
    col4.metric("Facturacion neta Salud", format_euro(total_salud_neta))

    col5, col6, col7 = st.columns(3)
    col5.metric("Primas netas Decesos", format_euro(total_primas_netas))
    col6.metric("Importe siniestros Decesos", format_euro(total_siniestros))
    col7.metric("Siniestralidad Decesos", format_percent(siniestralidad_total))

    vista = st.sidebar.radio(
        "Ranking",
        [
            "Ligas Decesos",
            "Ligas Salud",
            "Ranking completo Decesos",
            "Ranking completo Salud",
            "Detalles",
        ],
    )

    if vista == "Ligas Decesos":
        st.subheader("Ligas Decesos")
        st.caption("LIGA PRO: facturacion Decesos >= 30.000 € y siniestralidad <= 25%. LIGA ELITE: facturacion Decesos >= 60.000 € y siniestralidad <= 25%.")

        if ligas_decesos.empty:
            st.info("No hay mediadores clasificados en ligas de Decesos.")
        else:
            st.dataframe(
                ligas_decesos.style.format(
                    {
                        "FACTURACION_DECESOS_NETA": lambda value: format_euro(float(value)),
                        "SINIESTRALIDAD": lambda value: format_percent(float(value)),
                        "IMPORTE_SINIESTROS": lambda value: format_euro(float(value)),
                        "PRIMAS_NETAS": lambda value: format_euro(float(value)),
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    elif vista == "Ligas Salud":
        st.subheader("Ligas Salud")
        st.caption("LIGA PRO Salud: Salud >= 25.000 € y Decesos >= 12.000 €. LIGA ELITE Salud: Salud >= 80.000 € y Decesos >= 4.000 €.")

        if ligas_salud.empty:
            st.info("No hay mediadores clasificados en ligas de Salud.")
        else:
            st.dataframe(
                ligas_salud.style.format(
                    {
                        "FACTURACION_SALUD_NETA": lambda value: format_euro(float(value)),
                        "FACTURACION_DECESOS_NETA": lambda value: format_euro(float(value)),
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    elif vista == "Ranking completo Decesos":
        st.subheader("Ranking completo Decesos")

        if ranking_decesos.empty:
            st.info("No hay datos de Decesos para los filtros seleccionados.")
        else:
            st.dataframe(
                ranking_decesos.style.format(
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

    elif vista == "Ranking completo Salud":
        st.subheader("Ranking completo Salud")

        if ranking_salud.empty:
            st.info("No hay datos de Salud para los filtros seleccionados.")
        else:
            st.dataframe(
                ranking_salud.style.format(
                    {
                        "FACTURACION_SALUD_BRUTA": lambda value: format_euro(float(value)),
                        "FACTURACION_SALUD_ANULACIONES": lambda value: format_euro(float(value)),
                        "FACTURACION_SALUD_NETA": lambda value: format_euro(float(value)),
                        "FACTURACION_DECESOS_NETA": lambda value: format_euro(float(value)),
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    else:
        st.subheader("Detalles")

        with st.expander("Detalle Decesos - altas incluidas"):
            st.dataframe(
                altas_detail[detail_columns_for_display(altas_detail)].style.format(
                    {"PRIMA_NETA_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Decesos - anulaciones incluidas"):
            st.dataframe(
                anulaciones_detail[detail_columns_for_display(anulaciones_detail)].style.format(
                    {"PRIMA_NETA_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Decesos - siniestros incluidos"):
            st.dataframe(
                siniestros_detail[siniestros_columns_for_display(siniestros_detail)].style.format(
                    {
                        "PAGOSRZD_VALOR": lambda value: format_euro(float(value)),
                        "COSTESIN_VALOR": lambda value: format_euro(float(value)),
                        "IMPORTE_SINIESTRO": lambda value: format_euro(float(value)),
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Decesos - primas emitidas"):
            st.dataframe(
                primas_emitidas_detail[primas_columns_for_display(primas_emitidas_detail)].style.format(
                    {"POLIPNET_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Decesos - primas anuladas"):
            st.dataframe(
                primas_anuladas_detail[primas_columns_for_display(primas_anuladas_detail)].style.format(
                    {"POLIPNET_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Salud - cartera bruta incluida"):
            st.dataframe(
                salud_altas_detail[salud_altas_columns_for_display(salud_altas_detail)].style.format(
                    {"PRIMA_NETA_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Detalle Salud - bajas restadas"):
            st.dataframe(
                salud_bajas_detail[salud_bajas_columns_for_display(salud_bajas_detail)].style.format(
                    {"PRIMA_NETA_VALOR": lambda value: format_euro(float(value))}
                ),
                use_container_width=True,
                hide_index=True,
            )

    excel_bytes = dataframe_to_excel(
        ranking_decesos,
        ranking_salud,
        altas_detail,
        anulaciones_detail,
        siniestros_detail,
        primas_emitidas_detail,
        primas_anuladas_detail,
        salud_altas_detail,
        salud_bajas_detail,
        sheet_summary,
        ranking_date,
        excluded_products,
    )

    st.download_button(
        "Descargar ranking Decesos y Salud en Excel",
        data=excel_bytes,
        file_name=f"ranking_decesos_salud_{ranking_date:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
