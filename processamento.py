from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
import re
import unicodedata
from typing import Iterable, Sequence

import pandas as pd


STATUS_PRIORIDADE = {
    "Missing Cargo": 1,
    "Discrepância Criada": 2,
    "Pendente Embarque": 3,
    "Pendente Desembarque": 4,
    "Pendente Entrega": 5,
    "Atribuído à Rota": 6,
    "Saído para Entrega": 7,
    "Entregue": 8,
    "Baixado": 9,
    "Outros": 99,
}

STATUS_PENDENCIAS = [
    "Pendente Embarque",
    "Pendente Entrega",
    "Pendente Desembarque",
    "Missing Cargo",
    "Discrepância Criada",
]

COLUNAS_OBRIGATORIAS = {
    "AWB",
    "OriginCode",
    "ExecutionDateTime",
    "OPSStation",
    "StatusDescription",
    "ApproxSLA",
}

COLUNAS_DATA = ["ExecutionDateTime", "FltDt", "ApproxSLA", "DeliveryRequest"]
COLUNAS_NUMERICAS = ["PiecesCount", "No of Pieces", "GrossWt"]


@dataclass(frozen=True)
class ResultadoProcessamento:
    bruto: pd.DataFrame
    consolidado: pd.DataFrame
    pendencias: pd.DataFrame
    inconsistencias: pd.DataFrame
    arquivos: pd.DataFrame
    data_referencia: pd.Timestamp


def _sem_acento(valor: object) -> str:
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def normalizar_texto(valor: object) -> str:
    texto = _sem_acento(valor).upper().strip()
    return re.sub(r"\s+", " ", texto)


def _somente_digitos(valor: object) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return re.sub(r"\D", "", texto)


def padronizar_awb(valor: object) -> str:
    digitos = _somente_digitos(valor)
    if not digitos:
        return ""
    return digitos.zfill(8) if len(digitos) <= 8 else digitos


def padronizar_prefixo(valor: object) -> str:
    return _somente_digitos(valor)


def _converter_data(serie: pd.Series) -> pd.Series:
    convertido = pd.to_datetime(serie, errors="coerce", dayfirst=True, format="mixed")
    numericos = pd.to_numeric(serie, errors="coerce")
    mascara = convertido.isna() & numericos.notna()
    if mascara.any():
        convertido.loc[mascara] = pd.to_datetime(
            numericos.loc[mascara], unit="D", origin="1899-12-30", errors="coerce"
        )
    return convertido


def _status_canonico(valor: object) -> str:
    texto = normalizar_texto(valor)
    if "MISSING" in texto:
        return "Missing Cargo"
    if "DISCREP" in texto:
        return "Discrepância Criada"
    if "PENDENTE" in texto and "DESEMBAR" in texto:
        return "Pendente Desembarque"
    if "PENDENTE" in texto and "EMBAR" in texto:
        return "Pendente Embarque"
    if "PENDENTE" in texto and "ENTREGA" in texto:
        return "Pendente Entrega"
    if "ATRIBU" in texto and "ROTA" in texto:
        return "Atribuído à Rota"
    if "SAIDO" in texto and "ENTREGA" in texto:
        return "Saído para Entrega"
    if "ENTREG" in texto:
        return "Entregue"
    if "BAIXAD" in texto:
        return "Baixado"
    return "Outros"


def _identificar_cliente(linha: pd.Series) -> str:
    campos = [
        linha.get("BillTo", ""),
        linha.get("Shipper", ""),
        linha.get("Consignee", ""),
        linha.get("ServiceTaker", ""),
        linha.get("Consolidator", ""),
        linha.get("NotifyParty", ""),
    ]
    texto = normalizar_texto(" | ".join(str(v) for v in campos if pd.notna(v)))
    origem = normalizar_texto(linha.get("OriginCode", ""))

    if origem == "TRES1" or "TRES CORACOES" in texto or "3 CORACOES" in texto:
        return "TRÊS CORAÇÕES"
    if "RIACHUELO" in texto:
        return "RIACHUELO"
    if "TANIA BULHO" in texto or "TB COMERCIO DE PRESENTES" in texto:
        return "TANIA BULHÕES"
    if "DELLA VIA" in texto:
        return "DELLA VIA"
    return "OUTROS / NÃO MAPEADO"


def _identificar_operacao(origem: object) -> str:
    codigo = normalizar_texto(origem)
    if codigo == "TRES1":
        return "TRES1"
    if codigo == "SAO12":
        return "SAO12"
    return codigo or "NÃO INFORMADA"


def _ler_excel(arquivo) -> pd.DataFrame:
    # UploadedFile do Streamlit e BytesIO são aceitos diretamente pelo pandas.
    if hasattr(arquivo, "seek"):
        arquivo.seek(0)
    return pd.read_excel(arquivo, sheet_name=0, dtype=object, engine="openpyxl")


def validar_colunas(df: pd.DataFrame) -> list[str]:
    colunas = {str(c).strip() for c in df.columns}
    return sorted(COLUNAS_OBRIGATORIAS - colunas)


def preparar_arquivo_status(arquivo, nome_arquivo: str, data_referencia: pd.Timestamp) -> pd.DataFrame:
    df = _ler_excel(arquivo)
    df.columns = [str(c).strip() for c in df.columns]

    faltantes = validar_colunas(df)
    if faltantes:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não possui as colunas obrigatórias: "
            + ", ".join(faltantes)
        )

    for coluna in COLUNAS_DATA:
        if coluna in df.columns:
            df[coluna] = _converter_data(df[coluna])

    for coluna in COLUNAS_NUMERICAS:
        if coluna in df.columns:
            df[coluna] = pd.to_numeric(df[coluna], errors="coerce")

    if "AWBPrefix" not in df.columns:
        df["AWBPrefix"] = ""

    df["AWB"] = df["AWB"].map(padronizar_awb)
    df["AWBPrefix"] = df["AWBPrefix"].map(padronizar_prefixo)
    df["AWBCompleta"] = df["AWBPrefix"] + df["AWB"]
    df.loc[df["AWBPrefix"].eq(""), "AWBCompleta"] = df.loc[
        df["AWBPrefix"].eq(""), "AWB"
    ]

    df["StatusOperacional"] = df["StatusDescription"].map(_status_canonico)
    df["PrioridadeStatus"] = df["StatusOperacional"].map(STATUS_PRIORIDADE).fillna(99).astype(int)
    df["Cliente"] = df.apply(_identificar_cliente, axis=1)
    df["Operacao"] = df["OriginCode"].map(_identificar_operacao)
    df["ArquivoOrigem"] = nome_arquivo

    horas = (df["ApproxSLA"] - data_referencia).dt.total_seconds() / 3600
    df["HorasParaSLA"] = horas.round(1)
    df["SLAStatus"] = "NO PRAZO"
    df.loc[df["ApproxSLA"].isna(), "SLAStatus"] = "SEM SLA"
    df.loc[df["ApproxSLA"].notna() & (horas < 0), "SLAStatus"] = "VENCIDO"
    df.loc[df["ApproxSLA"].notna() & horas.between(0, 4, inclusive="both"), "SLAStatus"] = "EM RISCO"

    df["AWBValida"] = df["AWB"].str.len().ge(8)
    return df


def consolidar_awbs(bruto: pd.DataFrame, data_referencia: pd.Timestamp) -> pd.DataFrame:
    valido = bruto[bruto["AWBCompleta"].ne("")].copy()
    valido = valido.sort_values(
        ["AWBCompleta", "PrioridadeStatus", "ExecutionDateTime"],
        ascending=[True, True, False],
        na_position="last",
    )

    selecionado = valido.drop_duplicates("AWBCompleta", keep="first").copy()

    agrupado = valido.groupby("AWBCompleta", dropna=False)
    qtd_linhas = agrupado.size().rename("LinhasNoRelatorio")
    qtd_status = agrupado["StatusOperacional"].nunique().rename("QuantidadeStatus")
    lista_status = agrupado["StatusOperacional"].agg(
        lambda s: " | ".join(sorted(set(str(v) for v in s if pd.notna(v))))
    ).rename("StatusEncontrados")

    selecionado = selecionado.merge(qtd_linhas, on="AWBCompleta", how="left")
    selecionado = selecionado.merge(qtd_status, on="AWBCompleta", how="left")
    selecionado = selecionado.merge(lista_status, on="AWBCompleta", how="left")

    selecionado["Pendente"] = selecionado["StatusOperacional"].isin(STATUS_PENDENCIAS)
    selecionado["SLAStatus"] = selecionado["SLAStatus"].where(
        selecionado["Pendente"], "NÃO APLICÁVEL"
    )
    selecionado["HorasAtraso"] = 0.0
    mascara_atraso = selecionado["Pendente"] & selecionado["ApproxSLA"].notna() & (
        data_referencia > selecionado["ApproxSLA"]
    )
    selecionado.loc[mascara_atraso, "HorasAtraso"] = (
        (data_referencia - selecionado.loc[mascara_atraso, "ApproxSLA"]).dt.total_seconds()
        / 3600
    ).round(1)

    selecionado["Inconsistencia"] = ""
    selecionado.loc[selecionado["Cliente"].eq("OUTROS / NÃO MAPEADO"), "Inconsistencia"] += "CLIENTE NÃO MAPEADO; "
    selecionado.loc[selecionado["ApproxSLA"].isna() & selecionado["Pendente"], "Inconsistencia"] += "SEM SLA; "
    selecionado.loc[selecionado["OPSStation"].isna() | selecionado["OPSStation"].astype(str).str.strip().eq(""), "Inconsistencia"] += "SEM BASE OPS; "
    selecionado.loc[selecionado["QuantidadeStatus"].gt(1), "Inconsistencia"] += "MÚLTIPLOS STATUS; "
    selecionado["Inconsistencia"] = selecionado["Inconsistencia"].str.rstrip("; ")

    return selecionado.sort_values(
        ["Pendente", "PrioridadeStatus", "HorasAtraso"],
        ascending=[False, True, False],
    ).reset_index(drop=True)


def processar_status_awb(
    arquivos: Sequence,
    nomes_arquivos: Sequence[str] | None = None,
    data_referencia: datetime | pd.Timestamp | None = None,
) -> ResultadoProcessamento:
    if not arquivos:
        raise ValueError("Anexe pelo menos um relatório AWBStatus.")

    referencia = pd.Timestamp(data_referencia or datetime.now())
    nomes = list(nomes_arquivos or [getattr(a, "name", f"arquivo_{i+1}.xlsx") for i, a in enumerate(arquivos)])

    partes: list[pd.DataFrame] = []
    metadados: list[dict] = []
    for arquivo, nome in zip(arquivos, nomes):
        parte = preparar_arquivo_status(arquivo, nome, referencia)
        partes.append(parte)
        origens = ", ".join(sorted(parte["OriginCode"].dropna().astype(str).unique()))
        metadados.append(
            {
                "Arquivo": nome,
                "Linhas": len(parte),
                "AWBs": parte["AWBCompleta"].nunique(),
                "Origens": origens,
                "Situação": "PROCESSADO",
            }
        )

    bruto = pd.concat(partes, ignore_index=True, sort=False)
    consolidado = consolidar_awbs(bruto, referencia)
    pendencias = consolidado[consolidado["Pendente"]].copy()
    inconsistencias = consolidado[consolidado["Inconsistencia"].ne("")].copy()

    return ResultadoProcessamento(
        bruto=bruto,
        consolidado=consolidado,
        pendencias=pendencias,
        inconsistencias=inconsistencias,
        arquivos=pd.DataFrame(metadados),
        data_referencia=referencia,
    )


def ler_relatorio_complementar(arquivo, nome: str) -> tuple[pd.DataFrame, str]:
    df = _ler_excel(arquivo)
    df.columns = [str(c).strip() for c in df.columns]
    aliases_awb = {
        "AWB",
        "AWB NUMBER",
        "NUMERO AWB",
        "NÚMERO AWB",
        "CONHECIMENTO",
        "CTE",
        "CT-E",
    }
    encontrada = next((c for c in df.columns if normalizar_texto(c) in {normalizar_texto(a) for a in aliases_awb}), None)
    situacao = "RECEBIDO — AWB IDENTIFICADA" if encontrada else "RECEBIDO — AGUARDANDO MAPEAMENTO DE COLUNAS"
    return df, situacao


def aplicar_filtros(
    df: pd.DataFrame,
    clientes: Iterable[str] | None = None,
    operacoes: Iterable[str] | None = None,
    statuses: Iterable[str] | None = None,
    bases: Iterable[str] | None = None,
    destinos: Iterable[str] | None = None,
    sla: Iterable[str] | None = None,
) -> pd.DataFrame:
    filtrado = df.copy()
    filtros = [
        ("Cliente", clientes),
        ("Operacao", operacoes),
        ("StatusOperacional", statuses),
        ("OPSStation", bases),
        ("DestinationCode", destinos),
        ("SLAStatus", sla),
    ]
    for coluna, valores in filtros:
        valores = list(valores or [])
        if valores:
            filtrado = filtrado[filtrado[coluna].isin(valores)]
    return filtrado


def resumo_status(df: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    for status in STATUS_PENDENCIAS:
        fatia = df[df["StatusOperacional"].eq(status)]
        linhas.append(
            {
                "Status": status,
                "AWBs": fatia["AWBCompleta"].nunique(),
                "SLA Vencido": fatia["SLAStatus"].eq("VENCIDO").sum(),
                "Em Risco": fatia["SLAStatus"].eq("EM RISCO").sum(),
                "Sem SLA": fatia["SLAStatus"].eq("SEM SLA").sum(),
            }
        )
    return pd.DataFrame(linhas)


def resumo_clientes(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Cliente", *STATUS_PENDENCIAS, "Total Pendências", "SLA Vencido"])
    tabela = pd.crosstab(df["Cliente"], df["StatusOperacional"])
    for status in STATUS_PENDENCIAS:
        if status not in tabela.columns:
            tabela[status] = 0
    tabela = tabela[STATUS_PENDENCIAS]
    tabela["Total Pendências"] = tabela.sum(axis=1)
    vencido = df.assign(_v=df["SLAStatus"].eq("VENCIDO")).groupby("Cliente")["_v"].sum()
    tabela["SLA Vencido"] = vencido
    return tabela.reset_index().sort_values("Total Pendências", ascending=False)


def resumo_bases(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Base", "Total Pendências", "SLA Vencido", "% SLA Vencido", "Clientes Impactados"])
    base = df.copy()
    base["OPSStation"] = base["OPSStation"].fillna("SEM BASE").astype(str)
    tabela = pd.crosstab(base["OPSStation"], base["StatusOperacional"])
    for status in STATUS_PENDENCIAS:
        if status not in tabela.columns:
            tabela[status] = 0
    tabela = tabela[STATUS_PENDENCIAS]
    tabela["Total Pendências"] = tabela.sum(axis=1)
    tabela["SLA Vencido"] = base.assign(_v=base["SLAStatus"].eq("VENCIDO")).groupby("OPSStation")["_v"].sum()
    tabela["% SLA Vencido"] = (tabela["SLA Vencido"] / tabela["Total Pendências"]).fillna(0)
    tabela["Clientes Impactados"] = base.groupby("OPSStation")["Cliente"].nunique()
    return tabela.reset_index(names="Base").sort_values(
        ["SLA Vencido", "Total Pendências"], ascending=False
    )
