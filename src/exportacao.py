from __future__ import annotations

from io import BytesIO
import pandas as pd

from .processamento import resumo_bases, resumo_clientes, resumo_status


COLUNAS_DETALHE = [
    "Cliente",
    "Operacao",
    "AWBPrefix",
    "AWB",
    "AWBCompleta",
    "OriginCode",
    "DestinationCode",
    "OPSStation",
    "StatusOperacional",
    "StatusDescription",
    "ExecutionDateTime",
    "ApproxSLA",
    "SLAStatus",
    "HorasParaSLA",
    "HorasAtraso",
    "FltNo",
    "FltDt",
    "FltOrigin",
    "FltDestination",
    "PiecesCount",
    "No of Pieces",
    "GrossWt",
    "QuantidadeStatus",
    "StatusEncontrados",
    "Inconsistencia",
    "ArquivoOrigem",
]


def _colunas_existentes(df: pd.DataFrame, colunas: list[str]) -> list[str]:
    return [c for c in colunas if c in df.columns]


def gerar_excel(consolidado: pd.DataFrame, pendencias: pd.DataFrame, inconsistencias: pd.DataFrame) -> bytes:
    saida = BytesIO()
    resumo_s = resumo_status(pendencias)
    resumo_c = resumo_clientes(pendencias)
    resumo_b = resumo_bases(pendencias)

    with pd.ExcelWriter(saida, engine="xlsxwriter", datetime_format="dd/mm/yyyy hh:mm") as writer:
        workbook = writer.book
        navy = "#17365D"
        white = "#FFFFFF"
        gray = "#D9E1F2"
        colors = {
            "Pendente Embarque": "#F4B183",
            "Pendente Entrega": "#FFD966",
            "Pendente Desembarque": "#D9EAF7",
            "Missing Cargo": "#F4CCCC",
            "Discrepância Criada": "#E4DFEC",
        }
        title = workbook.add_format({"bold": True, "font_color": white, "bg_color": navy, "font_size": 18, "align": "center", "valign": "vcenter"})
        header = workbook.add_format({"bold": True, "font_color": white, "bg_color": navy, "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        number = workbook.add_format({"num_format": "#,##0", "align": "center"})
        percent = workbook.add_format({"num_format": "0.0%", "align": "center"})

        sheet = workbook.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = sheet
        sheet.merge_range("A1:N2", "CARGAS SAO12 — DASHBOARD DE CONTROLE", title)
        sheet.set_row(0, 28)

        for idx, status in enumerate(colors):
            valor = int((pendencias["StatusOperacional"] == status).sum())
            col = idx * 2
            fmt = workbook.add_format({"bold": True, "bg_color": colors[status], "font_color": navy, "font_size": 13, "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1})
            sheet.merge_range(4, col, 7, min(col + 1, 13), f"{status.upper()}\n{valor:,}".replace(",", "."), fmt)

        resumo_s.to_excel(writer, sheet_name="Resumo_Status", index=False)
        resumo_c.to_excel(writer, sheet_name="Resumo_Clientes", index=False)
        resumo_b.to_excel(writer, sheet_name="Bases_Ofensoras", index=False)
        pendencias[_colunas_existentes(pendencias, COLUNAS_DETALHE)].to_excel(writer, sheet_name="Pendencias", index=False)
        consolidado[_colunas_existentes(consolidado, COLUNAS_DETALHE)].to_excel(writer, sheet_name="Consolidado", index=False)
        inconsistencias[_colunas_existentes(inconsistencias, COLUNAS_DETALHE)].to_excel(writer, sheet_name="Inconsistencias", index=False)

        for nome, df in {
            "Resumo_Status": resumo_s,
            "Resumo_Clientes": resumo_c,
            "Bases_Ofensoras": resumo_b,
            "Pendencias": pendencias[_colunas_existentes(pendencias, COLUNAS_DETALHE)],
            "Consolidado": consolidado[_colunas_existentes(consolidado, COLUNAS_DETALHE)],
            "Inconsistencias": inconsistencias[_colunas_existentes(inconsistencias, COLUNAS_DETALHE)],
        }.items():
            ws = writer.sheets[nome]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))
            ws.set_row(0, 30, header)
            for i, coluna in enumerate(df.columns):
                largura = min(max(len(str(coluna)) + 2, 12), 34)
                if coluna in {"StatusEncontrados", "Inconsistencia", "StatusDescription"}:
                    largura = 32
                ws.set_column(i, i, largura)
            if nome == "Bases_Ofensoras" and "% SLA Vencido" in df.columns:
                c = df.columns.get_loc("% SLA Vencido")
                ws.set_column(c, c, 15, percent)

    return saida.getvalue()
