from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

from src.exportacao import gerar_excel
from src.processamento import (
    STATUS_PENDENCIAS,
    aplicar_filtros,
    ler_relatorio_complementar,
    processar_status_awb,
    resumo_bases,
    resumo_clientes,
    resumo_status,
)


st.set_page_config(
    page_title="CARGAS SAO12",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

STATUS_CORES = {
    "Pendente Embarque": "#F4B183",
    "Pendente Entrega": "#FFD966",
    "Pendente Desembarque": "#9DC3E6",
    "Missing Cargo": "#E06666",
    "Discrepância Criada": "#A64D79",
}

st.markdown(
    """
    <style>
      .block-container {padding-top: 1.1rem; padding-bottom: 2rem;}
      [data-testid="stSidebar"] {background: #F4F6F9;}
      .app-title {
        background: linear-gradient(90deg, #17365D, #2F75B5);
        color: white; padding: 18px 24px; border-radius: 12px;
        margin-bottom: 12px; box-shadow: 0 2px 8px rgba(0,0,0,.12);
      }
      .app-title h1 {margin: 0; font-size: 30px;}
      .app-title p {margin: 4px 0 0 0; opacity: .92;}
      .kpi-card {
        min-height: 128px; padding: 16px 12px; border-radius: 12px;
        text-align: center; border: 1px solid rgba(23,54,93,.18);
        box-shadow: 0 2px 7px rgba(0,0,0,.08);
        display:flex; flex-direction:column; justify-content:center;
      }
      .kpi-title {font-size: 13px; font-weight: 750; color: #17365D; text-transform: uppercase;}
      .kpi-value {font-size: 34px; line-height: 1.05; font-weight: 800; color: #17365D; margin-top: 8px;}
      .kpi-sub {font-size: 12px; color:#595959; margin-top:6px;}
      .small-card {background:#17365D; color:white; min-height:95px;}
      .small-card .kpi-title, .small-card .kpi-value, .small-card .kpi-sub {color:white;}
      .section-title {font-size:20px; font-weight:750; color:#17365D; margin-top:6px;}
      .status-ok {color:#2E7D32; font-weight:700;}
      .status-warn {color:#B26A00; font-weight:700;}
    </style>
    """,
    unsafe_allow_html=True,
)


def card(titulo: str, valor: int | float | str, cor: str, subtitulo: str = "") -> None:
    if isinstance(valor, (int, float)):
        valor_fmt = f"{valor:,.0f}".replace(",", ".")
    else:
        valor_fmt = str(valor)
    st.markdown(
        f"""
        <div class="kpi-card" style="background:{cor}">
            <div class="kpi-title">{titulo}</div>
            <div class="kpi-value">{valor_fmt}</div>
            <div class="kpi-sub">{subtitulo}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def card_secundario(titulo: str, valor: int | str, subtitulo: str = "") -> None:
    valor_fmt = f"{valor:,.0f}".replace(",", ".") if isinstance(valor, (int, float)) else valor
    st.markdown(
        f"""
        <div class="kpi-card small-card">
            <div class="kpi-title">{titulo}</div>
            <div class="kpi-value">{valor_fmt}</div>
            <div class="kpi-sub">{subtitulo}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def opcoes(df: pd.DataFrame, coluna: str) -> list[str]:
    if coluna not in df.columns:
        return []
    return sorted(v for v in df[coluna].dropna().astype(str).unique() if v.strip())


st.markdown(
    """
    <div class="app-title">
      <h1>CARGAS SAO12</h1>
      <p>Controle de pendências de embarque, desembarque, entrega, missing cargo, discrepâncias e SLA.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Entrada dos relatórios")
    arquivos_status = st.file_uploader(
        "AWBStatus — SAO12 e TRES1",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="Anexe um ou mais arquivos no mesmo modelo AWBStatusAtPieceLevel.",
    )
    arquivo_embarque = st.file_uploader(
        "Franchise",
        type=["xlsx", "xlsm"],
        help="Anexe o relatório Commission Report Franchise.",
    )
    arquivo_entrega = st.file_uploader(
        "Notas Integradas (EDI)",
        type=["xlsx", "xlsm"],
        help="Anexe o relatório de Notas Integradas (EDI).",
    )
    horas_risco = st.number_input(
        "Alertar SLA em risco quando faltarem até (horas)",
        min_value=1,
        max_value=24,
        value=10,
        help=(
            "Exemplo: com 10 horas, toda carga pendente cujo SLA vence entre agora "
            "e as próximas 10 horas será classificada como EM RISCO. "
            "Se o prazo já venceu, será VENCIDO; se faltarem mais de 10 horas, ficará NO PRAZO."
        ),
    )
    processar = st.button("PROCESSAR RELATÓRIOS", type="primary", use_container_width=True)

    st.divider()
    st.caption(
        "O painel principal usa o AWBStatus. Os relatórios Franchise e Notas Integradas (EDI) "
        "ficam disponíveis para o cruzamento operacional."
    )

if processar:
    if not arquivos_status:
        st.error("Anexe pelo menos um relatório AWBStatus.")
    else:
        try:
            with st.spinner("Validando, consolidando AWBs e calculando SLA..."):
                resultado = processar_status_awb(
                    arquivos_status,
                    [a.name for a in arquivos_status],
                    datetime.now(),
                )
                # Ajuste configurável da faixa de risco.
                consolidado = resultado.consolidado.copy()
                pendente_com_sla = consolidado["Pendente"] & consolidado["ApproxSLA"].notna()
                consolidado.loc[pendente_com_sla, "SLAStatus"] = "NO PRAZO"
                consolidado.loc[pendente_com_sla & consolidado["HorasParaSLA"].lt(0), "SLAStatus"] = "VENCIDO"
                consolidado.loc[
                    pendente_com_sla & consolidado["HorasParaSLA"].between(0, horas_risco, inclusive="both"),
                    "SLAStatus",
                ] = "EM RISCO"
                consolidado.loc[consolidado["Pendente"] & consolidado["ApproxSLA"].isna(), "SLAStatus"] = "SEM SLA"
                resultado = resultado.__class__(
                    bruto=resultado.bruto,
                    consolidado=consolidado,
                    pendencias=consolidado[consolidado["Pendente"]].copy(),
                    inconsistencias=consolidado[consolidado["Inconsistencia"].ne("")].copy(),
                    arquivos=resultado.arquivos,
                    data_referencia=resultado.data_referencia,
                )
                st.session_state["resultado"] = resultado

                complementares = {}
                for chave, arquivo in {"Franchise": arquivo_embarque, "Notas Integradas (EDI)": arquivo_entrega}.items():
                    if arquivo is not None:
                        df_comp, situacao = ler_relatorio_complementar(arquivo, arquivo.name)
                        complementares[chave] = {"nome": arquivo.name, "situacao": situacao, "dados": df_comp}
                st.session_state["complementares"] = complementares
            st.success("Relatórios processados.")
        except Exception as exc:
            st.exception(exc)

if "resultado" not in st.session_state:
    st.info("Anexe o relatório AWBStatus na barra lateral e clique em **PROCESSAR RELATÓRIOS**.")
    st.markdown(
        """
        **O painel ficará disponível com:**
        - cards de Pendente Embarque, Pendente Entrega, Pendente Desembarque, Missing Cargo e Discrepância;
        - filtros por cliente, operação, base, destino, status e SLA;
        - ranking de bases ofensoras;
        - relação analítica das AWBs;
        - exportação em Excel.
        """
    )
    st.stop()

resultado = st.session_state["resultado"]
base = resultado.consolidado.copy()

with st.expander("Filtros do painel", expanded=True):
    f1, f2, f3 = st.columns(3)
    f4, f5, f6 = st.columns(3)
    clientes = f1.multiselect("Cliente", opcoes(base, "Cliente"))
    operacoes = f2.multiselect("Operação", opcoes(base, "Operacao"))
    statuses = f3.multiselect("Status", STATUS_PENDENCIAS, default=STATUS_PENDENCIAS)
    bases = f4.multiselect("Base OPS", opcoes(base[base["Pendente"]], "OPSStation"))
    destinos = f5.multiselect("Destino", opcoes(base[base["Pendente"]], "DestinationCode"))
    sla = f6.multiselect("SLA", ["VENCIDO", "EM RISCO", "NO PRAZO", "SEM SLA"])

filtrado = aplicar_filtros(base, clientes, operacoes, statuses, bases, destinos, sla)
pendencias = filtrado[filtrado["Pendente"]].copy()

abas = st.tabs(["Dashboard", "Pendências", "Bases ofensoras", "Inconsistências", "Relatórios recebidos"])

with abas[0]:
    resumo = resumo_status(pendencias).set_index("Status")
    cols = st.columns(5)
    for coluna, status in zip(cols, STATUS_PENDENCIAS):
        with coluna:
            valor = int(resumo.loc[status, "AWBs"]) if status in resumo.index else 0
            vencido = int(resumo.loc[status, "SLA Vencido"]) if status in resumo.index else 0
            card(status, valor, STATUS_CORES[status], f"{vencido} com SLA vencido")

    st.write("")
    monitorados = pendencias[pendencias["Cliente"].ne("OUTROS / NÃO MAPEADO")]["AWBCompleta"].nunique()
    vencidos = int(pendencias["SLAStatus"].eq("VENCIDO").sum())
    outros = pendencias[pendencias["Cliente"].eq("OUTROS / NÃO MAPEADO")]["AWBCompleta"].nunique()
    operacoes_presentes = ", ".join(sorted(resultado.bruto["Operacao"].dropna().unique()))
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        card_secundario("Clientes monitorados", monitorados)
    with c2:
        card_secundario("SLA vencido", vencidos)
    with c3:
        card_secundario("Outros / não mapeados", outros)
    with c4:
        card_secundario("Operações carregadas", operacoes_presentes or "—")

    st.write("")
    graf1, graf2 = st.columns([1, 1.35])
    with graf1:
        dados_status = resumo_status(pendencias)
        fig = px.bar(
            dados_status,
            x="Status",
            y="AWBs",
            color="Status",
            color_discrete_map=STATUS_CORES,
            text_auto=True,
            title="Pendências por status",
        )
        fig.update_layout(showlegend=False, xaxis_title="", yaxis_title="AWBs", height=410)
        st.plotly_chart(fig, use_container_width=True)

    with graf2:
        bases_resumo = resumo_bases(pendencias).head(15).sort_values("Total Pendências")
        if bases_resumo.empty:
            st.info("Nenhuma base encontrada para os filtros selecionados.")
        else:
            fig = px.bar(
                bases_resumo,
                x="Total Pendências",
                y="Base",
                orientation="h",
                color="SLA Vencido",
                color_continuous_scale="Reds",
                text="Total Pendências",
                title="Top 15 bases ofensoras",
            )
            fig.update_layout(height=410, xaxis_title="AWBs pendentes", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title">Pendências por cliente</div>', unsafe_allow_html=True)
    clientes_resumo = resumo_clientes(pendencias)
    st.dataframe(clientes_resumo, use_container_width=True, hide_index=True)

with abas[1]:
    st.subheader(f"Relação analítica — {len(pendencias):,.0f} AWBs".replace(",", "."))
    colunas = [
        "Cliente", "Operacao", "AWB", "AWBCompleta", "OriginCode", "DestinationCode",
        "OPSStation", "StatusOperacional", "ExecutionDateTime", "ApproxSLA", "SLAStatus",
        "HorasParaSLA", "HorasAtraso", "FltNo", "FltDt", "FltOrigin", "FltDestination",
        "PiecesCount", "GrossWt", "StatusEncontrados", "Inconsistencia", "ArquivoOrigem",
    ]
    colunas = [c for c in colunas if c in pendencias.columns]
    st.dataframe(
        pendencias[colunas],
        use_container_width=True,
        hide_index=True,
        column_config={
            "ExecutionDateTime": st.column_config.DatetimeColumn("Última atualização", format="DD/MM/YYYY HH:mm"),
            "ApproxSLA": st.column_config.DatetimeColumn("SLA", format="DD/MM/YYYY HH:mm"),
            "FltDt": st.column_config.DatetimeColumn("Data do voo", format="DD/MM/YYYY HH:mm"),
            "HorasParaSLA": st.column_config.NumberColumn("Horas para SLA", format="%.1f"),
            "HorasAtraso": st.column_config.NumberColumn("Horas de atraso", format="%.1f"),
        },
    )

    excel = gerar_excel(filtrado, pendencias, filtrado[filtrado["Inconsistencia"].ne("")])
    d1, d2 = st.columns(2)
    d1.download_button(
        "Baixar controle em Excel",
        data=excel,
        file_name=f"CARGAS_SAO12_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    d2.download_button(
        "Baixar pendências em CSV",
        data=pendencias[colunas].to_csv(index=False, sep=";", encoding="utf-8-sig"),
        file_name=f"Pendencias_CARGAS_SAO12_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        use_container_width=True,
    )

with abas[2]:
    st.subheader("Ranking de bases ofensoras")
    base_of = resumo_bases(pendencias)
    base_of_exibicao = base_of.copy()
    if "% SLA Vencido" in base_of_exibicao.columns:
        base_of_exibicao["% SLA Vencido"] = base_of_exibicao["% SLA Vencido"] * 100
    st.dataframe(
        base_of_exibicao,
        use_container_width=True,
        hide_index=True,
        column_config={"% SLA Vencido": st.column_config.NumberColumn("% SLA vencido", format="%.1f%%")},
    )

with abas[3]:
    inc = filtrado[filtrado["Inconsistencia"].ne("")].copy()
    st.subheader(f"Inconsistências identificadas: {len(inc):,.0f}".replace(",", "."))
    st.caption("As AWBs permanecem no painel; a inconsistência apenas sinaliza ausência ou conflito de evidência.")
    col_inc = [
        "Cliente", "Operacao", "AWB", "AWBCompleta", "OPSStation", "StatusOperacional",
        "ApproxSLA", "QuantidadeStatus", "StatusEncontrados", "Inconsistencia", "ArquivoOrigem",
    ]
    col_inc = [c for c in col_inc if c in inc.columns]
    st.dataframe(inc[col_inc], use_container_width=True, hide_index=True)

with abas[4]:
    st.subheader("Arquivos AWBStatus processados")
    st.dataframe(resultado.arquivos, use_container_width=True, hide_index=True)
    origens = set(resultado.bruto["Operacao"].dropna().astype(str))
    o1, o2 = st.columns(2)
    o1.success("SAO12 carregado") if "SAO12" in origens else o1.warning("SAO12 não identificado")
    o2.success("TRES1 carregado") if "TRES1" in origens else o2.warning("TRES1 aguardando relatório")

    st.subheader("Relatórios complementares")
    complementares = st.session_state.get("complementares", {})
    if not complementares:
        st.info("Nenhum relatório complementar foi anexado nesta execução.")
    else:
        for tipo, info in complementares.items():
            with st.expander(f"{tipo}: {info['nome']} — {info['situacao']}"):
                st.dataframe(info["dados"].head(30), use_container_width=True, hide_index=True)

st.caption(
    f"Referência do cálculo: {resultado.data_referencia:%d/%m/%Y %H:%M}. "
    "Contagem consolidada por AWB; quando existem múltiplos status, prevalece o de maior criticidade."
)
