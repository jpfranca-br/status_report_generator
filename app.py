"""App web local para gerar o status report one-page.

Rodar:  streamlit run app.py
"""
import io
import tempfile
from datetime import date
from pathlib import Path

import streamlit as st

import criar_modelo
import gerador

MAX_ARQUIVOS = 5  # planilhas por envio

st.set_page_config(page_title="Status report one-page", page_icon="📊", layout="wide")
st.title("Status report one-page")
st.caption("Suba a planilha preenchida e baixe o slide em PowerPoint.")

with st.sidebar:
    st.subheader("Planilha modelo")
    with tempfile.TemporaryDirectory() as tmp:
        caminho = criar_modelo.criar(str(Path(tmp) / "modelo_status_report.xlsx"))
        modelo = Path(caminho).read_bytes()
    st.download_button("Baixar modelo (.xlsx)", modelo, "modelo_status_report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown(
        "A aba **Instruções** explica cada campo. Status do cronograma: "
        "`concluido`, `realizado`, `sem_sucesso`, `critico`, `planejado`, `janela`."
    )
    data_ref = st.date_input("Data da linha 'Hoje'", value=date.today(), format="DD/MM/YYYY",
                             help="Usada quando a planilha não define data_referencia.")

arquivos = st.file_uploader(f"Planilhas (.xlsx), até {MAX_ARQUIVOS}. Mais de uma = um slide por planilha, no mesmo arquivo.",
                            type=["xlsx"], accept_multiple_files=True)
if not arquivos:
    st.info("Aguardando planilha.")
    st.stop()
if len(arquivos) > MAX_ARQUIVOS:
    st.error(f"Envie no máximo {MAX_ARQUIVOS} planilhas por vez.")
    st.stop()

relatorios, erros = [], []
for a in arquivos:
    try:
        dados = a.getvalue()
        gerador.checar_xlsx(dados)
        r = gerador.ler_planilha(io.BytesIO(dados), a.name)
        r.projeto["data_referencia"] = r.projeto.get("data_referencia") or data_ref
        relatorios.append(r)
    except Exception as e:  # planilha fora do formato
        erros.append(f"{a.name}: {e}")
for e in erros:
    st.error(e)
if not relatorios:
    st.stop()

with tempfile.TemporaryDirectory() as tmp:
    saida = Path(tmp) / "status_report.pptx"
    avisos = gerador.gerar(relatorios, saida)
    pptx_bytes = saida.read_bytes()
    pngs = None
    try:
        pngs = gerador.previa_png(saida)
    except Exception as e:
        st.warning(f"Não consegui gerar a prévia: {e}")

nome = Path(arquivos[0].name).stem if len(arquivos) == 1 else "status_reports"
st.download_button("⬇️ Baixar PowerPoint", pptx_bytes, f"{nome}.pptx", type="primary",
                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if avisos:
    with st.expander(f"⚠️ {len(avisos)} aviso(s) sobre o conteúdo", expanded=True):
        for a in avisos:
            st.write("• " + a)
else:
    st.success("Tudo coube no slide.")

if pngs:
    for i, png in enumerate(pngs, start=1):
        st.image(png, caption=f"Slide {i}", use_container_width=True)
else:
    st.caption("Prévia indisponível: instale o LibreOffice para ver o slide aqui. "
               "O arquivo PowerPoint é gerado normalmente.")

with st.expander("Dados lidos"):
    for r in relatorios:
        st.markdown(f"**{r.projeto.get('nome', r.origem)}**")
        st.dataframe([e.__dict__ for e in r.cronograma], use_container_width=True)
