import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
import os
import copy
from docx.table import _Row

st.set_page_config(page_title="Gerador de Documentos SST", layout="wide")

EXCEL_PATH = "banco_aprs.xlsx"

# -----------------------------
# MAPEAMENTO DOS TEMPLATES POR EMPRESA
# -----------------------------
TEMPLATES_EMPRESAS = {
    "Benteler": "T.SHE.046 Safe Job Analyses Bentler.docx",
    "Macromaq": "T.SHE.046 Safe Job Analyses Macromaq.docx"
}

# -----------------------------
# LEITURA DO EXCEL
# -----------------------------
@st.cache_data
def carregar_dados():
    if os.path.exists(EXCEL_PATH):
        df = pd.read_excel(EXCEL_PATH)
        df.columns = df.columns.str.lower().str.strip()
        df.rename(columns={
            'ações': 'acoes', 'acao': 'acoes',
            'atividade': 'atividade', 'tarefa': 'tarefa',
            'risco': 'risco', 'área': 'area'
        }, inplace=True)
        return df
    return pd.DataFrame()

df_excel = carregar_dados()
LISTA_ATIVIDADES = df_excel["atividade"].dropna().unique().tolist() if not df_excel.empty else []


# -----------------------------
# FUNÇÃO AUXILIAR DE SUBSTITUIÇÃO
# -----------------------------
def substituir_texto_paragrafo(p, mapeamento):
    texto_completo = "".join([run.text for run in p.runs])
    houve_mudanca = False
    
    for tag, valor in mapeamento.items():
        if tag in texto_completo:
            texto_completo = texto_completo.replace(tag, str(valor))
            houve_mudanca = True
            
    if houve_mudanca:
        if p.runs:
            p.runs[0].text = texto_completo
            for run in p.runs[1:]:
                p._p.remove(run._r)
        else:
            p.add_run(texto_completo)


# -----------------------------
# PROCESSAMENTO DO DOCX
# -----------------------------
def substituir_docx(doc, mapeamento, passos_atividade):
    for p in doc.paragraphs:
        substituir_texto_paragrafo(p, mapeamento)

    for table in doc.tables:
        linha_modelo = None
        idx_tarefa, idx_risco, idx_acao = None, None, None
        
        for row in table.rows:
            encontrou_na_linha = False
            for idx, cell in enumerate(row.cells):
                texto_celula = cell.text.upper()
                if "{{TAREFAS}}" in texto_celula:
                    linha_modelo, idx_tarefa = row, idx
                    encontrou_na_linha = True
                if "{{RISCOS}}" in texto_celula:
                    linha_modelo, idx_risco = row, idx
                    encontrou_na_linha = True
                if "{{ACOES}}" in texto_celula:
                    linha_modelo, idx_acao = row, idx
                    encontrou_na_linha = True
            if encontrou_na_linha:
                break
        
        if linha_modelo is not None:
            passos_lista = passos_atividade.to_dict(orient="records")
            linha_referencia = linha_modelo
            
            for i, item in enumerate(passos_lista):
                tarefa = "" if pd.isna(item.get("tarefa")) else str(item["tarefa"]).strip()
                risco = "" if pd.isna(item.get("risco")) else str(item["risco"]).strip()
                acao = "" if pd.isna(item.get("acoes")) else str(item["acoes"]).strip()

                if i == 0:
                    nova_linha = linha_modelo
                else:
                    tr_copiado = copy.deepcopy(linha_modelo._tr)
                    linha_referencia._tr.addnext(tr_copiado)
                    nova_linha = _Row(tr_copiado, table)
                    linha_referencia = nova_linha
                
                def preencher_celula_limpa(celula, texto):
                    if not celula.paragraphs:
                        celula.add_paragraph()
                    p = celula.paragraphs[0]
                    p.text = ""
                    for extra_p in celula.paragraphs[1:]:
                        p_xml = extra_p._p
                        p_xml.getparent().remove(p_xml)
                    
                    run = p.add_run(texto)
                    run.font.size = Pt(7)
                
                if idx_tarefa is not None:
                    preencher_celula_limpa(nova_linha.cells[idx_tarefa], f"{i+1}. {tarefa}" if tarefa else "")
                if idx_risco is not None:
                    preencher_celula_limpa(nova_linha.cells[idx_risco], f"• {risco}" if risco else "")
                if idx_acao is not None:
                    preencher_celula_limpa(nova_linha.cells[idx_acao], f"• {acao}" if acao else "")
                    
        else:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        substituir_texto_paragrafo(p, mapeamento)


# -----------------------------
# GERADOR PRINCIPAL
# -----------------------------
def gerar_documento(dados, atividade_selecionada, passos, caminho_template):
    if not os.path.exists(caminho_template):
        st.error(f"Template não encontrado no caminho: {caminho_template}")
        return None

    doc = Document(caminho_template)

    mapeamento = {
        "{{CONTRATADA}}": dados["contratada"].upper(),
        "{{LOCAL}}": dados["local"].upper(),
        "{{AREA}}": dados["area"].upper(),
        "{{ÁREA}}": dados["area"].upper(),
        "{{ATIVIDADE}}": atividade_selecionada.upper(),
        "{{PROCESSO}}": dados["processo"].upper()
    }

    substituir_docx(doc, mapeamento, passos)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# INTERFACE STREAMLIT
# -----------------------------
st.title("🛠️ Gerador ATS - SSMA")

st.subheader("📋 Configuração do Template")
empresa_selecionada = st.selectbox("Selecione a Empresa cliente (Template):", list(TEMPLATES_EMPRESAS.keys()))
caminho_template_escolhido = TEMPLATES_EMPRESAS[empresa_selecionada]

st.divider()

col1, col2 = st.columns(2)

with col1:
    atividade = st.selectbox("Atividade", LISTA_ATIVIDADES)
    processo = st.text_input("Processo", "MANUTENÇÃO MECÂNICA")

with col2:
    contratada = st.text_input("Contratada", empresa_selecionada.upper())
    local = st.text_input("Local", "OFICINA")
    area = st.text_input("Área", "EMPILHADEIRA")

passos = df_excel[df_excel["atividade"] == atividade] if not df_excel.empty else pd.DataFrame()

st.subheader("Prévia dos dados do Excel")
st.dataframe(passos, use_container_width=True)

if st.button("🚀 Gerar Documento", type="primary"):
    if df_excel.empty:
        st.error("O banco de dados do Excel não foi carregado corretamente.")
    elif passos.empty:
        st.warning("Nenhum passo encontrado para a atividade selecionada.")
    else:
        dados = {"contratada": contratada, "local": local, "area": area, "processo": processo}
        
        resultado = gerar_documento(dados, atividade, passos, caminho_template_escolhido)

        if resultado:
            st.success(f"Documento gerado com sucesso utilizando o template da {empresa_selecionada}!")
            st.download_button(
                "📥 Baixar DOCX",
                resultado,
                file_name=f"ATS_{empresa_selecionada}_{atividade.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
