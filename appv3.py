import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table
import difflib
import pandas as pd

# Configuração da página em modo Wide
st.set_page_config(page_title="Comparador Word Avançado", page_icon="📝", layout="wide")

st.title("📝 Comparador Word Avançado (Texto e Tabelas Lado a Lado)")
st.write("Esta versão deteta parágrafos e tabelas na ordem correta e alinha as alterações correspondentes.")

# Função para iterar sobre os elementos do documento mantendo a ordem real (texto/tabela)
def iterar_elementos_documento(documento):
    """Gera parágrafos e tabelas na ordem exata em que aparecem no documento."""
    for nó in documento.element.body:
        if nó.tag.endswith('p'):
            yield Paragraph(nó, documento)
        elif nó.tag.endswith('tbl'):
            yield Table(nó, documento)

# Converter uma tabela do Word numa string/lista estruturada para o difflib conseguir comparar
def converter_tabela_para_dados(tabela):
    dados = []
    for linha in tabela.rows:
        dados.append([celula.text.strip() for celula in linha.cells])
    return dados

# Extrai e estrutura os blocos do documento
def processar_documento(ficheiro_docx):
    try:
        doc = Document(ficheiro_docx)
        elementos_estruturados = []
        
        for elem in iterar_elementos_documento(doc):
            if isinstance(elem, Paragraph):
                if elem.text.strip() != "":
                    elementos_estruturados.append({"tipo": "texto", "conteudo": elem.text.strip()})
            elif isinstance(elem, Table):
                dados_tabela = converter_tabela_para_dados(elem)
                if dados_tabela:
                    elementos_estruturados.append({"tipo": "tabela", "conteudo": dados_tabela})
                    
        return elementos_estruturados
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro: {e}")
        return []

# Helper para renderizar tabelas estilizadas
def mostrar_tabela_estilizada(dados_tabela, estilo_borda):
    df = pd.DataFrame(dados_tabela)
    # Renderiza como HTML puro para permitir aplicar cores de fundo personalizadas nas bordas
    html = df.to_html(header=False, index=False, classes="table-design")
    st.markdown(f"<div style='{estilo_borda} overflow-x: auto; padding: 5px;'>{html}</div>", unsafe_allow_html=True)

# Upload dos ficheiros
col_up1, col_up2 = st.columns(2)
with col_up1:
    file_a = st.file_uploader("📄 Versão Original (A)", type=["docx"], key="file_a")
with col_up2:
    file_b = st.file_uploader("📄 Versão Modificada (B)", type=["docx"], key="file_b")

st.markdown("---")

if file_a and file_b:
    doc_a = processar_documento(file_a)
    doc_b = processar_documento(file_b)
    
    # Criamos chaves textuais simples para o SequenceMatcher alinhar os blocos estruturais
    chaves_a = [f"[{item['tipo']}] {str(item['conteudo'])}" for item in doc_a]
    chaves_b = [f"[{item['tipo']}] {str(item['conteudo'])}" for item in doc_b]
    
    matcher = difflib.SequenceMatcher(None, chaves_a, chaves_b)
    
    # Cabeçalhos das colunas
    col_visual_a, col_visual_b = st.columns(2)
    with col_visual_a:
        st.markdown("### 🔴 Original")
    with col_visual_b:
        st.markdown("### 🟢 Modificado")
        
    # Estilos CSS
    estilo_removido = "color:#333; background-color:#ffcdd2; padding:8px; margin:4px 0; border-radius:4px; border-left: 5px solid #ef5350;"
    estilo_adicionado = "color:#333; background-color:#c8e6c9; padding:8px; margin:4px 0; border-radius:4px; border-left: 5px solid #66bb6a;"
    estilo_igual = "color:#555; background-color:#f9f9f9; padding:8px; margin:4px 0; border-radius:4px; border-left: 5px solid #ccc;"

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        
        if tag == 'equal':
            for item_a, item_b in zip(doc_a[i1:i2], doc_b[j1:j2]):
                c1, c2 = st.columns(2)
                if item_a["tipo"] == "texto":
                    c1.markdown(f"<div style='{estilo_igual}'>{item_a['conteudo']}</div>", unsafe_allow_html=True)
                    c2.markdown(f"<div style='{estilo_igual}'>{item_b['conteudo']}</div>", unsafe_allow_html=True)
                else:
                    with c1: mostrar_tabela_estilizada(item_a['conteudo'], estilo_igual)
                    with c2: mostrar_tabela_estilizada(item_b['conteudo'], estilo_igual)
                    
        elif tag == 'replace':
            lista_a = doc_a[i1:i2]
            lista_b = doc_b[j1:j2]
            max_len = max(len(lista_a), len(lista_b))
            for idx in range(max_len):
                c1, c2 = st.columns(2)
                if idx < len(lista_a):
                    item = lista_a[idx]
                    if item["tipo"] == "texto":
                        c1.markdown(f"<div style='{estilo_removido}'><strong>[Texto Alterado]</strong> {item['conteudo']}</div>", unsafe_allow_html=True)
                    else:
                        with c1: 
                            st.caption("🚨 **Tabela Alterada/Removida:**")
                            mostrar_tabela_estilizada(item['conteudo'], estilo_removido)
                if idx < len(lista_b):
                    item = lista_b[idx]
                    if item["tipo"] == "texto":
                        c2.markdown(f"<div style='{estilo_adicionado}'><strong>[Texto Novo]</strong> {item['conteudo']}</div>", unsafe_allow_html=True)
                    else:
                        with c2:
                            st.caption("✅ **Tabela Nova/Modificada:**")
                            mostrar_tabela_estilizada(item['conteudo'], estilo_adicionado)
                            
        elif tag == 'delete':
            for item in doc_a[i1:i2]:
                c1, c2 = st.columns(2)
                if item["tipo"] == "texto":
                    c1.markdown(f"<div style='{estilo_removido}'><strong>[Apagado]</strong> {item['conteudo']}</div>", unsafe_allow_html=True)
                else:
                    with c1:
                        st.caption("🚨 **Tabela Removida:**")
                        mostrar_tabela_estilizada(item['conteudo'], estilo_removido)
                c2.write("")
                
        elif tag == 'insert':
            for item in doc_b[j1:j2]:
                c1, c2 = st.columns(2)
                c1.write("")
                if item["tipo"] == "texto":
                    c2.markdown(f"<div style='{estilo_adicionado}'><strong>[Inserido]</strong> {item['conteudo']}</div>", unsafe_allow_html=True)
                else:
                    with c2:
                        st.caption("✅ **Tabela Inserida:**")
                        mostrar_tabela_estilizada(item['conteudo'], estilo_adicionado)
else:
    st.info("💡 Insira os documentos com tabelas para testar o alinhamento avançado.")
