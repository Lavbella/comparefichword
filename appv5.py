import streamlit as st
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
import difflib

# Configuração da página em modo Wide
st.set_page_config(page_title="Comparador Word Premium", page_icon="📝", layout="wide")

st.title("📝 Comparador Word Inteligente")
st.write("Compare textos e tabelas lado a lado com filtros avançados de visualização.")

# Função para iterar sobre os elementos mantendo a ordem original do documento
def iterar_elementos_documento(documento):
    for nó in documento.element.body:
        if nó.tag.endswith('p'):
            yield Paragraph(nó, documento)
        elif nó.tag.endswith('tbl'):
            yield Table(nó, documento)

# Converter uma tabela do Word numa matriz de strings (lista de listas)
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

# Compara duas células de texto e gera HTML com os destaques inline e um booleano se há alterações
def comparar_celulas_html(texto_a, texto_b):
    if texto_a == texto_b:
        return f"<span style='color:#555;'>{texto_a}</span>", f"<span style='color:#555;'>{texto_b}</span>", False
    
    palavras_a = texto_a.split()
    palavras_b = texto_b.split()
    
    diff = list(difflib.ndiff(palavras_a, palavras_b))
    
    html_a = []
    html_b = []
    ha_alteracao = False
    
    for item in diff:
        if item.startswith('  '):
            html_a.append(item[2:])
            html_b.append(item[2:])
        elif item.startswith('- '):
            ha_alteracao = True
            html_a.append(f"<span style='background-color:#ffcdd2; color:#b71c1c; font-weight:bold; padding:2px; border-radius:2px;'>{item[2:]}</span>")
        elif item.startswith('+ '):
            ha_alteracao = True
            html_b.append(f"<span style='background-color:#c8e6c9; color:#1b5e20; font-weight:bold; padding:2px; border-radius:2px;'>{item[2:]}</span>")
            
    return " ".join(html_a), " ".join(html_b), ha_alteracao

# Verifica se duas matrizes de tabela têm qualquer diferença
def tabela_tem_diferencas(dados_a, dados_b):
    if dados_a != dados_b:
        return True
    return False

# Converte uma matriz de dados em HTML de tabela
def gerar_html_tabela_comparada(dados_a, dados_b, modo):
    max_linhas = max(len(dados_a), len(dados_b)) if dados_a and dados_b else (len(dados_a) if dados_a else len(dados_b))
    
    html = "<table style='width:100%; border-collapse: collapse; margin: 5px 0; font-family: sans-serif; font-size: 14px;'>"
    
    for r in range(max_linhas):
        html += "<tr>"
        
        linha_a = dados_a[r] if (dados_a and r < len(dados_a)) else []
        linha_b = dados_b[r] if (dados_b and r < len(dados_b)) else []
        max_colunas = max(len(linha_a), len(linha_b))
        
        for c in range(max_colunas):
            celula_a = linha_a[c] if c < len(linha_a) else ""
            celula_b = linha_b[c] if c < len(linha_b) else ""
            
            html_celula_a, html_celula_b, _ = comparar_celulas_html(celula_a, celula_b)
            estilo_td = "border: 1px solid #ddd; padding: 8px; vertical-align: top;"
            
            if modo == "original":
                if celula_a != celula_b and celula_b == "":
                    estilo_td += "background-color: #ffebee;"
                html += f"<td style='{estilo_td}'>{html_celula_a}</td>"
            elif modo == "modificado":
                if celula_a != celula_b and celula_a == "":
                    estilo_td += "background-color: #e8f5e9;"
                html += f"<td style='{estilo_td}'>{html_celula_b}</td>"
                
        html += "</tr>"
    html += "</table>"
    return html

# Painel de upload superior
col_up1, col_up2 = st.columns(2)
with col_up1:
    file_a = st.file_uploader("📄 Versão Original (A)", type=["docx"], key="file_a")
with col_up2:
    file_b = st.file_uploader("📄 Versão Modificada (B)", type=["docx"], key="file_b")

# Filtro de visualização colocado em destaque no topo
st.markdown("### ⚙️ Opções de Visualização")
esconder_iguais = st.checkbox("🔍 Mostrar APENAS o que mudou (Esconder linhas 100% iguais)", value=False)

st.markdown("---")

if file_a and file_b:
    doc_a = processar_documento(file_a)
    doc_b = processar_documento(file_b)
    
    chaves_a = [f"[{item['tipo']}]" for item in doc_a]
    chaves_b = [f"[{item['tipo']}]" for item in doc_b]
    
    matcher = difflib.SequenceMatcher(None, chaves_a, chaves_b)
    
    col_visual_a, col_visual_b = st.columns(2)
    with col_visual_a: st.markdown("### 🔴 Original")
    with col_visual_b: st.markdown("### 🟢 Modificado")
        
    estilo_removido = "color:#333; background-color:#ffcdd2; padding:8px; margin:4px 0; border-radius:4px; border-left: 5px solid #ef5350;"
    estilo_adicionado = "color:#333; background-color:#c8e6c9; padding:8px; margin:4px 0; border-radius:4px; border-left: 5px solid #66bb6a;"
    estilo_igual = "color:#555; background-color:#f9f9f9; padding:8px; margin:4px 0; border-radius:4px; border-left: 5px solid #ccc;"

    total_alteracoes_exibidas = 0

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        
        if tag == 'equal':
            for item_a, item_b in zip(doc_a[i1:i2], doc_b[j1:j2]):
                
                if item_a["tipo"] == "texto":
                    txt_html_a, txt_html_b, mudou = comparar_celulas_html(item_a['conteudo'], item_b['conteudo'])
                    
                    # Filtro ativo: Se o filtro estiver ligado e NÃO mudou nada, salta para o próximo
                    if esconder_iguais and not mudou:
                        continue
                        
                    c1, c2 = st.columns(2)
                    c1.markdown(f"<div style='{estilo_igual}'>{txt_html_a}</div>", unsafe_allow_html=True)
                    c2.markdown(f"<div style='{estilo_igual}'>{txt_html_b}</div>", unsafe_allow_html=True)
                    total_alteracoes_exibidas += 1
                else:
                    mudou = tabela_tem_diferencas(item_a['conteudo'], item_b['conteudo'])
                    
                    if esconder_iguais and not mudou:
                        continue
                        
                    c1, c2 = st.columns(2)
                    html_tab_a = gerar_html_tabela_comparada(item_a['conteudo'], item_b['conteudo'], "original")
                    html_tab_b = gerar_html_tabela_comparada(item_a['conteudo'], item_b['conteudo'], "modificado")
                    c1.markdown(html_tab_a, unsafe_allow_html=True)
                    c2.markdown(html_tab_b, unsafe_allow_html=True)
                    total_alteracoes_exibidas += 1
                    
        elif tag == 'replace':
            lista_a = doc_a[i1:i2]
            lista_b = doc_b[j1:j2]
            max_len = max(len(lista_a), len(lista_b))
            for idx in range(max_len):
                c1, c2 = st.columns(2)
                if idx < len(lista_a):
                    item = lista_a[idx]
                    if item["tipo"] == "texto":
                        c1.markdown(f"<div style='{estilo_removido}'><strong>[Removido]</strong> {item['conteudo']}</div>", unsafe_allow_html=True)
                    else:
                        c1.markdown(gerar_html_tabela_comparada(item['conteudo'], None, "original"), unsafe_allow_html=True)
                if idx < len(lista_b):
                    item = lista_b[idx]
                    if item["tipo"] == "texto":
                        c2.markdown(f"<div style='{estilo_adicionado}'><strong>[Adicionado]</strong> {item['conteudo']}</div>", unsafe_allow_html=True)
                    else:
                        c2.markdown(gerar_html_tabela_comparada(None, item['conteudo'], "modificado"), unsafe_allow_html=True)
                total_alteracoes_exibidas += 1
                            
        elif tag == 'delete':
            for item in doc_a[i1:i2]:
                c1, c2 = st.columns(2)
                if item["tipo"] == "texto":
                    c1.markdown(f"<div style='{estilo_removido}'><strong>[Apagado]</strong> {item['conteudo']}</div>", unsafe_allow_html=True)
                else:
                    c1.markdown(gerar_html_tabela_comparada(item['conteudo'], None, "original"), unsafe_allow_html=True)
                c2.write("")
                total_alteracoes_exibidas += 1
                
        elif tag == 'insert':
            for item in doc_b[j1:j2]:
                c1, c2 = st.columns(2)
                c1.write("")
                if item["tipo"] == "texto":
                    c2.markdown(f"<div style='{estilo_adicionado}'><strong>[Inserido]</strong> {item['conteudo']}</div>", unsafe_allow_html=True)
                else:
                    c2.markdown(gerar_html_tabela_comparada(None, item['conteudo'], "modificado"), unsafe_allow_html=True)
                total_alteracoes_exibidas += 1

    # Mensagem caso o filtro esconda absolutamente tudo
    if esconder_iguais and total_alteracoes_exibidas == 0:
        st.success("🎉 Não há alterações detetadas no documento! Tudo ocultado pelo filtro.")
    else:
        st.info("💡 Insira os documentos para testar o novo filtro de supressão de conteúdo idêntico.")