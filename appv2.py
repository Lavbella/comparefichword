import streamlit as st
from docx import Document
import difflib

# Configuração da página em modo Wide (Largo) para caber as duas colunas
st.set_page_config(page_title="Comparador Word Lado a Lado", page_icon="📝", layout="wide")

st.title("📝 Comparador de Word Lado a Lado")
st.write("Compare as duas versões visualmente alinhadas na mesma zona.")

# Função para extrair texto de um ficheiro .docx
def extrair_texto_docx(docx_file):
    try:
        doc = Document(docx_file)
        return [p.text for p in doc.paragraphs if p.text.strip() != ""]
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro: {e}")
        return []

# Upload dos ficheiros nas colunas superiores
col_up1, col_up2 = st.columns(2)
with col_up1:
    file_a = st.file_uploader("📄 Versão Original (A)", type=["docx"], key="file_a")
with col_up2:
    file_b = st.file_uploader("📄 Versão Modificada (B)", type=["docx"], key="file_b")

st.markdown("---")

if file_a and file_b:
    texto_a = extrair_texto_docx(file_a)
    texto_b = extrair_texto_docx(file_b)
    
    # O SequenceMatcher ajuda-nos a alinhar os blocos correspondentes
    matcher = difflib.SequenceMatcher(None, texto_a, texto_b)
    
    # Títulos das colunas de visualização
    col_visual_a, col_visual_b = st.columns(2)
    with col_visual_a:
        st.markdown("### 🔴 Original")
    with col_visual_b:
        st.markdown("### 🟢 Modificado")
        
    # Estilos CSS reutilizáveis para os blocos de texto
    estilo_removido = "color:#333; background-color:#ffcdd2; padding:8px; margin:4px 0; border-radius:4px; border-left: 5px solid #ef5350;"
    estilo_adicionado = "color:#333; background-color:#c8e6c9; padding:8px; margin:4px 0; border-radius:4px; border-left: 5px solid #66bb6a;"
    estilo_igual = "color:#555; background-color:#f9f9f9; padding:8px; margin:4px 0; border-radius:4px; border-left: 5px solid #ccc;"
    
    # Percorrer as operações de alteração estrutural (equal, replace, delete, insert)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        
        # 1. TEXTO IGUAL: Alinha lado a lado perfeitamente
        if tag == 'equal':
            for a, b in zip(texto_a[i1:i2], texto_b[j1:j2]):
                c1, c2 = st.columns(2)
                c1.markdown(f"<div style='{estilo_igual}'>{a}</div>", unsafe_allow_html=True)
                c2.markdown(f"<div style='{estilo_igual}'>{b}</div>", unsafe_allow_html=True)
                
        # 2. TEXTO SUBSTITUÍDO: Mostra o antigo à esquerda e o novo à direita na mesma linha
        elif tag == 'replace':
            lista_a = texto_a[i1:i2]
            lista_b = texto_b[j1:j2]
            # Sincroniza o máximo de linhas possíveis lado a lado
            max_len = max(len(lista_a), len(lista_b))
            for idx in range(max_len):
                c1, c2 = st.columns(2)
                if idx < len(lista_a):
                    c1.markdown(f"<div style='{estilo_removido}'><strong>[Alterado]</strong> {lista_a[idx]}</div>", unsafe_allow_html=True)
                if idx < len(lista_b):
                    c2.markdown(f"<div style='{estilo_adicionado}'><strong>[Novo]</strong> {lista_b[idx]}</div>", unsafe_allow_html=True)
                    
        # 3. TEXTO APAGADO: Mostra à esquerda, deixa a direita em branco
        elif tag == 'delete':
            for a in texto_a[i1:i2]:
                c1, c2 = st.columns(2)
                c1.markdown(f"<div style='{estilo_removido}'><strong>[Apagado]</strong> {a}</div>", unsafe_allow_html=True)
                c2.write("") # Espaço em branco correspondente
                
        # 4. TEXTO INSERIDO: Deixa a esquerda em branco, mostra à direita
        elif tag == 'insert':
            for b in texto_b[j1:j2]:
                c1, c2 = st.columns(2)
                c1.write("") # Espaço em branco correspondente
                c2.markdown(f"<div style='{estilo_adicionado}'><strong>[Inserido]</strong> {b}</div>", unsafe_allow_html=True)

else:
    st.info("💡 Carregue os dois ficheiros acima para gerar a comparação síncrona.")
