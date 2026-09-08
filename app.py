import streamlit as st
from docx import Document
import difflib

# Configuração da página do Streamlit
st.set_page_config(page_title="Comparador de Word", page_icon="📝", layout="wide")

st.title("📝 Comparador de Ficheiros Word (.docx)")
st.write("Faça o upload de dois ficheiros Word para ver as diferenças detalhadas entre eles.")

# Função para extrair texto de um ficheiro .docx
def extrair_texto_docx(docx_file):
    try:
        doc = Document(docx_file)
        # Junta o texto de todos os parágrafos separados por quebras de linha
        texto = [p.text for p in doc.paragraphs if p.text.strip() != ""]
        return texto
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro: {e}")
        return []

# Criar duas colunas para o upload dos ficheiros
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Ficheiro Original (Versão A)")
    file_a = st.file_uploader("Selecione o primeiro ficheiro", type=["docx"], key="file_a")

with col2:
    st.subheader("📄 Ficheiro Modificado (Versão B)")
    file_b = st.file_uploader("Selecione o segundo ficheiro", type=["docx"], key="file_b")

# Separador visual
st.markdown("---")

# Processar os ficheiros se ambos forem carregados
if file_a and file_b:
    st.subheader("🔍 Diferenças Encontradas")
    
    # Extrair o texto de cada documento (retorna uma lista de parágrafos)
    texto_a = extrair_texto_docx(file_a)
    texto_b = extrair_texto_docx(file_b)
    
    # Gerar a comparação linha por linha (parágrafo por parágrafo)
    diferencas = list(difflib.ndiff(texto_a, texto_b))
    
    tem_diferencas = False
    
    # Contentor para exibir o resultado
    with st.container():
        for linha in diferencas:
            # O difflib adiciona um prefixo de 2 caracteres a cada linha:
            # '- ' significa que existia no A mas foi removido no B
            # '+ ' significa que foi adicionado no B
            # '  ' significa que é igual em ambos
            # '? ' são linhas de apoio do difflib para guiar o olhar (podemos ignorar)
            
            if linha.startswith('- '):
                tem_diferencas = True
                # Texto removido em fundo vermelho
                st.markdown(f"<p style='color:white; background-color:#ef5350; padding:5px; margin:2px; border-radius:3px;'><strong>[-] Removido:</strong> {linha[2:]}</p>", unsafe_allow_html=True)
                
            elif linha.startswith('+ '):
                tem_diferencas = True
                # Texto adicionado em fundo verde
                st.markdown(f"<p style='color:white; background-color:#66bb6a; padding:5px; margin:2px; border-radius:3px;'><strong>[+] Adicionado:</strong> {linha[2:]}</p>", unsafe_allow_html=True)
                
            elif linha.startswith('  '):
                # Texto igual (opcional: pode comentar esta linha se quiser ver APENAS as diferenças)
                st.write(f"ℹ️ {linha[2:]}")
                
        if not tem_diferencas:
            st.success("🎉 Os dois ficheiros são exatamente iguais!")
else:
    st.info("💡 Por favor, faça o upload de ambos os ficheiros acima para iniciar a comparação.")
