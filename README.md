# 📝 Word File Comparator (v5) / Comparador de Ficheiros Word (v5)

[Português](#português) | [English](#english)

---

## Português

Uma aplicação em Python desenvolvida com **Streamlit** para comparar dois ficheiros Word (`.docx`) lado a lado, destacando visualmente todas as diferenças de texto. Esta versão (v5) é a mais completa e inclui suporte para geração de um executável autónomo (`.exe` / aplicação nativa).

### Funcionalidades
* **Visualização Lado a Lado:** Interface limpa que permite ver os dois documentos em simultâneo.
* **Destaque de Diferenças:** Identificação visual precisa de inserções, remoções e alterações de texto.
* **Suporte a Executável:** Arquitetura preparada para conversão em executável independente através do PyInstaller.

### Estrutura do Projeto
* `appV5.py`: Código principal da aplicação Streamlit (v5).
* `app_spec.py`: Criar o Script de bootstrap/inicialização para o executável.
* `streamlit_app.spec`: Criar o Ficheiro de configuração do PyInstaller.
* `requirements.txt`: Lista de dependências do ecossistema Python.

### Como Executar Localmente
1. Crie e ative o seu ambiente virtual:
   ```bash
   python -m venv env
   # No Windows: .\env\Scripts\activate
   # No macOS/Linux: source env/bin/activate
   ```
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Inicie a aplicação:
   ```bash
   streamlit run appV5.py
   ```

### Como Criar o Executável
Para compilar a aplicação e gerar o executável final utilizando o PyInstaller, execute o seguinte comando na raiz do projeto:
```bash
pyinstaller --clean streamlit_app.spec
```
O executável final será gerado dentro da pasta `dist/`.

### Exemplos de Referência
A estrutura de empacotamento a utilizar neste repositório (configuração do ficheiro `.spec` e o script de bootstrap `app_spec.py` / `run_app.py`) segue o mesmo modelo testado e implementado em **outros projetos públicos disponíveis no meu perfil do GitHub**. Pode consultar esses repositórios para ver exemplos adicionais de como converter aplicações Streamlit em executáveis autónomos.

---

## English

A Python application built with **Streamlit** to compare two Word files (`.docx`) side-by-side, visually highlighting all text differences. This version (v5) is the most complete release and includes full support for generating a standalone executable.

### Features
* **Side-by-Side View:** Clean user interface allowing simultaneous viewing of both documents.
* **Difference Highlighting:** Precise visual identification of text insertions, deletions, and modifications.
* **Executable Ready:** Architecture fully configured for standalone executable conversion using PyInstaller.

### Project Structure
* `appV5.py`: Main Streamlit application code (v5).
* `app_spec.py`: Build a Bootstrap/entry-point script for the executable.
* `streamlit_app.spec`: Build a PyInstaller configuration specification file.
* `requirements.txt`: Python environment dependencies.

### How to Run Locally
1. Create and activate your virtual environment:
   ```bash
   python -m venv env
   # On Windows: .\env\Scripts\activate
   # On macOS/Linux: source env/bin/activate
   ```
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```
3. Launch the application:
   ```bash
   streamlit run appV5.py
   ```

### How to Build the Executable
To compile the application into a standalone executable using PyInstaller, run the following command in the project root:
```bash
pyinstaller --clean streamlit_app.spec
```
The final standalone bundle will be available inside the `dist/` directory.

### Reference Examples
The packaging workflow utilized in this repository (the configuration of the `.spec` file and the `app_spec.py` / `run_app.py` bootstrap script) follows the exact same model tested and deployed across **other public projects available on my GitHub profile**. You can explore those repositories for additional hands-on examples of converting Streamlit applications into standalone executables.

