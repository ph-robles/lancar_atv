import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------
# Configurações do aplicativo
# ---------------------------
ARQUIVO_EXCEL_PADRAO = "funcionarios.xlsx"
NOME_ABA = "dados"
MAX_ID_CHARS = 15

st.set_page_config(page_title="Cadastro em Planilha - Funcionários", page_icon="📒", layout="centered")

st.title("📒 Cadastro em Planilha (Nome e ID)")
st.caption("Preencha os campos, escolha quantas vezes repetir e grave na planilha Excel.")

# ---------------------------
# Funções utilitárias
# ---------------------------
def garantir_planilha(arquivo: str, aba: str):
    """
    Garante que o arquivo Excel exista com a aba e colunas necessárias.
    Se não existir, cria com colunas ['Nome', 'ID'].
    """
    if not os.path.exists(arquivo):
        df = pd.DataFrame(columns=["Nome", "ID"])
        with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=aba, index=False)

def carregar_planilha(arquivo: str, aba: str) -> pd.DataFrame:
    """
    Carrega a planilha; se não existir, cria.
    Sempre retorna um DataFrame com colunas ['Nome', 'ID'].
    """
    garantir_planilha(arquivo, aba)
    try:
        df = pd.read_excel(arquivo, sheet_name=aba, engine="openpyxl")
    except Exception:
        # Se a aba não existir ou der erro, recria o formato mínimo
        df = pd.DataFrame(columns=["Nome", "ID"])
        with pd.ExcelWriter(arquivo, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=aba, index=False)
    # Normaliza colunas
    df = df.reindex(columns=["Nome", "ID"])
    return df

def salvar_planilha(arquivo: str, aba: str, df: pd.DataFrame):
    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=aba, index=False)

def gerar_buffer_excel(df: pd.DataFrame, aba: str) -> bytes:
    """
    Gera um arquivo Excel em memória (bytes) para download.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=aba, index=False)
    buffer.seek(0)
    return buffer.read()

# ---------------------------
# Barra lateral (opções)
# ---------------------------
st.sidebar.header("⚙️ Configurações")
arquivo_excel = st.sidebar.text_input(
    "Nome do arquivo Excel",
    value=ARQUIVO_EXCEL_PADRAO,
    help="O arquivo será criado se não existir."
)
aba_excel = st.sidebar.text_input("Nome da aba", value=NOME_ABA)

st.sidebar.divider()
st.sidebar.caption("💡 Dica: deixe o arquivo padrão e a aba padrão para simplicidade.")

# ---------------------------
# Formulário de entrada
# ---------------------------
with st.form("form_cadastro", clear_on_submit=False):
    col1, col2 = st.columns([2, 1])
    with col1:
        nome = st.text_input("Nome do funcionário", max_chars=120, placeholder="Ex.: Maria Silva")
    with col2:
        qtd = st.number_input("Quantidade de repetições", min_value=1, value=1, step=1, help="Quantas linhas serão criadas.")

    id_func = st.text_input(
        f"ID (até {MAX_ID_CHARS} caracteres)",
        max_chars=MAX_ID_CHARS,
        placeholder="Ex.: 123456789ABCDEF"
    )

    # Preview do que será inserido
    if nome.strip() and id_func.strip() and qtd > 0:
        preview_df = pd.DataFrame({
            "Nome": [nome.strip()] * qtd,
            "ID": [id_func.strip()] * qtd
        })
        st.subheader("Pré-visualização das linhas a inserir")
        st.dataframe(preview_df, use_container_width=True, height=120 + min(qtd, 5) * 30)

    colb1, colb2, colb3 = st.columns([1, 1, 2])
    with colb1:
        submitted = st.form_submit_button("✅ Gravar")
    with colb2:
        limpar = st.form_submit_button("🧹 Limpar formulário")

# Limpar formulário (reseta estado visual; os inputs permanecem conforme Streamlit)
if limpar:
    st.experimental_rerun()

# ---------------------------
# Ação: gravar
# ---------------------------
if submitted:
    # Validações
    erros = []
    if not nome.strip():
        erros.append("Informe o **Nome do funcionário**.")
    if not id_func.strip():
        erros.append("Informe o **ID**.")
    if len(id_func.strip()) > MAX_ID_CHARS:
        erros.append(f"O **ID** não pode ultrapassar **{MAX_ID_CHARS}** caracteres.")
    if qtd < 1:
        erros.append("A **Quantidade de repetições** deve ser pelo menos 1.")

    if erros:
        for e in erros:
            st.error(e)
    else:
        # Carrega planilha atual, concatena e salva
        df_atual = carregar_planilha(arquivo_excel, aba_excel)
        novas_linhas = pd.DataFrame({
            "Nome": [nome.strip()] * qtd,
            "ID": [id_func.strip()] * qtd
        })
        df_final = pd.concat([df_atual, novas_linhas], ignore_index=True)

        try:
            salvar_planilha(arquivo_excel, aba_excel, df_final)
            st.success(f"✅ {qtd} linha(s) adicionada(s) com sucesso em **{arquivo_excel}** (aba **{aba_excel}**).")
            st.balloons()
        except PermissionError:
            st.error("❌ Não foi possível salvar. Verifique se o arquivo está **aberto** em outro programa (como o Excel) e feche-o.")
        except Exception as exc:
            st.error(f"❌ Erro ao salvar: {exc}")

# ---------------------------
# Exibir conteúdo atual e permitir download
# ---------------------------
st.subheader("📄 Conteúdo atual da planilha")
df_conteudo = carregar_planilha(arquivo_excel, aba_excel)
st.dataframe(df_conteudo, use_container_width=True, height=240)

# Botão de download
excel_bytes = gerar_buffer_excel(df_conteudo, aba_excel)
nome_download = f"funcionarios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
st.download_button(
    "⬇️ Baixar Excel atualizado",
    data=excel_bytes,
    file_name=nome_download,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Baixe uma cópia do arquivo atualizado."
)

st.caption("Pronto! Agora é só repetir o processo para inserir mais nomes e IDs.")