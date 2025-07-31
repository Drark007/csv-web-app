import streamlit as st
import pandas as pd
import io
from datetime import datetime
import chardet

st.set_page_config(page_title="Extração de dados CSV", layout="wide")
st.title("🔍 Extração de Dados de Arquivos CSV")

st.markdown("""
Esta aplicação permite extrair colunas específicas de múltiplos arquivos CSV com a mesma estrutura e gerar uma planilha Excel consolidada, com suporte à conversão de unidades e verificação de estrutura.
""")

uploaded_files = st.file_uploader("Selecione um ou mais arquivos CSV", accept_multiple_files=True, type="csv")

if uploaded_files:
    colunas_iguais = True
    nomes_colunas = None
    arquivos = []
    nomes_arquivos = []
    erro_decodificacao = []

    for arquivo in uploaded_files:
        content = arquivo.read()
        result = chardet.detect(content)
        encoding = result['encoding'] if result['confidence'] > 0.5 else 'utf-8'
        arquivo.seek(0)

        try:
            df = pd.read_csv(arquivo, encoding=encoding, sep=None, engine='python')
        except Exception as e:
            erro_decodificacao.append(arquivo.name)
            continue

        arquivos.append(df)
        nomes_arquivos.append(arquivo.name)

        if nomes_colunas is None:
            nomes_colunas = list(df.columns)
        elif list(df.columns) != nomes_colunas:
            colunas_iguais = False
            break

    if erro_decodificacao:
        st.warning(f"Os seguintes arquivos não puderam ser lidos: {', '.join(erro_decodificacao)}")

    if not arquivos:
        st.stop()

    if not colunas_iguais:
        st.error("Os arquivos selecionados não possuem a mesma estrutura de colunas.")
    else:
        st.success("Todos os arquivos possuem estrutura idêntica.")
        st.markdown("**Arquivos carregados:**")
        st.write(nomes_arquivos)

        colunas_escolhidas = st.multiselect("Selecione as colunas a serem extraídas", nomes_colunas)

        converter_volume = st.checkbox("Converter volumes de uL para mL (quando aplicável)")
        remover_duplicatas = st.checkbox("Remover linhas duplicadas")

        if colunas_escolhidas:
            if st.button("Gerar arquivo Excel"):
                with st.spinner("Processando os dados..."):
                    excel_buffer = io.BytesIO()
                    writer = pd.ExcelWriter(excel_buffer, engine='openpyxl')
                    planilha = pd.DataFrame()

                    for idx, df in enumerate(arquivos, start=1):
                        dados = df[colunas_escolhidas].copy()

                        # Conversão de volume
                        for col in dados.columns:
                            if "vol" in col.lower() and converter_volume:
                                dados[col] = dados[col].apply(pd.to_numeric, errors='coerce') / 1000

                        # Sugestão de unidade
                        for col in dados.columns:
                            if "pH" in col:
                                st.info(f"Coluna '{col}' pode conter dados de potencial (pH). Confirme a unidade esperada.")

                        if remover_duplicatas:
                            dados = dados.drop_duplicates()

                        # Nomear grupo
                        nome_grupo = f"Documento {idx}"
                        header = pd.DataFrame({col: [nome_grupo] if i == 0 else [None] for i, col in enumerate(colunas_escolhidas)})

                        planilha = pd.concat([planilha, header, dados, pd.DataFrame([[]])], ignore_index=True)

                    planilha.columns = colunas_escolhidas
                    planilha.to_excel(writer, index=False, sheet_name="Dados")
                    writer.close()

                    nome_arquivo = datetime.now().strftime("%Y_%m_%d") + "_extracao.xlsx"
                    st.download_button(
                        label="📥 Baixar Excel",
                        data=excel_buffer.getvalue(),
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("Arquivo gerado com sucesso!")
        else:
            st.info("Selecione pelo menos uma coluna para extrair.")
