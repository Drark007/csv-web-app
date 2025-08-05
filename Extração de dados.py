import streamlit as st
import pandas as pd
import io
from datetime import datetime
import chardet
import numpy as np

st.set_page_config(page_title="Extração de dados CSV", layout="wide")
st.title("🔍 Extração de Dados CSV")

st.markdown("""
Esta aplicação permite extrair colunas específicas de múltiplos arquivos CSV com estrutura idêntica ou diferente,
com suporte à verificação de estrutura, conversão de unidades, visualização prévia, exportação para Excel
em abas organizadas por estrutura.
""")

uploaded_files = st.file_uploader("Selecione um ou mais arquivos CSV", accept_multiple_files=True, type="csv")

if uploaded_files:
    grupos = {}
    arquivos_info = []
    erro_decodificacao = []

    for arquivo in uploaded_files:
        content = arquivo.read()
        result = chardet.detect(content)
        encoding = result['encoding'] if result['confidence'] > 0.5 else 'utf-8'
        arquivo.seek(0)

        try:
            df = pd.read_csv(arquivo, encoding=encoding, sep=None, engine='python')
        except Exception:
            erro_decodificacao.append(arquivo.name)
            continue

        estrutura = tuple(df.columns)
        if estrutura not in grupos:
            grupos[estrutura] = []
        grupos[estrutura].append((arquivo.name, df))
        arquivos_info.append((arquivo.name, df, estrutura))

    if erro_decodificacao:
        st.warning(f"Arquivos não lidos: {', '.join(erro_decodificacao)}")

    if not arquivos_info:
        st.stop()

    st.success(f"{len(grupos)} estrutura(s) distinta(s) detectada(s).")
    st.write("**Arquivos agrupados por estrutura:**")
    for i, (estrutura, arquivos) in enumerate(grupos.items(), start=1):
        st.markdown(f"**Estrutura {i}**: {len(arquivos)} arquivo(s) - Colunas: {list(estrutura)}")

    estrutura_alertas = []
    for i, (estrutura, arquivos) in enumerate(grupos.items(), start=1):
        alerta = []
        for col in estrutura:
            if "ph" in col.lower():
                alerta.append("pH")
            if "temp" in col.lower():
                alerta.append("Temperatura")
            if "press" in col.lower():
                alerta.append("Pressão")
            if "vol" in col.lower() or "volume" in col.lower():
                alerta.append("Volume")
        if alerta:
            estrutura_alertas.append(f"Estrutura {i} pode conter dados de: {', '.join(set(alerta))}")

    if estrutura_alertas:
        with st.expander("⚠️ Alertas detectados nas estruturas"):
            for msg in estrutura_alertas:
                st.info(msg)

    converter_dados = {}
    for i, estrutura in enumerate(grupos.keys(), start=1):
        converter_dados[estrutura] = {}
        st.markdown(f"### Estrutura {i} – Definir unidades e conversões")
        for col in estrutura:
            unidade = st.selectbox(f"Coluna '{col}': Qual a unidade?", ["Não definido", "µL", "mL", "L", "°C", "K", "nm", "µm", "Hz", "kHz"], key=f"unidade_{i}_{col}")
            if unidade != "Não definido":
                alvo = st.selectbox(f"Converter '{col}' para:", ["Não converter", "µL", "mL", "L", "°C", "K", "nm", "µm", "Hz", "kHz"], key=f"converter_{i}_{col}")
                if alvo != "Não converter" and alvo != unidade:
                    converter_dados[estrutura][col] = (unidade, alvo)

    gerar_excel = st.button("Gerar arquivo Excel")

    if gerar_excel:
        with st.spinner("Processando..."):
            excel_buffer = io.BytesIO()
            writer = pd.ExcelWriter(excel_buffer, engine='openpyxl')

            for idx, (estrutura, arquivos) in enumerate(grupos.items(), start=1):
                planilha = pd.DataFrame()
                blocos = []
                for nome, df in arquivos:
                    dados = df.copy()
                    for col in dados.columns:
                        if col in converter_dados[estrutura]:
                            origem, destino = converter_dados[estrutura][col]
                            try:
                                dados[col] = pd.to_numeric(dados[col], errors='coerce')
                                if origem == "µL" and destino == "mL":
                                    dados[col] = dados[col] / 1000
                                elif origem == "mL" and destino == "µL":
                                    dados[col] = dados[col] * 1000
                                elif origem == "mL" and destino == "L":
                                    dados[col] = dados[col] / 1000
                                elif origem == "°C" and destino == "K":
                                    dados[col] = dados[col] + 273.15
                                elif origem == "K" and destino == "°C":
                                    dados[col] = dados[col] - 273.15
                                elif origem == "nm" and destino == "µm":
                                    dados[col] = dados[col] / 1000
                                elif origem == "µm" and destino == "nm":
                                    dados[col] = dados[col] * 1000
                                elif origem == "Hz" and destino == "kHz":
                                    dados[col] = dados[col] / 1000
                                elif origem == "kHz" and destino == "Hz":
                                    dados[col] = dados[col] * 1000
                            except Exception:
                                pass
                    dados.columns = [f"{col} ({nome})" for col in dados.columns]
                    blocos.append(dados)
                    blocos.append(pd.DataFrame(np.nan, index=range(len(dados)), columns=[""]))

                final = pd.concat(blocos, axis=1)
                final.to_excel(writer, index=False, sheet_name=f"Estrutura {idx}")

            writer.close()

            nome_arquivo = datetime.now().strftime("%Y_%m_%d") + "_extracao_completa.xlsx"
            st.download_button(
                label="📥 Baixar Excel",
                data=excel_buffer.getvalue(),
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Arquivo Excel gerado com sucesso.")
