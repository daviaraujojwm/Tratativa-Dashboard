import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re

# =========================
# Configuração inicial
# =========================
st.set_page_config(page_title="Relatório de Faturamento", layout="wide")
st.title("📊 Automação - Relatório de Faturamento")

# Upload dos arquivos
arquivo1 = st.sidebar.file_uploader("Carregar 1ª planilha (ESL Desktop)", type=["xlsx", "xls"])
arquivo2 = st.sidebar.file_uploader("Carregar 2ª planilha (Modelo SIG)", type=["xlsx", "xls"])

# =========================
# Funções auxiliares
# =========================
def normalize_col(name):
    """Remove acentos, caracteres especiais e normaliza nomes de colunas."""
    if not isinstance(name, str):
        return str(name)
    name = name.strip()
    name = unicodedata.normalize("NFKD", name)
    name = name.encode("ASCII", "ignore").decode("utf-8")
    name = name.lower()
    name = re.sub(r'[^0-9a-z]+', '_', name)
    name = name.strip('_')
    return name

def remover_duplicadas(df):
    """Renomeia colunas duplicadas automaticamente (sem remover linhas)."""
    seen = {}
    new_cols = []
    renomeadas = {}

    for col in df.columns:
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            novo_nome = f"{col}_{seen[col]}"
            new_cols.append(novo_nome)
            renomeadas[col] = novo_nome

    df.columns = new_cols
    return df, renomeadas

def to_excel(df):
    """Exporta DataFrame para Excel com destaque."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Planilha Final")

        workbook = writer.book
        worksheet = writer.sheets["Planilha Final"]
        amarelo = workbook.add_format({"bg_color": "#FFFF00"})

        if "NM NF NCTE" in df.columns:
            col_j_idx = df.columns.get_loc("NM NF NCTE")
            worksheet.set_column(col_j_idx, col_j_idx, None, amarelo)

        if "Emissão CT-e" in df.columns:
            col_l_idx = df.columns.get_loc("Emissão CT-e")
            worksheet.set_column(col_l_idx, col_l_idx, None, amarelo)

    return output.getvalue()

# =========================
# Processamento principal
# =========================
if arquivo1 and arquivo2:
    df1 = pd.read_excel(arquivo1)
    df2 = pd.read_excel(arquivo2)

    # Mantém todas as linhas — sem remover nada
    df1 = df1.copy()
    df2 = df2.copy()

    # Tratamento de duplicadas apenas nos nomes
    df2, renomeadas = remover_duplicadas(df2)
    if renomeadas:
        st.warning("⚠️ Colunas duplicadas na 2ª planilha foram renomeadas automaticamente:")
        st.table(pd.DataFrame(list(renomeadas.items()), columns=["Coluna Original", "Coluna Renomeada"]))

    st.subheader("📋 Quantidade de Linhas Originais")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Planilha 1 (ESL Desktop)", len(df1))
    with col2:
        st.metric("Planilha 2 (Modelo SIG)", len(df2))

    # =========================
    # Normalização
    # =========================
    norm_df1 = {normalize_col(c): c for c in df1.columns}
    norm_df2 = {normalize_col(c): c for c in df2.columns}

    # =========================
    # TRATATIVAS PLANILHA 1
    # =========================
    df1 = df1.replace("", pd.NA)
    colunas = list(df1.columns)

    if len(colunas) >= 9:
        col_g, col_h, col_i = colunas[6], colunas[7], colunas[8]

        def combinar_valores(x):
            partes = []
            if pd.notna(x[col_h]):
                partes.append(f"{x[col_g]}/{x[col_h]}")
            if pd.notna(x[col_i]):
                partes.append(f"{x[col_g]}/{x[col_i]}")
            return " | ".join(partes)

        df1["NM NF NCTE"] = df1.apply(combinar_valores, axis=1)

    if len(colunas) >= 3:
        df1["Emissão CT-e"] = df1[colunas[2]]

    if "Classificação" in df1.columns:
        df1["Classificação"] = df1["Classificação"].fillna("EMBARQUE")

    if "Total Frete" in df1.columns:
        df1["Total Frete"] = pd.to_numeric(df1["Total Frete"], errors="coerce").fillna(0)

    # =========================
    # MAPEAMENTO ENTRE PLANILHAS
    # =========================
    mapeamento = {
        "NM NF NCTE": "Número CT-e",
        "nº de referência": "Número Coleta",
        "Classificação": "Tipo Movimento",
        "Remetente": "Remetente",
        "Cidade Remetente": "Cidade Remetente",
        "Cidade Origem": "Cidade Origem",
        "Bairro Remetente": "Bairro Remetente",
        "Destinatário": "Destinatário",
        "Cidade Destinatária": "Cidade Destinatário",
        "Bairro Destinatário": "Bairro Destinatário",
        "Cidade Destino": "Cidade Destino",
        "Tabela de Preço": "Tabela de Preço",
        "Nota Fiscal": "Nota Fiscal",
        "Valor NF": "Valor N.F.",
        "CFOP": "CFOP",
        "Volume": "Volume",
        "Peso Taxado": "Peso Taxado",
        "Peso Real": "Peso Real",
        "Frete Peso": "Frete Peso",
        "AD Valorem": "ADValorem",
        "Natureza": "Natureza",
        "Emissão CT-e": "Emissão CT-e",
        "CNPJ Pagador": "CPF/CNPJ Faturado",
        "Pagador de Frete": "Cliente Faturado",
        "Total Frete": "Valor Frete",
        "%Imposto": "Taxa ICMS",
        "Valor Imposto": "Valor ICMS",
        "Outros Valores": "Outros",
        "Valor Pedágio": "Pedágio"
    }

    df_final = pd.concat([df2], ignore_index=True)

    for _, row in df1.iterrows():
        nova_linha = {}
        for col_origem, col_destino in mapeamento.items():
            n_src = normalize_col(col_origem)
            src_colname = norm_df1.get(n_src)
            if src_colname and src_colname in row:
                nova_linha[col_destino] = row[src_colname]
        df_final = pd.concat([df_final, pd.DataFrame([nova_linha])], ignore_index=True)

    # =========================
    # VERIFICAÇÃO DE LINHAS
    # =========================
    total_esperado = len(df1) + len(df2)
    total_gerado = len(df_final)

    st.subheader("📊 Verificação de Linhas")
    st.write(f"Total esperado: **{total_esperado}**")
    st.write(f"Total gerado: **{total_gerado}**")
    if total_esperado == total_gerado:
        st.success("✅ Quantidade de linhas confere exatamente com a soma das planilhas originais!")
    else:
        st.warning("⚠️ Quantidade de linhas não confere — verifique mapeamentos ou dados vazios!")

    # =========================
    # Valor Frete Consolidado
    # =========================
    frete_total = pd.to_numeric(df_final.get("Valor Frete", pd.Series()), errors="coerce").sum()
    st.subheader("💰 Resumo do Valor de Frete")
    st.write(f"📦 Valor Total Consolidado: **{frete_total:,.2f}**")

    # =========================
    # Download e Preview
    # =========================
    st.subheader("🔎 Pré-visualização (5 primeiras linhas)")
    st.dataframe(df_final.head())

    st.sidebar.download_button(
        label="📥 Baixar Planilha Final",
        data=to_excel(df_final),
        file_name="planilha_final_consolidada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
