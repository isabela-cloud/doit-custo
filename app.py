import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO

st.set_page_config(page_title="DOit - Atualização de Custos", layout="wide")
st.title("📊 DOit - Atualização de Custos")

# --- Carregar planilha DOit (fixa no servidor) ---
CAMINHO_DOIT = "ListagemdeProdutos DOit.xlsx"


@st.cache_data
def carregar_doit():
    df = pd.read_excel(CAMINHO_DOIT)
    df["# Referência"] = df["# Referência"].astype(str).str.strip()
    return df


try:
    df_doit = carregar_doit()
except FileNotFoundError:
    st.error("Arquivo 'ListagemdeProdutos DOit.xlsx' não encontrado na pasta do projeto.")
    st.stop()

st.success(f"Planilha DOit carregada: {len(df_doit)} produtos")

# --- Upload da planilha do fornecedor ---
st.divider()
st.subheader("1. Upload da planilha do fornecedor")

arquivo_fornecedor = st.file_uploader(
    "Arraste ou selecione a planilha do fornecedor (.xlsx, .xls)",
    type=["xlsx", "xls"],
)

if arquivo_fornecedor is None:
    st.stop()

# Ler todas as abas do arquivo do fornecedor
todas_abas = pd.read_excel(arquivo_fornecedor, header=None, sheet_name=None)
nomes_abas = list(todas_abas.keys())

st.info(f"📑 Abas encontradas: {', '.join(nomes_abas)} ({len(nomes_abas)} aba{'s' if len(nomes_abas) > 1 else ''})")

# Pré-visualização da primeira aba
df_raw = todas_abas[nomes_abas[0]]
st.write("**Pré-visualização da primeira aba (primeiras 15 linhas):**")
st.dataframe(df_raw.head(15), use_container_width=True)

# --- Configuração do fornecedor ---
st.divider()
st.subheader("2. Configuração")

col1, col2 = st.columns(2)

with col1:
    linha_header = st.number_input(
        "Linha do cabeçalho (0 = primeira linha)",
        min_value=0,
        max_value=len(df_raw) - 1,
        value=1,
        help="Indique em qual linha estão os nomes das colunas",
    )

# Recarregar todas as abas com header correto e concatenar
# A primeira aba usa o header indicado; as demais usam os mesmos nomes de coluna
dfs_abas = []
df_primeira = pd.read_excel(arquivo_fornecedor, header=int(linha_header), sheet_name=nomes_abas[0])
df_primeira.columns = [str(c).strip() for c in df_primeira.columns]
df_primeira["_aba_origem"] = nomes_abas[0]
dfs_abas.append(df_primeira)

colunas_base = df_primeira.columns.drop("_aba_origem")

for nome_aba in nomes_abas[1:]:
    df_aba = pd.read_excel(arquivo_fornecedor, header=None, sheet_name=nome_aba)
    # Aplicar os mesmos nomes de coluna da primeira aba
    if len(df_aba.columns) >= len(colunas_base):
        df_aba = df_aba.iloc[:, :len(colunas_base)]
        df_aba.columns = colunas_base
    else:
        # Se tiver menos colunas, usar os nomes disponíveis
        df_aba.columns = colunas_base[:len(df_aba.columns)]
    df_aba["_aba_origem"] = nome_aba
    dfs_abas.append(df_aba)

df_forn = pd.concat(dfs_abas, ignore_index=True)

colunas_forn = [c for c in df_forn.columns if c != "_aba_origem"]

with col2:
    nome_fornecedor = st.text_input(
        "Nome do fornecedor (para nome do arquivo)",
        value=arquivo_fornecedor.name.split(".")[0],
    )

col3, col4, col5, col6 = st.columns(4)

with col3:
    col_codigo = st.selectbox("Coluna do CÓDIGO (referência)", options=colunas_forn)

with col4:
    col_preco = st.selectbox("Coluna do PREÇO", options=colunas_forn)

with col5:
    # Caso o preço esteja separado em duas colunas (ex: "R$" em uma e valor em outra)
    usar_col_valor = st.checkbox("Valor em coluna separada?", value=False,
                                  help="Marque se o R$ está em uma coluna e o número em outra")

with col6:
    ipi = st.number_input("IPI (%)", min_value=0.0, max_value=100.0, value=0.0, step=0.25)

col_valor_separado = None
if usar_col_valor:
    col_valor_separado = st.selectbox("Coluna com o VALOR numérico", options=colunas_forn)

# --- Seleção do fabricante ---
fabricantes_doit = (
    df_doit[["Id do Fabricante", "Fabricante"]]
    .dropna(subset=["Id do Fabricante"])
    .drop_duplicates()
    .sort_values("Fabricante")
    .reset_index(drop=True)
)
fabricantes_doit["_label"] = (
    fabricantes_doit["Fabricante"].astype(str)
    + " (ID: "
    + fabricantes_doit["Id do Fabricante"].astype(int).astype(str)
    + ")"
)

fabricante_escolhido = st.selectbox(
    "Fabricante (selecione o fornecedor correspondente no DOit):",
    options=fabricantes_doit["_label"].tolist(),
    help="Escolha o fabricante cadastrado no DOit que corresponde a esta planilha.",
)

idx_selecionado = fabricantes_doit["_label"].tolist().index(fabricante_escolhido)
id_fabricante = fabricantes_doit.iloc[idx_selecionado]["Id do Fabricante"]

st.write("**Planilha do fornecedor (com cabeçalho aplicado):**")
st.dataframe(df_forn.head(10), use_container_width=True)


# --- Funções auxiliares ---
def parse_preco(valor):
    if pd.isna(valor):
        return None
    s = str(valor).strip()
    # Remover R$, espaços extras
    s = s.replace("R$", "").replace("r$", "").strip()
    s = s.replace(" ", "")
    if not s:
        return None

    # Formato brasileiro: 1.234,56 ou 1.234
    if "," in s:
        # 1.234,56 -> remove ponto (milhar), troca vírgula por ponto (decimal)
        s = s.replace(".", "").replace(",", ".")
    elif "." in s:
        if s.count(".") > 1:
            s = s.replace(".", "")
        else:
            # Um único ponto: verificar se é milhar ou decimal
            partes = s.split(".")
            if len(partes[1]) == 3:
                s = s.replace(".", "")  # milhar
    try:
        return float(s)
    except ValueError:
        return None


# --- Processamento ---
st.divider()
st.subheader("3. Resultado")

if st.button("🔄 Processar", type="primary"):
    # Limpar código do fornecedor
    df_forn["_codigo_limpo"] = df_forn[col_codigo].astype(str).str.strip()
    df_forn["_preco_limpo"] = df_forn[col_preco].apply(parse_preco)

    # Remover linhas sem código ou preço válido
    df_forn_valido = df_forn.dropna(subset=["_preco_limpo"]).copy()
    df_forn_valido = df_forn_valido[
        (df_forn_valido["_codigo_limpo"] != "nan")
        & (df_forn_valido["_codigo_limpo"] != "")
    ]

    # Cruzar com Doit pela referência (usar apenas 1 preço por código do fornecedor)
    df_forn_unico = df_forn_valido.drop_duplicates(subset=["_codigo_limpo"], keep="first")

    df_merge = df_doit.merge(
        df_forn_unico[["_codigo_limpo", "_preco_limpo"]],
        left_on="# Referência",
        right_on="_codigo_limpo",
        how="inner",
    )

    # Produtos do fornecedor que NÃO estão no DOit (precisam ser criados)
    refs_doit = set(df_doit["# Referência"].astype(str).str.strip())
    mask_nao_encontrado = ~df_forn_valido["_codigo_limpo"].isin(refs_doit)
    df_precisam_criar = df_forn_valido[mask_nao_encontrado].copy()

    # Guardar no session_state para uso após seleção do fabricante
    st.session_state["df_merge_raw"] = df_merge
    st.session_state["df_forn_valido"] = df_forn_valido
    st.session_state["df_precisam_criar"] = df_precisam_criar
    st.session_state["processado"] = True

# --- Seleção do fabricante e resultados ---
if st.session_state.get("processado", False):
    df_merge = st.session_state["df_merge_raw"]
    df_precisam_criar = st.session_state["df_precisam_criar"]

    # Filtrar o merge para manter apenas produtos do fabricante selecionado
    if not df_merge.empty:
        df_merge = df_merge[df_merge["Id do Fabricante"] == id_fabricante].copy()

    # Todos os produtos desse fabricante no Doit
    df_fabricante_doit = df_doit[df_doit["Id do Fabricante"] == id_fabricante].copy()
    # Refs que foram atualizadas
    refs_atualizadas = set(df_merge["# Referência"].astype(str).str.strip()) if not df_merge.empty else set()
    # Não atualizados = fabricante no Doit - atualizados
    df_nao_atualizados = df_fabricante_doit[
        ~df_fabricante_doit["# Referência"].isin(refs_atualizadas)
    ].copy()

    # --- Cálculos ---
    # Custo Líquido = Valor fornecedor * 1.10
    # CUSTO (bruto) = Custo Líquido * (1 + IPI%)
    ipi_fator = 1 + (ipi / 100)

    if not df_merge.empty:
        df_merge["_custo_liquido"] = (df_merge["_preco_limpo"] * 1.10).round(2)
        df_merge["_custo_bruto"] = (df_merge["_custo_liquido"] * ipi_fator).round(2)

    # --- Métricas ---
    st.success(f"✅ Processamento concluído!")

    col_info1, col_info2, col_info3, col_info4 = st.columns(4)
    col_info1.metric("Atualizados", len(df_merge))
    col_info2.metric("Precisam ser criados", len(df_precisam_criar))
    col_info3.metric("Não atualizados", len(df_nao_atualizados))
    col_info4.metric("IPI", f"{ipi}%")

    # --- Planilha 1: Modelo Custo para importação no DOit ---
    hoje = date.today().strftime("%d/%m/%Y")

    if not df_merge.empty:
        df_modelo_custo = pd.DataFrame(
            {
                "SKU": df_merge["SKU"].fillna(0).astype(int).astype(str),
                "FORNECEDOR": df_merge["Id do Fabricante"].fillna(0).astype(int).astype(str),
                "NOME ORIGINAL": "",
                "PART NUMBER": "",
                "CONDIÇÃO": "DDP",
                "CUSTO": df_merge["_custo_bruto"].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")),
                "MOEDA": "BRL",
                "CUSTO FINAL?": "",
                "CUSTO LÍQUIDO": df_merge["_custo_liquido"].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")),
                "MODIFICADO EM": hoje,
            }
        )
    else:
        df_modelo_custo = pd.DataFrame()

    # --- Planilha 2: Relatório completo para o cliente ---
    # Aba "Produtos DOit" - produtos do DOit atualizados (com preços novos)
    if not df_merge.empty:
        df_produtos_doit = df_merge.drop(
            columns=["_codigo_limpo", "_preco_limpo", "_custo_liquido", "_custo_bruto"],
            errors="ignore",
        )
    else:
        df_produtos_doit = pd.DataFrame()

    # Aba "Produtos" - planilha do fornecedor original
    df_forn_valido = st.session_state["df_forn_valido"]
    df_produtos_forn = df_forn_valido.drop(
        columns=["_codigo_limpo", "_preco_limpo", "_aba_origem"], errors="ignore"
    )

    # Aba "Precisam ser criados" - formato do fornecedor
    df_criar_saida = df_precisam_criar.drop(
        columns=["_codigo_limpo", "_preco_limpo", "_aba_origem"], errors="ignore"
    )

    # --- Exibição ---
    tab1, tab2, tab3, tab4 = st.tabs(
        ["📋 Modelo Custo (DOit)", "✅ Atualizados", "🆕 Precisam ser criados", "⚠️ Não atualizados"]
    )

    with tab1:
        st.dataframe(df_modelo_custo, use_container_width=True)

    with tab2:
        if not df_produtos_doit.empty:
            st.dataframe(
                df_produtos_doit[["SKU", "# Referência", "Nome", "Preço"]],
                use_container_width=True,
            )

    with tab3:
        st.dataframe(df_criar_saida, use_container_width=True)

    with tab4:
        if not df_nao_atualizados.empty:
            st.dataframe(
                df_nao_atualizados[["SKU", "# Referência", "Nome", "Preço"]],
                use_container_width=True,
            )

    # --- Texto padrão para envio ---
    st.divider()
    st.subheader("4. Texto para o cliente")

    texto_padrao = (
        f"Referente a {nome_fornecedor}, os custos foram atualizados:\n\n"
        f"Produtos que foram atualizados: {len(df_merge)}\n"
        f"Produtos que precisam ser criados: {len(df_precisam_criar)}\n"
        f"Produtos que não foram atualizados: {len(df_nao_atualizados)}\n"
        f"Atualizado na Luminata 1 e 2."
    )

    st.text_area("Copie e envie ao cliente:", value=texto_padrao, height=150)

    # --- Downloads ---
    st.divider()
    st.subheader("5. Downloads")

    col_dl1, col_dl2 = st.columns(2)

    # Download 1: Planilha para importação no DOit
    with col_dl1:
        st.write("**Planilha para importação no DOit:**")
        buffer1 = BytesIO()
        with pd.ExcelWriter(buffer1, engine="xlsxwriter") as writer:
            df_modelo_custo.to_excel(writer, index=False, sheet_name="Modelo Custo")
        buffer1.seek(0)

        st.download_button(
            label="📥 Baixar - Importação DOit",
            data=buffer1,
            file_name=f"custo_doit_{nome_fornecedor}_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Download 2: Relatório completo para o cliente
    with col_dl2:
        st.write("**Relatório completo para o cliente:**")
        buffer2 = BytesIO()
        with pd.ExcelWriter(buffer2, engine="xlsxwriter") as writer:
            if not df_produtos_doit.empty:
                df_produtos_doit.to_excel(writer, index=False, sheet_name="Produtos DOit")
            df_produtos_forn.to_excel(writer, index=False, sheet_name="Produtos")
            df_modelo_custo.to_excel(writer, index=False, sheet_name="Modelo Custo")
            if not df_criar_saida.empty:
                df_criar_saida.to_excel(writer, index=False, sheet_name="Precisam ser criados")
            if not df_nao_atualizados.empty:
                df_nao_atualizados.to_excel(
                    writer, index=False, sheet_name="Não foram atualizados"
                )
        buffer2.seek(0)

        st.download_button(
            label="📥 Baixar - Relatório Cliente",
            data=buffer2,
            file_name=f"relatorio_{nome_fornecedor}_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
