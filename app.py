import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO

# ─── Configuração da página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="DOit - Atualização de Custos",
    page_icon="💰",
    layout="wide",
)

# ─── Estilo customizado ──────────────────────────────────────────────────────
st.markdown("""
<style>
    .main .block-container { padding-top: 1rem; max-width: 1200px; }
    h1 { color: #2c3e50; }
    .stTabs [data-baseweb="tab-list"] { gap: 1.5rem; }
    .section-header {
        font-size: 1.1rem;
        font-weight: 600;
        color: #2c3e50;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    .info-box {
        background: #eef6ff;
        border-left: 4px solid #3b82f6;
        border-radius: 6px;
        padding: 0.8rem 1rem;
        font-size: 0.85rem;
        color: #1e40af;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ─── Carregar planilha DOit ───────────────────────────────────────────────────
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

# ─── Header ───────────────────────────────────────────────────────────────────
st.title("💰 Atualização de Custos")
st.markdown(f'<div class="info-box">📦 Base DOit carregada com <strong>{len(df_doit):,}</strong> produtos | {df_doit["Fabricante"].nunique()} fabricantes</div>', unsafe_allow_html=True)

# ─── Upload ───────────────────────────────────────────────────────────────────
st.divider()
st.markdown('<div class="section-header">📁 Upload da planilha do fornecedor</div>', unsafe_allow_html=True)

arquivo_fornecedor = st.file_uploader(
    "Arraste ou selecione a planilha (.xlsx, .xls)",
    type=["xlsx", "xls"],
)

if arquivo_fornecedor is None:
    st.stop()

# ─── Leitura de abas ─────────────────────────────────────────────────────────
todas_abas = pd.read_excel(arquivo_fornecedor, header=None, sheet_name=None)
nomes_abas = list(todas_abas.keys())

st.markdown(
    f'<div class="info-box">📑 Arquivo: <strong>{arquivo_fornecedor.name}</strong> — '
    f'{len(nomes_abas)} aba{"s" if len(nomes_abas) > 1 else ""} encontrada{"s" if len(nomes_abas) > 1 else ""} '
    f'({", ".join(nomes_abas)})</div>',
    unsafe_allow_html=True,
)

df_raw = todas_abas[nomes_abas[0]]

with st.expander("👁️ Pré-visualização (primeiras 15 linhas)", expanded=False):
    st.dataframe(df_raw.head(15), use_container_width=True)

# ─── Configuração ────────────────────────────────────────────────────────────
st.divider()
st.markdown('<div class="section-header">⚙️ Configuração</div>', unsafe_allow_html=True)

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
dfs_abas = []
df_primeira = pd.read_excel(arquivo_fornecedor, header=int(linha_header), sheet_name=nomes_abas[0])
df_primeira.columns = [str(c).strip() for c in df_primeira.columns]
df_primeira["_aba_origem"] = nomes_abas[0]
dfs_abas.append(df_primeira)

colunas_base = df_primeira.columns.drop("_aba_origem")

for nome_aba in nomes_abas[1:]:
    df_aba = pd.read_excel(arquivo_fornecedor, header=None, sheet_name=nome_aba)
    if len(df_aba.columns) >= len(colunas_base):
        df_aba = df_aba.iloc[:, :len(colunas_base)]
        df_aba.columns = colunas_base
    else:
        df_aba.columns = colunas_base[:len(df_aba.columns)]
    df_aba["_aba_origem"] = nome_aba
    dfs_abas.append(df_aba)

df_forn = pd.concat(dfs_abas, ignore_index=True)
colunas_forn = [c for c in df_forn.columns if c != "_aba_origem"]

with col2:
    nome_fornecedor = st.text_input(
        "Nome do fornecedor",
        value=arquivo_fornecedor.name.split(".")[0],
    )

col3, col4, col5, col6 = st.columns(4)

with col3:
    col_codigo = st.selectbox("Coluna do CÓDIGO", options=colunas_forn)

with col4:
    col_preco = st.selectbox("Coluna do PREÇO", options=colunas_forn)

with col5:
    ipi = st.number_input("IPI (%)", min_value=0.0, max_value=100.0, value=0.0, step=0.25)

# ─── Opções avançadas ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">🔧 Opções de compatibilização</div>', unsafe_allow_html=True)

col_opt1, col_opt2, col_opt3 = st.columns(3)

with col_opt1:
    normalizar_codigos = st.checkbox(
        "🔄 Normalizar códigos (remover acabamentos e hifens)",
        value=False,
        help="Ex: Revoluz envia 'RI-H54414-1-BFM OU PTO' e no DOit é 'RI-H54414-1'. "
             "Remove sufixos de acabamento e hifens para compatibilizar. "
             "Usar com: Revoluz, Revolux e similares.",
    )

with col_opt2:
    concatenar_colunas = st.checkbox(
        "🔗 Concatenar colunas para formar código",
        value=False,
        help="Ex: Spotline tem ID=84 e Descrição='385/2 PLAFON SMART...', no DOit é 'SL-84-385-2'. "
             "Junta o ID + primeira palavra da descrição para formar o código completo. "
             "Usar com: Spotline.",
    )

with col_opt3:
    usar_col_valor = st.checkbox(
        "💲 Valor em coluna separada",
        value=False,
        help="Ex: Golden Art tem 'R$' em uma coluna e o valor numérico em outra. "
             "Selecione a coluna com o número após ativar. "
             "Usar com: Golden Art.",
    )

col_valor_separado = None
if usar_col_valor:
    col_valor_separado = st.selectbox("Coluna com o VALOR numérico", options=colunas_forn)

col_concat_segunda = None
if concatenar_colunas:
    col_concat_segunda = st.selectbox(
        "Segunda coluna (será extraída a primeira palavra e concatenada ao código)",
        options=colunas_forn,
        help="A primeira palavra desta coluna será unida ao código com hífen. Ex: ID=84, Descrição='385/2 PLAFON...' → 84-385-2",
    )

# ─── Seleção do fabricante ────────────────────────────────────────────────────
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
    "🏭 Fabricante no DOit",
    options=fabricantes_doit["_label"].tolist(),
    help="Selecione o fabricante cadastrado no DOit correspondente a esta planilha.",
)

idx_selecionado = fabricantes_doit["_label"].tolist().index(fabricante_escolhido)
id_fabricante = fabricantes_doit.iloc[idx_selecionado]["Id do Fabricante"]

with st.expander("📋 Planilha do fornecedor (com cabeçalho aplicado)", expanded=False):
    st.dataframe(df_forn[colunas_forn].head(10), use_container_width=True)


# ─── Funções auxiliares ───────────────────────────────────────────────────────
def parse_preco(valor):
    if pd.isna(valor):
        return None
    s = str(valor).strip()
    s = s.replace("R$", "").replace("r$", "").strip()
    s = s.replace(" ", "")
    if not s:
        return None

    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "." in s:
        if s.count(".") > 1:
            s = s.replace(".", "")
        else:
            partes = s.split(".")
            if len(partes[1]) == 3:
                s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


# ─── Processamento ────────────────────────────────────────────────────────────
st.divider()
st.markdown('<div class="section-header">🔄 Processamento</div>', unsafe_allow_html=True)

if st.button("▶️ Processar atualização", type="primary", use_container_width=True):
    # Limpar código do fornecedor
    df_forn["_codigo_limpo"] = df_forn[col_codigo].astype(str).str.strip()
    # Remover .0 de números inteiros lidos como float (ex: 9766.0 -> 9766)
    df_forn["_codigo_limpo"] = df_forn["_codigo_limpo"].str.replace(r"\.0$", "", regex=True)

    # Concatenar colunas se ativado (ex: Spotline: ID + primeira palavra da descrição)
    if concatenar_colunas and col_concat_segunda:
        def _extrair_primeira_palavra(val):
            s = str(val).strip()
            return s.split()[0] if s and s != "nan" else ""

        df_forn["_segunda_parte"] = df_forn[col_concat_segunda].apply(_extrair_primeira_palavra)
        # Concatenar: código + "-" + primeira palavra (com / trocado por -)
        df_forn["_codigo_limpo"] = (
            df_forn["_codigo_limpo"] + "-" + df_forn["_segunda_parte"].str.replace("/", "-", regex=False)
        )
        # Limpar casos onde uma das partes é vazia
        df_forn["_codigo_limpo"] = df_forn["_codigo_limpo"].str.strip("-")

    df_forn["_preco_limpo"] = df_forn[col_preco].apply(parse_preco)

    # Remover linhas sem código ou preço válido
    df_forn_valido = df_forn.dropna(subset=["_preco_limpo"]).copy()
    df_forn_valido = df_forn_valido[
        (df_forn_valido["_codigo_limpo"] != "nan")
        & (df_forn_valido["_codigo_limpo"] != "")
    ]

    # Normalização de códigos (se ativada)
    if normalizar_codigos:
        import re

        # Acabamentos conhecidos (Revoluz e similares)
        _acabamentos = [
            "BCO", "BCX", "BFM", "CRT", "DSB", "CRG", "GRM", "MCX",
            "PTB", "PTO", "PTX", "VBE", "VCR", "VOC", "CORES",
            "BR", "PT", "DO", "CR", "CZ", "VD", "AM", "AZ", "LR", "MR",
        ]

        def remover_acabamento(codigo):
            """Remove sufixos de acabamento do código do fornecedor."""
            s = str(codigo).strip().upper()
            partes = s.split("-")
            while len(partes) > 1 and any(acab in partes[-1] for acab in _acabamentos + ["OU"]):
                partes.pop()
            return "-".join(partes)

        def normalizar(codigo):
            """Remove prefixos de fabricante, hifens, traços, barras e espaços para comparação."""
            s = str(codigo).strip().upper()
            # Remover prefixos comuns de fabricante (ex: SL-, SP-)
            s = re.sub(r"^[A-Z]{2,3}-", "", s)
            # Remover hifens, barras e espaços
            s = re.sub(r"[-/\s]", "", s)
            return s

        # Normalizar código do fornecedor: remover acabamento + remover hifens
        df_forn_valido["_codigo_norm"] = df_forn_valido["_codigo_limpo"].apply(
            lambda x: normalizar(remover_acabamento(x))
        )

        # Normalizar referências do DOit: remover hifens (usar cópia para não alterar cache)
        df_doit_norm = df_doit.copy()
        df_doit_norm["_ref_norm"] = df_doit_norm["# Referência"].apply(normalizar)

        # Cruzar usando códigos normalizados
        df_forn_unico = df_forn_valido.drop_duplicates(subset=["_codigo_norm"], keep="first")

        df_merge = df_doit_norm.merge(
            df_forn_unico[["_codigo_norm", "_preco_limpo", "_codigo_limpo"]],
            left_on="_ref_norm",
            right_on="_codigo_norm",
            how="inner",
        )

        # Produtos do fornecedor que NÃO estão no DOit
        refs_doit_norm = set(df_doit_norm["_ref_norm"])
        mask_nao_encontrado = ~df_forn_valido["_codigo_norm"].isin(refs_doit_norm)
        df_precisam_criar = df_forn_valido[mask_nao_encontrado].copy()

    else:
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

    # Guardar no session_state
    st.session_state["df_merge_raw"] = df_merge
    st.session_state["df_forn_valido"] = df_forn_valido
    st.session_state["df_precisam_criar"] = df_precisam_criar
    st.session_state["processado"] = True

# ─── Resultados ───────────────────────────────────────────────────────────────
if st.session_state.get("processado", False):
    df_merge = st.session_state["df_merge_raw"]
    df_precisam_criar = st.session_state["df_precisam_criar"]

    # Filtrar o merge para manter apenas produtos do fabricante selecionado
    if not df_merge.empty:
        df_merge = df_merge[df_merge["Id do Fabricante"] == id_fabricante].copy()

    # Todos os produtos desse fabricante no Doit
    df_fabricante_doit = df_doit[df_doit["Id do Fabricante"] == id_fabricante].copy()
    refs_atualizadas = set(df_merge["# Referência"].astype(str).str.strip()) if not df_merge.empty else set()
    df_nao_atualizados = df_fabricante_doit[
        ~df_fabricante_doit["# Referência"].isin(refs_atualizadas)
    ].copy()

    # ─── Cálculos ─────────────────────────────────────────────────────────────
    ipi_fator = 1 + (ipi / 100)

    if not df_merge.empty:
        df_merge["_custo_liquido"] = (df_merge["_preco_limpo"] * 1.10).round(2)
        df_merge["_custo_bruto"] = (df_merge["_custo_liquido"] * ipi_fator).round(2)

    # ─── Métricas ─────────────────────────────────────────────────────────────
    st.divider()
    st.markdown('<div class="section-header">📊 Resultado</div>', unsafe_allow_html=True)

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("✅ Atualizados", f"{len(df_merge):,}")
    col_m2.metric("🆕 Precisam ser criados", f"{len(df_precisam_criar):,}")
    col_m3.metric("⚠️ Não atualizados", f"{len(df_nao_atualizados):,}")
    col_m4.metric("📦 IPI", f"{ipi}%")

    # ─── Modelo Custo ─────────────────────────────────────────────────────────
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

    # ─── Relatório ────────────────────────────────────────────────────────────
    if not df_merge.empty:
        df_produtos_doit = df_merge.drop(
            columns=["_codigo_limpo", "_preco_limpo", "_custo_liquido", "_custo_bruto", "_codigo_norm", "_ref_norm", "_segunda_parte"],
            errors="ignore",
        )
    else:
        df_produtos_doit = pd.DataFrame()

    df_forn_valido = st.session_state["df_forn_valido"]
    df_produtos_forn = df_forn_valido.drop(
        columns=["_codigo_limpo", "_preco_limpo", "_aba_origem", "_codigo_norm", "_segunda_parte"], errors="ignore"
    )

    df_criar_saida = df_precisam_criar.drop(
        columns=["_codigo_limpo", "_preco_limpo", "_aba_origem", "_codigo_norm", "_segunda_parte"], errors="ignore"
    )

    # ─── Tabs de visualização ─────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs(
        ["📋 Modelo Custo", "✅ Atualizados", "🆕 Precisam ser criados", "⚠️ Não atualizados"]
    )

    with tab1:
        if not df_modelo_custo.empty:
            st.dataframe(df_modelo_custo, use_container_width=True, height=400)
        else:
            st.info("Nenhum produto para atualizar.")

    with tab2:
        if not df_produtos_doit.empty:
            st.dataframe(
                df_produtos_doit[["SKU", "# Referência", "Nome", "Preço"]],
                use_container_width=True,
                height=400,
            )
        else:
            st.info("Nenhum produto atualizado.")

    with tab3:
        if not df_criar_saida.empty:
            st.dataframe(df_criar_saida, use_container_width=True, height=400)
        else:
            st.info("Todos os produtos do fornecedor já existem no DOit.")

    with tab4:
        if not df_nao_atualizados.empty:
            st.dataframe(
                df_nao_atualizados[["SKU", "# Referência", "Nome", "Preço"]],
                use_container_width=True,
                height=400,
            )
        else:
            st.info("Todos os produtos do fabricante foram atualizados.")

    # ─── Texto para o cliente ─────────────────────────────────────────────────
    st.divider()
    st.markdown('<div class="section-header">💬 Texto para o cliente</div>', unsafe_allow_html=True)

    texto_padrao = (
        f"Referente a {nome_fornecedor}, os custos foram atualizados:\n\n"
        f"Produtos que foram atualizados: {len(df_merge)}\n"
        f"Produtos que precisam ser criados: {len(df_precisam_criar)}\n"
        f"Produtos que não foram atualizados: {len(df_nao_atualizados)}\n"
        f"Atualizado na Luminata 1 e 2."
    )

    st.text_area("Copie e envie ao cliente:", value=texto_padrao, height=140, label_visibility="collapsed")

    # ─── Downloads ────────────────────────────────────────────────────────────
    st.divider()
    st.markdown('<div class="section-header">📥 Downloads</div>', unsafe_allow_html=True)

    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        buffer1 = BytesIO()
        with pd.ExcelWriter(buffer1, engine="xlsxwriter") as writer:
            df_modelo_custo.to_excel(writer, index=False, sheet_name="Modelo Custo")
        buffer1.seek(0)

        st.download_button(
            label="📥 Importação DOit",
            data=buffer1,
            file_name=f"custo_doit_{nome_fornecedor}_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_dl2:
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
            label="📥 Relatório Cliente",
            data=buffer2,
            file_name=f"relatorio_{nome_fornecedor}_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


