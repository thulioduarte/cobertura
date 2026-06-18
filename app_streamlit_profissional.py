# -*- coding: utf-8 -*-
"""
Interface Streamlit profissional para o Estudo de Cobertura.

Como usar:
1) Salve este arquivo na mesma pasta do motor:
   gerar_estudo_cobertura_anexo_corrigido.py
2) Instale as dependências:
   pip install streamlit pandas openpyxl xlsxwriter numpy
3) Execute:
   streamlit run app_streamlit_profissional.py

Observação:
- Este app não altera a lógica do estudo.
- Ele apenas cria uma camada visual mais profissional por cima do arquivo-base.
"""

from __future__ import annotations

import importlib
import inspect
import os
import tempfile
from pathlib import Path
from typing import Dict, Iterable, Optional, Tuple

import pandas as pd
import streamlit as st


# ============================================================
# Configurações gerais
# ============================================================

LOGO_URL = "https://attachments.gupy.io/production/companies/37162/career/77428/images/2023-10-05_14-45_companyLogoUrl.png"
ENGINE_MODULE = os.getenv("COBERTURA_ENGINE_MODULE", "gerar_estudo_cobertura_anexo_corrigido")

APP_TITLE = "Estudo de Cobertura"
APP_SUBTITLE = "Análise Sell-in x Sell-out com validações, escala de volumetria e geração de Excel analítico."
MANUAL_FILE = "Manual de Uso - Estudo de Cobertura.docx"
BOAS_PRATICAS_FILE = "Boas Praticas - Estudo de Cobertura.md"

METRICAS = {
    "Volume": "volume",
    "Quantidade": "quantia",
    "Volume variável": "volume_variavel",
}

NIVEIS = {
    "Categoria": "CATEGORIA",
    "Nível 1": "NIVEL1",
    "Nível 2": "NIVEL2",
    "Est Mer 7": "ESTMER7",
}

OPCOES_SAIDA = {
    "resumo_categorias": "Resumo Categorias",
    "abas_categorias": "Abas por categoria/PROD",
    "base_skus": "Base SKUs",
    "base_contribuicao_sellout": "Base Contribuição Sell-out",
    "skus_por_categoria": "SKUs por Categoria",
    "crosschecks": "Crosschecks",
    "parametros": "Parâmetros",
    "descricao_calculos": "Descrição Cálculos",
    "avisos": "Avisos",
    "graficos_cobertura": "Gráficos Cobertura",
    "abas_auxiliares_comparacao": "Abas auxiliares da comparação",
    "top20_sku_canal_uf": "TOP 20 SKU por UF/Canal",
}

OPCOES_PADRAO = {
    "resumo_categorias": True,
    "abas_categorias": True,
    "base_skus": True,
    "base_contribuicao_sellout": False,
    "skus_por_categoria": True,
    "crosschecks": True,
    "parametros": True,
    "descricao_calculos": True,
    "avisos": True,
    "graficos_cobertura": True,
    "abas_auxiliares_comparacao": True,
    "top20_sku_canal_uf": False,
}


# ============================================================
# Inicialização visual
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon=LOGO_URL,
    layout="wide",
    initial_sidebar_state="expanded",
)


def aplicar_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --brand-blue: #062B49;
            --brand-cyan: #00A6C8;
            --brand-green: #35C48B;
            --brand-bg: #F5F7FA;
            --brand-card: #FFFFFF;
            --brand-border: #E3E8EF;
            --brand-text: #17202A;
            --muted-text: #667085;
        }

        .stApp {
            background: linear-gradient(180deg, #F5F7FA 0%, #FFFFFF 42%, #F8FAFC 100%);
        }

        div[data-testid="stToolbar"] {
            visibility: hidden;
            height: 0%;
            position: fixed;
        }

        .block-container {
            padding-top: 1.35rem;
            padding-bottom: 3rem;
            max-width: 1320px;
        }

        .hero {
            background:
                radial-gradient(circle at top right, rgba(0, 166, 200, 0.16), transparent 30%),
                linear-gradient(135deg, #062B49 0%, #0B3D61 42%, #0D5E7E 100%);
            border: 1px solid rgba(255,255,255,0.15);
            border-radius: 26px;
            padding: 28px 30px;
            color: white;
            box-shadow: 0 18px 48px rgba(6, 43, 73, 0.18);
            margin-bottom: 22px;
        }

        .hero-grid {
            display: grid;
            grid-template-columns: 140px 1fr;
            gap: 24px;
            align-items: center;
        }

        .logo-box {
            background: rgba(255,255,255,0.96);
            border-radius: 22px;
            padding: 18px;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 104px;
            box-shadow: 0 16px 36px rgba(0,0,0,0.16);
        }

        .logo-box img {
            max-width: 104px;
            max-height: 76px;
            object-fit: contain;
        }

        .hero h1 {
            font-size: 2.2rem;
            line-height: 1.08;
            margin: 0 0 10px 0;
            letter-spacing: -0.03em;
            font-weight: 800;
        }

        .hero p {
            margin: 0;
            color: rgba(255,255,255,0.85);
            font-size: 1.03rem;
            max-width: 860px;
        }

        .hero-badges {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-top: 18px;
        }

        .hero-badge {
            border: 1px solid rgba(255,255,255,0.20);
            background: rgba(255,255,255,0.10);
            color: rgba(255,255,255,0.95);
            border-radius: 999px;
            padding: 7px 12px;
            font-size: 0.82rem;
            font-weight: 650;
        }

        .app-card {
            background: var(--brand-card);
            border: 1px solid var(--brand-border);
            border-radius: 22px;
            padding: 18px 20px;
            box-shadow: 0 12px 34px rgba(15, 23, 42, 0.06);
            margin-bottom: 16px;
        }

        .app-card h3 {
            margin: 0 0 6px 0;
            font-size: 1.05rem;
            color: var(--brand-text);
        }

        .app-card p {
            margin: 0;
            color: var(--muted-text);
            font-size: 0.92rem;
        }

        .step-card {
            background: #FFFFFF;
            border: 1px solid #E3E8EF;
            border-radius: 18px;
            padding: 14px 16px;
            min-height: 102px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
        }

        .step-number {
            width: 28px;
            height: 28px;
            border-radius: 10px;
            background: #E9F8FB;
            color: #007E9B;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-weight: 800;
            margin-bottom: 10px;
        }

        .step-title {
            font-weight: 800;
            color: #17202A;
            margin-bottom: 4px;
        }

        .step-text {
            color: #667085;
            font-size: 0.88rem;
            line-height: 1.36;
        }

        .status-ok {
            color: #027A48;
            background: #ECFDF3;
            border: 1px solid #ABEFC6;
            padding: 6px 10px;
            border-radius: 999px;
            font-size: 0.82rem;
            font-weight: 700;
            display: inline-block;
        }

        .status-warn {
            color: #B54708;
            background: #FFFAEB;
            border: 1px solid #FEDF89;
            padding: 6px 10px;
            border-radius: 999px;
            font-size: 0.82rem;
            font-weight: 700;
            display: inline-block;
        }

        .muted-small {
            color: #667085;
            font-size: 0.84rem;
            margin-top: 4px;
        }

        .section-title {
            font-size: 1.22rem;
            font-weight: 850;
            letter-spacing: -0.02em;
            color: #17202A;
            margin: 4px 0 8px 0;
        }

        .section-subtitle {
            color: #667085;
            font-size: 0.95rem;
            margin-bottom: 14px;
        }

        .success-box {
            border: 1px solid #ABEFC6;
            background: #ECFDF3;
            border-radius: 18px;
            padding: 16px 18px;
        }

        .error-box {
            border: 1px solid #FECDCA;
            background: #FEF3F2;
            border-radius: 18px;
            padding: 16px 18px;
            color: #B42318;
        }

        .stButton > button {
            border-radius: 14px;
            font-weight: 800;
            border: 1px solid #062B49;
            background: #062B49;
            color: white;
            padding: 0.65rem 1.1rem;
        }

        .stDownloadButton > button {
            border-radius: 14px;
            font-weight: 800;
            background: #027A48;
            color: white;
            border: 1px solid #027A48;
            padding: 0.65rem 1.1rem;
        }

        div[data-testid="stFileUploader"] section {
            border-radius: 18px;
            border: 1.5px dashed #B8C4D3;
            background: #FBFCFE;
        }

        div[data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid #E3E8EF;
            padding: 14px;
            border-radius: 18px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
        }

        @media (max-width: 780px) {
            .hero-grid {
                grid-template-columns: 1fr;
            }
            .logo-box {
                max-width: 150px;
            }
            .hero h1 {
                font-size: 1.7rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


aplicar_css()


# ============================================================
# Utilitários de interface
# ============================================================

def render_header() -> None:
    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-grid">
                <div class="logo-box">
                    <img src="{LOGO_URL}" alt="Logo">
                </div>
                <div>
                    <h1>{APP_TITLE}</h1>
                    <p>{APP_SUBTITLE}</p>
                    <div class="hero-badges">
                        <span class="hero-badge">Sell-in x Sell-out</span>
                        <span class="hero-badge">Volume ou Quantidade</span>
                        <span class="hero-badge">Prévia de arquivos</span>
                        <span class="hero-badge">Excel final formatado</span>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def card_inicio(titulo: str, texto: str) -> None:
    st.markdown(
        f"""
        <div class="app-card">
            <h3>{titulo}</h3>
            <p>{texto}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def etapa(numero: int, titulo: str, texto: str) -> str:
    return f"""
    <div class="step-card">
        <div class="step-number">{numero}</div>
        <div class="step-title">{titulo}</div>
        <div class="step-text">{texto}</div>
    </div>
    """


def render_steps() -> None:
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(etapa(1, "Selecione o modo", "Escolha entre estudo individual, dois sell-outs, dash ou comparação."), unsafe_allow_html=True)
    c2.markdown(etapa(2, "Envie os arquivos", "Carregue as bases necessárias e confira a prévia das primeiras linhas."), unsafe_allow_html=True)
    c3.markdown(etapa(3, "Configure o cálculo", "Defina métrica, agrupamento, fabricante e abas desejadas no Excel."), unsafe_allow_html=True)
    c4.markdown(etapa(4, "Gere o resultado", "O app processa o motor original e libera o Excel para download."), unsafe_allow_html=True)


@st.cache_resource(show_spinner=False)
def carregar_motor():
    return importlib.import_module(ENGINE_MODULE)


def motor_disponivel() -> Tuple[bool, Optional[object], str]:
    try:
        mod = carregar_motor()
        return True, mod, ""
    except Exception as exc:
        return False, None, str(exc)


def safe_call(func, *args, **kwargs):
    """
    Chama funções do motor aceitando versões com assinaturas diferentes.
    Kwargs não existentes na assinatura são ignorados.
    """
    assinatura = inspect.signature(func)
    kwargs_filtrados = {k: v for k, v in kwargs.items() if k in assinatura.parameters}
    return func(*args, **kwargs_filtrados)


def extensao_upload(uploaded_file) -> str:
    if uploaded_file is None:
        return ".xlsx"
    nome = uploaded_file.name or ""
    ext = Path(nome).suffix.lower()
    return ext if ext else ".xlsx"


def salvar_upload(uploaded_file, pasta: Path, nome_base: str) -> str:
    ext = extensao_upload(uploaded_file)
    destino = pasta / f"{nome_base}{ext}"
    destino.write_bytes(uploaded_file.getvalue())
    return str(destino)


def nome_saida(prefixo: str) -> str:
    data = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefixo}_{data}.xlsx"


def arquivo_status(uploaded_file) -> str:
    if uploaded_file is None:
        return '<span class="status-warn">Pendente</span>'
    tamanho_mb = len(uploaded_file.getvalue()) / (1024 * 1024)
    return f'<span class="status-ok">Carregado · {tamanho_mb:.2f} MB</span>'


def render_file_status(label: str, uploaded_file) -> None:
    nome = uploaded_file.name if uploaded_file is not None else "Nenhum arquivo selecionado"
    st.markdown(
        f"""
        <div class="app-card">
            <h3>{label}</h3>
            <p>{nome}</p>
            <div class="muted-small">{arquivo_status(uploaded_file)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def carregar_preview_bytes(nome: str, conteudo: bytes, linhas: int = 12) -> Tuple[pd.DataFrame, str]:
    """
    Lê uma amostra leve do arquivo, sem aplicar a lógica de negócio.
    Para Excel, lê a primeira aba como linhas brutas.
    Para CSV/TXT, tenta separadores e encodings comuns.
    """
    import io

    ext = Path(nome or "").suffix.lower()

    if ext in {".xlsx", ".xlsm", ".xls"}:
        try:
            df = pd.read_excel(io.BytesIO(conteudo), nrows=linhas, header=None, dtype=str)
            df.columns = [f"Coluna {i + 1}" for i in range(len(df.columns))]
            return df.fillna(""), "Excel · primeira aba · leitura bruta"
        except Exception as exc:
            return pd.DataFrame({"Erro na prévia": [str(exc)]}), "Erro"

    encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]
    separadores = [",", ";", "\t", "|"]

    ultimo_erro = ""
    for encoding in encodings:
        for sep in separadores:
            try:
                df = pd.read_csv(
                    io.BytesIO(conteudo),
                    nrows=linhas,
                    sep=sep,
                    encoding=encoding,
                    dtype=str,
                    engine="python",
                )
                if len(df.columns) > 1:
                    return df.fillna(""), f"CSV/TXT · sep='{sep}' · encoding={encoding}"
            except Exception as exc:
                ultimo_erro = str(exc)

    try:
        texto = conteudo.decode("latin1", errors="replace").splitlines()[:linhas]
        return pd.DataFrame({"Linhas brutas": texto}), "Texto bruto"
    except Exception:
        return pd.DataFrame({"Erro na prévia": [ultimo_erro or "Não foi possível ler a prévia."]}), "Erro"


def render_preview(uploaded_file, titulo: str, linhas: int) -> None:
    if uploaded_file is None:
        return

    with st.expander(f"Prévia de linhas · {titulo}", expanded=False):
        df_prev, detalhe = carregar_preview_bytes(uploaded_file.name, uploaded_file.getvalue(), linhas)
        st.caption(detalhe)
        st.dataframe(df_prev, use_container_width=True, height=min(420, 82 + 28 * max(len(df_prev), 3)))


def download_arquivo_lateral(caminho: Path, rotulo: str, mime: str) -> None:
    if caminho.exists():
        st.download_button(
            rotulo,
            data=caminho.read_bytes(),
            file_name=caminho.name,
            mime=mime,
            use_container_width=True,
        )
    else:
        st.caption(f"Arquivo não localizado: {caminho.name}")


def seletor_congelado_opcional(prefixo_key: str, linhas_preview: int, texto_ajuda: str):
    usar_congelado = st.checkbox(
        "Usar Congelado opcional",
        value=False,
        key=f"{prefixo_key}_usar_congelado",
        help=texto_ajuda,
    )
    if not usar_congelado:
        st.caption("Congelado não será usado neste processamento. Marque a opção acima somente quando precisar apoiar o mapeamento por SKU/EAN.")
        return None

    with st.expander("Carregar Congelado opcional", expanded=True):
        congelado = st.file_uploader(
            "Congelado",
            type=["xlsx", "xlsm", "xls", "csv", "txt", "tsv"],
            key=f"{prefixo_key}_congelado",
        )
        render_file_status("Congelado opcional", congelado)
        render_preview(congelado, "Congelado opcional", linhas_preview)
        return congelado


def seletor_configuracoes(prefixo_key: str = "") -> Tuple[str, str, str]:
    c1, c2, c3 = st.columns([1.0, 1.0, 1.4])

    with c1:
        metrica_label = st.selectbox(
            "Cálculo da cobertura",
            list(METRICAS.keys()),
            index=0,
            key=f"{prefixo_key}_metrica",
            help="Volume usa Vendas_em_volume. Quantidade usa Qtd_de_Vendas. Volume variável aplica a regra de volume variável do motor.",
        )

    with c2:
        nivel_label = st.selectbox(
            "Agrupar por",
            list(NIVEIS.keys()),
            index=0,
            key=f"{prefixo_key}_nivel",
        )

    with c3:
        fabricante = st.text_input(
            "Filtro opcional de fabricante",
            value="",
            key=f"{prefixo_key}_fabricante",
            placeholder="Ex.: Camil, Softys, BAT...",
            help="Use somente quando quiser limitar o Sell-out por fabricante.",
        ).strip()

    return METRICAS[metrica_label], NIVEIS[nivel_label], fabricante


def seletor_opcoes_saida(prefixo_key: str = "") -> Dict[str, bool]:
    st.markdown('<div class="section-title">Saídas do Excel</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Marque as abas que devem ser geradas. O padrão já vem com as principais saídas ligadas.</div>', unsafe_allow_html=True)

    opcoes = {}

    with st.expander("Configurar abas do arquivo final", expanded=False):
        colunas = st.columns(2)
        for i, (chave, rotulo) in enumerate(OPCOES_SAIDA.items()):
            with colunas[i % 2]:
                opcoes[chave] = st.checkbox(
                    rotulo,
                    value=OPCOES_PADRAO.get(chave, True),
                    key=f"{prefixo_key}_opcao_{chave}",
                )

    return opcoes


def painel_lateral(ok_motor: bool, erro_motor: str) -> None:
    with st.sidebar:
        st.image(LOGO_URL, width=112)
        st.markdown("### Painel de execução")
        st.caption("Geração do Estudo de Cobertura em Excel.")

        if not ok_motor:
            st.error("Motor não carregado.")
            st.caption(erro_motor)

        st.divider()
        st.markdown("#### Manual de Uso")
        st.caption("Baixe o manual atualizado com o fluxo do Streamlit, parâmetros e leitura dos resultados.")
        pasta_app = Path(__file__).resolve().parent
        download_arquivo_lateral(
            pasta_app / MANUAL_FILE,
            "Baixar Manual de Uso",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.divider()
        st.markdown("#### Boas práticas")
        with st.expander("Ver boas práticas", expanded=True):
            st.caption("• Use Sell-in mensal, sem MAT/YTD agregado como mês.")
            st.caption("• Escolha a mesma métrica entre Sell-in e Sell-out: volume ou quantidade.")
            st.caption("• Use filtro de fabricante para reduzir arquivos grandes.")
            st.caption("• Use Congelado somente quando precisar apoiar o mapeamento por SKU/EAN.")
            st.caption("• Revise a aba Avisos antes de interpretar a cobertura final.")
        download_arquivo_lateral(
            pasta_app / BOAS_PRATICAS_FILE,
            "Baixar Boas Práticas",
            "text/markdown",
        )

        st.divider()
        if st.button("Limpar cache da interface"):
            st.cache_data.clear()
            st.cache_resource.clear()
            st.rerun()

def log_streamlit_factory(status_placeholder, progress_bar, log_area):
    linhas_log = []

    def log(msg: str, pct: Optional[float] = None):
        texto = str(msg)
        linhas_log.append(texto)
        status_placeholder.info(texto)

        if pct is not None:
            try:
                valor = max(0, min(100, int(round(float(pct)))))
                progress_bar.progress(valor)
            except Exception:
                pass

        with log_area:
            st.code("\n".join(linhas_log[-18:]), language="text")

    return log


def executar_e_preparar_download(label_botao: str, nome_download: str, executar_callback) -> None:
    st.divider()
    col_btn, col_info = st.columns([0.32, 0.68])

    with col_btn:
        gerar = st.button(label_botao, type="primary", use_container_width=True)

    with col_info:
        st.caption("Após gerar, o arquivo ficará disponível para download logo abaixo.")

    if not gerar:
        return

    status_placeholder = st.empty()
    progress_bar = st.progress(0)
    log_area = st.empty()
    log_callback = log_streamlit_factory(status_placeholder, progress_bar, log_area)

    try:
        with st.spinner("Processando estudo..."):
            bytes_saida = executar_callback(log_callback)

        progress_bar.progress(100)
        status_placeholder.markdown(
            """
            <div class="success-box">
                <strong>Arquivo gerado com sucesso.</strong><br>
                Faça o download do Excel abaixo.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.download_button(
            "Baixar Excel gerado",
            data=bytes_saida,
            file_name=nome_download,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as exc:
        status_placeholder.markdown(
            f"""
            <div class="error-box">
                <strong>Erro ao gerar o estudo.</strong><br>
                {str(exc)}
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.exception(exc)


def requeridos_ok(arquivos: Iterable) -> bool:
    return all(arq is not None for arq in arquivos)


# ============================================================
# Modos de execução
# ============================================================

def tela_estudo_individual(motor, linhas_preview: int, output_options: Dict[str, bool]) -> None:
    st.markdown('<div class="section-title">Estudo individual</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Use um Sell-in e um Sell-out. O Congelado fica recolhido e só aparece se você ativar a opção.</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        sellin = st.file_uploader("Sell-in", type=["xlsx", "xlsm", "xls"], key="estudo_sellin")
    with c2:
        sellout = st.file_uploader("Sell-out", type=["xlsx", "xlsm", "xls", "csv", "txt", "tsv"], key="estudo_sellout")

    s1, s2 = st.columns(2)
    with s1:
        render_file_status("Sell-in", sellin)
    with s2:
        render_file_status("Sell-out", sellout)

    render_preview(sellin, "Sell-in", linhas_preview)
    render_preview(sellout, "Sell-out", linhas_preview)

    congelado = seletor_congelado_opcional(
        "estudo",
        linhas_preview,
        "Use apenas se precisar que a categoria dos SKUs seja apoiada pelo Congelado.",
    )

    metrica, nivel, fabricante = seletor_configuracoes("estudo")

    if not requeridos_ok([sellin, sellout]):
        st.warning("Envie pelo menos o Sell-in e o Sell-out para liberar a geração.")
        return

    def executar(log_callback):
        with tempfile.TemporaryDirectory(prefix="streamlit_cobertura_") as tmpdir:
            tmp = Path(tmpdir)
            p_sellin = salvar_upload(sellin, tmp, "sellin")
            p_sellout = salvar_upload(sellout, tmp, "sellout")
            p_congelado = salvar_upload(congelado, tmp, "congelado") if congelado is not None else ""
            p_saida = tmp / nome_saida("estudo_cobertura")

            resultado = safe_call(
                motor.executar_estudo,
                p_sellin,
                p_sellout,
                str(p_saida),
                metrica,
                nivel,
                fabricante,
                congelado_path=p_congelado,
                log_callback=log_callback,
                output_options=output_options,
            )

            return Path(resultado).read_bytes()

    executar_e_preparar_download("Gerar estudo individual", nome_saida("estudo_cobertura"), executar)


def tela_estudo_dois_sellouts(motor, linhas_preview: int, output_options: Dict[str, bool]) -> None:
    st.markdown('<div class="section-title">Estudo com 2 Sell-outs</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Use o mesmo Sell-in para comparar Sell-out 2.0 x Sell-out 3.0 em um único Excel final.</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        sellin = st.file_uploader("Sell-in", type=["xlsx", "xlsm", "xls"], key="dois_sellin")
    with c2:
        sellout20 = st.file_uploader("Sell-out 2.0", type=["xlsx", "xlsm", "xls", "csv", "txt", "tsv"], key="dois_sellout20")
    with c3:
        sellout30 = st.file_uploader("Sell-out 3.0", type=["xlsx", "xlsm", "xls", "csv", "txt", "tsv"], key="dois_sellout30")

    s1, s2, s3 = st.columns(3)
    with s1:
        render_file_status("Sell-in", sellin)
    with s2:
        render_file_status("Sell-out 2.0", sellout20)
    with s3:
        render_file_status("Sell-out 3.0", sellout30)

    render_preview(sellin, "Sell-in", linhas_preview)
    render_preview(sellout20, "Sell-out 2.0", linhas_preview)
    render_preview(sellout30, "Sell-out 3.0", linhas_preview)

    congelado = seletor_congelado_opcional(
        "dois",
        linhas_preview,
        "Use apenas se quiser apoiar a categoria dos SKUs pelo Congelado nos dois Sell-outs.",
    )

    metrica, nivel, fabricante = seletor_configuracoes("dois")
    gerar_top20 = st.checkbox(
        "Gerar TOP 20 SKU por UF/Canal",
        value=output_options.get("top20_sku_canal_uf", False),
        key="dois_top20",
    )

    if not requeridos_ok([sellin, sellout20, sellout30]):
        st.warning("Envie Sell-in, Sell-out 2.0 e Sell-out 3.0 para liberar a geração.")
        return

    def executar(log_callback):
        with tempfile.TemporaryDirectory(prefix="streamlit_cobertura_2out_") as tmpdir:
            tmp = Path(tmpdir)
            p_sellin = salvar_upload(sellin, tmp, "sellin")
            p_sellout20 = salvar_upload(sellout20, tmp, "sellout20")
            p_sellout30 = salvar_upload(sellout30, tmp, "sellout30")
            p_congelado = salvar_upload(congelado, tmp, "congelado") if congelado is not None else ""
            p_saida = tmp / nome_saida("comparacao_2_sellouts")

            resultado = safe_call(
                motor.executar_estudo_dois_sellouts,
                p_sellin,
                p_sellout20,
                p_sellout30,
                str(p_saida),
                metrica,
                nivel,
                fabricante,
                congelado_path=p_congelado,
                log_callback=log_callback,
                gerar_top20_sku_canal_uf=gerar_top20,
                output_options={**output_options, "top20_sku_canal_uf": gerar_top20},
            )

            return Path(resultado).read_bytes()

    executar_e_preparar_download("Gerar comparação 2 Sell-outs", nome_saida("comparacao_2_sellouts"), executar)


def tela_dash(motor, linhas_preview: int, output_options: Dict[str, bool]) -> None:
    st.markdown('<div class="section-title">Cobertura Dash</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Use Sell-in, Vendas UF e SKU. Vendas SKU e Congelado são opcionais conforme a necessidade da análise.</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        sellin = st.file_uploader("Sell-in", type=["xlsx", "xlsm", "xls"], key="dash_sellin")
    with c2:
        vendas_uf = st.file_uploader("Vendas UF", type=["xlsx", "xlsm", "xls", "csv", "txt", "tsv"], key="dash_vendas_uf")
    with c3:
        vendas_sku = st.file_uploader("Vendas SKU opcional", type=["xlsx", "xlsm", "xls", "csv", "txt", "tsv"], key="dash_vendas_sku")

    c4, c5 = st.columns([1, 1])
    with c4:
        sku = st.file_uploader("Arquivo SKU", type=["xlsx", "xlsm", "xls", "csv", "txt", "tsv"], key="dash_sku")
    with c5:
        st.info("Congelado é opcional. Ative abaixo somente se precisar usar Marca/Fabricante/Est Mer 7 da Base Congelada.")

    s1, s2, s3, s4 = st.columns(4)
    with s1:
        render_file_status("Sell-in", sellin)
    with s2:
        render_file_status("Vendas UF", vendas_uf)
    with s3:
        render_file_status("Vendas SKU opcional", vendas_sku)
    with s4:
        render_file_status("SKU", sku)

    render_preview(sellin, "Sell-in", linhas_preview)
    render_preview(vendas_uf, "Vendas UF", linhas_preview)
    render_preview(vendas_sku, "Vendas SKU opcional", linhas_preview)
    render_preview(sku, "SKU", linhas_preview)

    congelado = seletor_congelado_opcional(
        "dash",
        linhas_preview,
        "Use se precisar complementar Marca, Fabricante, categoria congelada ou Est Mer 7 a partir da Base Congelada.",
    )

    metrica, nivel, fabricante = seletor_configuracoes("dash")

    if not requeridos_ok([sellin, vendas_uf, sku]):
        st.warning("Envie Sell-in, Vendas UF e SKU para liberar a geração do modo Dash. O Congelado é opcional.")
        return

    def executar(log_callback):
        with tempfile.TemporaryDirectory(prefix="streamlit_cobertura_dash_") as tmpdir:
            tmp = Path(tmpdir)
            p_sellin = salvar_upload(sellin, tmp, "sellin")
            p_vendas_uf = salvar_upload(vendas_uf, tmp, "vendas_uf")
            p_vendas_sku = salvar_upload(vendas_sku, tmp, "vendas_sku") if vendas_sku is not None else None
            p_sku = salvar_upload(sku, tmp, "sku")
            p_congelado = salvar_upload(congelado, tmp, "congelado") if congelado is not None else ""
            p_saida = tmp / nome_saida("cobertura_dash")

            resultado = safe_call(
                motor.executar_cobertura_dash,
                p_sellin,
                p_vendas_uf,
                p_sku,
                p_congelado,
                str(p_saida),
                metrica,
                nivel,
                fabricante,
                vendas_sku_path=p_vendas_sku,
                log_callback=log_callback,
                output_options=output_options,
            )

            return Path(resultado).read_bytes()

    executar_e_preparar_download("Gerar cobertura Dash", nome_saida("cobertura_dash"), executar)


def tela_comparacao(motor, linhas_preview: int, output_options: Dict[str, bool]) -> None:
    st.markdown('<div class="section-title">Comparação de estudos já gerados</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Use dois arquivos de estudo já processados para gerar uma comparação consolidada.</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        estudo20 = st.file_uploader("Estudo 2.0", type=["xlsx", "xlsm"], key="comp_estudo20")
    with c2:
        estudo30 = st.file_uploader("Estudo 3.0", type=["xlsx", "xlsm"], key="comp_estudo30")

    s1, s2 = st.columns(2)
    with s1:
        render_file_status("Estudo 2.0", estudo20)
    with s2:
        render_file_status("Estudo 3.0", estudo30)

    render_preview(estudo20, "Estudo 2.0", linhas_preview)
    render_preview(estudo30, "Estudo 3.0", linhas_preview)

    gerar_top20 = st.checkbox(
        "Gerar TOP 20 SKU por UF/Canal",
        value=output_options.get("top20_sku_canal_uf", False),
        key="comp_top20",
    )

    if not requeridos_ok([estudo20, estudo30]):
        st.warning("Envie os dois estudos para liberar a comparação.")
        return

    def executar(log_callback):
        with tempfile.TemporaryDirectory(prefix="streamlit_comparacao_") as tmpdir:
            tmp = Path(tmpdir)
            p_estudo20 = salvar_upload(estudo20, tmp, "estudo20")
            p_estudo30 = salvar_upload(estudo30, tmp, "estudo30")
            p_saida = tmp / nome_saida("comparacao_estudos")

            resultado = safe_call(
                motor.gerar_comparacao_estudos,
                p_estudo20,
                p_estudo30,
                str(p_saida),
                log_callback=log_callback,
                gerar_top20_sku_canal_uf=gerar_top20,
                output_options={**output_options, "top20_sku_canal_uf": gerar_top20},
            )

            return Path(resultado).read_bytes()

    executar_e_preparar_download("Gerar comparação", nome_saida("comparacao_estudos"), executar)


# ============================================================
# App principal
# ============================================================

def main() -> None:
    render_header()
    render_steps()

    ok_motor, motor, erro_motor = motor_disponivel()
    painel_lateral(ok_motor, erro_motor)

    st.markdown("---")

    if not ok_motor:
        st.error(
            f"Não consegui carregar o motor `{ENGINE_MODULE}.py`. "
            "Coloque este arquivo Streamlit na mesma pasta do código-base ou defina a variável COBERTURA_ENGINE_MODULE."
        )
        st.code(erro_motor, language="text")
        return

    col_a, col_b = st.columns([1.4, 1.0])
    with col_a:
        modo = st.radio(
            "Modo de geração",
            ["Estudo individual", "2 Sell-outs", "Cobertura Dash", "Comparação"],
            horizontal=True,
        )
    with col_b:
        linhas_preview = st.slider("Linhas de demonstração", 5, 30, 12, 1)

    card_inicio(
        "Prévia antes do processamento",
        "As linhas de demonstração são apenas uma leitura bruta inicial do arquivo. A validação real continua sendo feita pelo motor do estudo.",
    )

    output_options = seletor_opcoes_saida("global")

    st.markdown("---")

    if modo == "Estudo individual":
        tela_estudo_individual(motor, linhas_preview, output_options)
    elif modo == "2 Sell-outs":
        tela_estudo_dois_sellouts(motor, linhas_preview, output_options)
    elif modo == "Cobertura Dash":
        tela_dash(motor, linhas_preview, output_options)
    else:
        tela_comparacao(motor, linhas_preview, output_options)


if __name__ == "__main__":
    main()
