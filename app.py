import base64
import streamlit as st
from formatter_engine import convert

st.set_page_config(
    page_title="Formatador — Revista Educar Mais",
    page_icon="app_logo.png",
    layout="centered",
)

# ---- Estilo (cores da revista + componentes mais arredondados) ----
st.markdown(
    """
    <style>
    .block-container { padding-top: 2.5rem; max-width: 760px; }

    .em-header { display: flex; align-items: center; gap: 1rem; margin-bottom: .25rem; }
    .em-header img { width: 64px; }
    .em-header .em-title h1 {
        font-size: 1.5rem; margin: 0; color: #2B2B2B; font-weight: 700;
    }
    .em-header .em-title p {
        margin: .15rem 0 0 0; color: #757575; font-size: .95rem;
    }

    hr.em-rule {
        border: none; height: 3px; background: #FF7F26;
        border-radius: 2px; margin: 1.1rem 0 1.6rem 0;
    }

    div[data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #FFBF8F; border-radius: 12px; background: #FFF8F2;
    }

    .stButton > button, .stDownloadButton > button {
        background: #FF7F26; color: white; border: none; border-radius: 8px;
        font-weight: 600; padding: .6rem 1.4rem;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #ED7D31; color: white;
    }

    .em-checklist { color: #757575; font-size: .92rem; line-height: 1.6; }
    .em-footer { text-align: center; color: #B0B0B0; font-size: .8rem; margin-top: 3rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---- Cabeçalho com logo ----
try:
    with open("app_logo.png", "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode()
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" />'
except FileNotFoundError:
    logo_html = ""

st.markdown(
    f"""
    <div class="em-header">
        {logo_html}
        <div class="em-title">
            <h1>Formatador de Artigos</h1>
            <p>Revista Educar Mais</p>
        </div>
    </div>
    <hr class="em-rule" />
    """,
    unsafe_allow_html=True,
)

st.write(
    "Envie o artigo **.docx** recebido do autor e baixe uma versão já formatada "
    "no layout oficial da revista — logo, cabeçalho, ORCID, notas de rodapé, "
    "tabelas, figuras e referências, tudo automático."
)

uploaded = st.file_uploader("Artigo do autor (.docx)", type=["docx"], label_visibility="collapsed")

if uploaded is not None:
    with open("_input.docx", "wb") as f:
        f.write(uploaded.read())

    with st.spinner("Formatando o artigo..."):
        try:
            blocks = convert("_input.docx", "_output.docx", modelo_path="Modelo.docx")
        except Exception as e:
            blocks = None
            error = e

    if blocks is not None:
        st.success("✅ Artigo formatado com sucesso!")

        col1, col2, col3 = st.columns(3)
        col1.metric("Autores", len(blocks["authors"]))
        col2.metric("Seções", sum(1 for it in blocks["body"] if it[0] == "heading"))
        col3.metric("Referências", len(blocks["references"]))

        with st.expander("Conferir o que foi identificado no artigo"):
            st.write("**Títulos encontrados:**")
            for t in blocks["titles"]:
                st.write("-", t)
            st.write("**Autores:**")
            for name, orcid_url, bio in blocks["authors"]:
                st.write("-", name)
            st.write(f"**Resumos:** {', '.join(a['label'] for a in blocks['abstracts'])}")
            st.write(f"**Tabelas/quadros no corpo:** {sum(1 for it in blocks['body'] if it[0] == 'table')}")
            st.write(f"**Figuras no corpo:** {sum(1 for it in blocks['body'] if it[0] == 'figure')}")

        st.markdown(
            """
            <div class="em-checklist">
            ⚠️ <b>Confira sempre antes de publicar:</b> contagem de caracteres do resumo
            (550–1.200), anonimização do texto, ORCID autenticado e normas ABNT das
            referências. O formatador cuida do layout, não dessas validações.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.write("")

        with open("_output.docx", "rb") as f:
            st.download_button(
                "⬇️  Baixar artigo formatado (.docx)",
                data=f.read(),
                file_name=f"formatado_{uploaded.name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    else:
        st.error(
            "Não consegui reconhecer a estrutura do artigo automaticamente. "
            f"Detalhe técnico: {error}"
        )
        st.info(
            "Isso costuma acontecer quando o autor não seguiu a ordem "
            "título → autores → resumos → seções numeradas → referências."
        )

st.markdown(
    '<div class="em-footer">Revista Educar Mais · Formatador interno</div>',
    unsafe_allow_html=True,
)
