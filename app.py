import io
import streamlit as st
from formatter_engine import convert

st.set_page_config(page_title="Formatador — Revista Educar Mais", page_icon="📄")

st.title("📄 Formatador de Artigos — Revista Educar Mais")
st.write(
    "Envie o artigo .docx recebido do autor e baixe uma versão já formatada "
    "conforme as Diretrizes da revista (Times New Roman 12, espaçamento duplo, "
    "A4, margens de 3 cm, páginas numeradas)."
)

uploaded = st.file_uploader("Artigo do autor (.docx)", type=["docx"])

if uploaded is not None:
    with open("_input.docx", "wb") as f:
        f.write(uploaded.read())

    try:
        blocks = convert("_input.docx", "_output.docx", modelo_path="Modelo.docx")

        st.success("Artigo formatado com sucesso!")

        with st.expander("Conferir o que foi identificado no artigo"):
            st.write("**Títulos encontrados:**")
            for t in blocks["titles"]:
                st.write("-", t)
            st.write(f"**Autores encontrados:** {len(blocks['authors'])}")
            for name, bio in blocks["authors"]:
                st.write("-", name)
            st.write(f"**Resumos encontrados:** {', '.join(a['label'] for a in blocks['abstracts'])}")
            st.write(f"**Seções no corpo:** {sum(1 for it in blocks['body'] if it[0] == 'heading')}")
            st.write(f"**Tabelas/quadros no corpo:** {sum(1 for it in blocks['body'] if it[0] == 'table')}")
            st.write(f"**Figuras no corpo:** {sum(1 for it in blocks['body'] if it[0] == 'figure')}")
            st.write(f"**Referências encontradas:** {len(blocks['references'])}")

        st.warning(
            "⚠️ Confira sempre: contagem de caracteres do resumo (550–1.200), "
            "anonimização do texto, ORCID autenticado e normas ABNT das "
            "referências. O formatador cuida do layout, não dessas validações."
        )

        with open("_output.docx", "rb") as f:
            st.download_button(
                "⬇️ Baixar artigo formatado (.docx)",
                data=f.read(),
                file_name=f"formatado_{uploaded.name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    except Exception as e:
        st.error(
            "Não consegui reconhecer a estrutura do artigo automaticamente. "
            f"Detalhe técnico: {e}"
        )
        st.info(
            "Isso costuma acontecer quando o autor não seguiu a ordem "
            "título → autores → resumos → seções numeradas → referências."
        )
