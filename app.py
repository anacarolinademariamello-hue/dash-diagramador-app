import io
import streamlit as st
from diagramador import diagramar

st.set_page_config(page_title="Diagramador de Simulados", page_icon="📄", layout="centered")

st.title("📄 Diagramador de Simulados")
st.write(
    "Envie o arquivo Word **cru** (sem diagramação) e baixe a versão já formatada: "
    "o cabeçalho/rodapé originais do arquivo (banner e logo) são preservados, uma faixa "
    "azul com o título da matéria é adicionada repetindo em todas as páginas, as questões "
    "e alternativas ficam em destaque, e o Gabarito Simplificado é gerado automaticamente."
)

with st.expander("⚙️ Personalizar título da faixa azul (opcional)"):
    titulo_custom = st.text_input("Título (deixe vazio para detectar automaticamente)")
    subtitulo_custom = st.text_input("Ementa / subtítulo (opcional)")

arquivo = st.file_uploader("Arquivo .docx cru", type=["docx"])

if arquivo is not None:
    if st.button("Diagramar", type="primary"):
        with st.spinner("Diagramando..."):
            try:
                doc, n_gabarito = diagramar(
                    arquivo,
                    titulo=titulo_custom or None,
                    subtitulo=subtitulo_custom or None,
                )
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.success(f"Pronto! {n_gabarito} questões detectadas no gabarito.")

                nome_saida = arquivo.name.replace(".docx", "_diagramado.docx")
                st.download_button(
                    "⬇️ Baixar arquivo diagramado",
                    data=buffer,
                    file_name=nome_saida,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"Não consegui diagramar este arquivo: {e}")
                st.info(
                    "Confira se o documento segue o padrão esperado (Blocos, Questões numeradas, "
                    "alternativas A) a E), Gabarito Comentado com 'Alternativa correta: X')."
                )

st.divider()
st.caption(
    "Padrão esperado no arquivo de entrada: títulos de bloco (\"BLOCO 1 — ...\"), "
    "questões (\"Questão 1 ...\"), alternativas (\"A) ...\" a \"E) ...\") e seções de "
    "gabarito comentado (\"Questão 1 — Alternativa correta: D\")."
)
