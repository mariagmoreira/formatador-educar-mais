# Formatador de Artigos — Revista Educar Mais

Ferramenta web simples: você faz upload do .docx do autor e baixa o artigo
já formatado conforme as Diretrizes da revista (Times New Roman 12,
espaçamento duplo, papel A4, margens de 3 cm, páginas numeradas, estrutura
título/autores/resumos/seções/referências).

## Como rodar no seu computador

1. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```
2. Rode a ferramenta:
   ```
   streamlit run app.py
   ```
3. O navegador abre automaticamente em `http://localhost:8501`.

## Como publicar como site (grátis, sem servidor próprio)

1. Crie um repositório no GitHub com estes 3 arquivos (`app.py`,
   `formatter_engine.py`, `requirements.txt`).
2. Acesse https://streamlit.io/cloud, conecte sua conta do GitHub e
   selecione o repositório.
3. Em poucos minutos você tem uma URL pública (ex.:
   `https://educar-mais-formatador.streamlit.app`) que qualquer editor
   da revista pode usar direto do navegador, sem instalar nada.

## O que a ferramenta faz automaticamente

- Reconhece título (PT/EN/ES), autores, RESUMO/ABSTRACT/RESUMEN com
  palavras-chave, seções numeradas do corpo e REFERÊNCIAS.
- Aplica a formatação exigida pelas Diretrizes: fonte, espaçamento,
  margens, tamanho de papel, numeração de página.
- Insere a logo da revista no topo e o ícone do ORCID ao lado de cada
  autor, usando as mesmas imagens do `Modelo.docx` (incluído nesta pasta —
  se você atualizar o modelo da revista, é só substituir esse arquivo).

## O que ainda exige revisão humana

- Contagem de caracteres do resumo (550–1.200).
- Anonimização do manuscrito (remoção de identificação dos autores).
- Conformidade das referências com a ABNT.
- Autenticação do ORCID.
- Artigos com estrutura muito diferente do padrão podem não ser
  reconhecidos corretamente — nesse caso a ferramenta avisa o motivo.
