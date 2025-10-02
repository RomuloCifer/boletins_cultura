import os, re
import pandas as pd
from docx import Document
# Se quiser converter para PDF via Word, descomente:
# from docx2pdf import convert

ARQUIVO_RESULTADOS = "resultados_boletim.xlsx"
MODELO = "modelo_boletim.docx"
PASTA_SAIDA = "boletins_pdf"

# quais chaves esperamos encontrar no Excel e no modelo
CHAVES = ["Aluno","Turma","Nivel","Professor",
          "Comunicacao","Compreensao","Interesse","Colaboracao","Engajamento"]

def _replace_all(texto: str, dados: dict) -> str:
    """Substitui todos os <<CAMPO>> presentes em 'texto' usando 'dados'."""
    for k in CHAVES:
        val = dados.get(k, "")
        if pd.isna(val):
            val = ""
        texto = texto.replace(f"<<{k}>>", str(val))
    return texto

def _replace_in_paragraph(par, dados):
    # Concatena todos os runs, substitui e regrava com segurança
    old = "".join(run.text for run in par.runs) or par.text
    new = _replace_all(old, dados)
    if new != old:
        if not par.runs:
            par.add_run()
        par.runs[0].text = new
        for run in par.runs[1:]:
            run.text = ""

def _replace_in_cell(cell, dados):
    # Garante ao menos um parágrafo
    if not cell.paragraphs:
        cell.add_paragraph("")
    for p in list(cell.paragraphs):
        _replace_in_paragraph(p, dados)

def substituir_texto(doc: Document, dados: dict):
    # Parágrafos fora de tabelas
    for p in doc.paragraphs:
        _replace_in_paragraph(p, dados)
    # Dentro das tabelas
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                _replace_in_cell(cell, dados)

def safe_filename(s: str) -> str:
    s = str(s)
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)  # caracteres inválidos no Windows
    return s.strip()

def gerar_boletins():
    df = pd.read_excel(ARQUIVO_RESULTADOS)

    # Harmoniza coluna "Nome" -> "Aluno" caso venha assim
    if "Nome" in df.columns and "Aluno" not in df.columns:
        df = df.rename(columns={"Nome": "Aluno"})

    # Cria/limpa pasta de saída
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    for arq in os.listdir(PASTA_SAIDA):
        try:
            os.remove(os.path.join(PASTA_SAIDA, arq))
        except:
            pass

    for _, row in df.iterrows():
        dados = {k: row[k] for k in CHAVES if k in row.index}
        doc = Document(MODELO)
        substituir_texto(doc, dados)

        # Nome do arquivo: Aluno_Turma_boletim.docx
        nome = safe_filename(dados.get("Aluno", "Aluno"))
        turma = safe_filename(dados.get("Turma", "Turma"))
        nome_docx = f"{nome}_{turma}_boletim.docx"
        caminho_docx = os.path.join(PASTA_SAIDA, nome_docx)

        doc.save(caminho_docx)

        # Para PDF via Word, descomente a linha abaixo e instale/tenha Word:
        # convert(caminho_docx, os.path.join(PASTA_SAIDA, f"{nome}_{turma}_boletim.pdf"))

    print(f"✅ Boletins gerados em: {PASTA_SAIDA}")

if __name__ == "__main__":
    gerar_boletins()
