import os, re
import pandas as pd
from docx import Document
# Se quiser converter para PDF via Word, descomente:
# from docx2pdf import convert

ARQUIVO_RESULTADOS = "resultados_boletim.xlsx"
PASTA_MODELOS = "Modelos"   # nome da pasta igual ao que está no seu diretório
PASTA_SAIDA = "boletins_pdf"

# Mapeamento de nível fixo -> modelo correspondente
MAPA_MODELOS = {
    "Lion Stars": "Modelo Boletim - Lion stars.docx",
    "Junior": "Modelo Boletim - Junior.docx",
    "Adultos": "Modelo Boletim - Adolescentes e Adultos.docx"  # único modelo para todos os adultos
}

def _replace_all(texto: str, dados: dict) -> str:
    for k, val in dados.items():
        if pd.isna(val):
            val = ""
        texto = texto.replace(f"<<{k}>>", str(val))
    return texto

def _replace_in_paragraph(par, dados):
    old = "".join(run.text for run in par.runs) or par.text
    new = _replace_all(old, dados)
    if new != old:
        if not par.runs:
            par.add_run()
        par.runs[0].text = new
        for run in par.runs[1:]:
            run.text = ""

def _replace_in_cell(cell, dados):
    if not cell.paragraphs:
        cell.add_paragraph("")
    for p in list(cell.paragraphs):
        _replace_in_paragraph(p, dados)

def substituir_texto(doc: Document, dados: dict):
    for p in doc.paragraphs:
        _replace_in_paragraph(p, dados)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                _replace_in_cell(cell, dados)

def safe_filename(s: str) -> str:
    s = str(s)
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    return s.strip()

def gerar_boletins():
    df = pd.read_excel(ARQUIVO_RESULTADOS)

    # Renomeia se vier "Nome" ao invés de "Aluno"
    if "Nome" in df.columns and "Aluno" not in df.columns:
        df = df.rename(columns={"Nome": "Aluno"})

    os.makedirs(PASTA_SAIDA, exist_ok=True)

    for _, row in df.iterrows():
        nivel = str(row.get("Nivel", "")).strip()

        # regra especial: qualquer Adultos usa o mesmo modelo
        if nivel.startswith("Adultos"):
            caminho_modelo = os.path.join(PASTA_MODELOS, MAPA_MODELOS["Adultos"])
        elif nivel in MAPA_MODELOS:
            caminho_modelo = os.path.join(PASTA_MODELOS, MAPA_MODELOS[nivel])
        else:
            print(f"⚠️ Nível {nivel} não tem modelo configurado, pulando {row['Aluno']}")
            continue

        if not os.path.exists(caminho_modelo):
            print(f"❌ Modelo não encontrado: {caminho_modelo}")
            continue

        dados = {col: row[col] for col in df.columns if col in row.index}
        doc = Document(caminho_modelo)
        substituir_texto(doc, dados)

        nome = safe_filename(dados.get("Aluno", "Aluno"))
        turma = safe_filename(dados.get("Turma", "Turma"))
        nome_docx = f"{nome}_{turma}_{nivel}.docx".replace("/", "-")
        caminho_docx = os.path.join(PASTA_SAIDA, nome_docx)

        doc.save(caminho_docx)

        # Se quiser gerar também PDF → descomente:
        # convert(caminho_docx, os.path.join(PASTA_SAIDA, f"{nome}_{turma}_{nivel}.pdf"))

        print(f"✅ Boletim gerado: {caminho_docx}")

if __name__ == "__main__":
    gerar_boletins()
