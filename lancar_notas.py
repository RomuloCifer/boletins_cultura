import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd

# ====== CONFIGURAÇÕES ======
ARQUIVO_ALUNOS = "alunos.xlsx"

# Critérios específicos por nível
CRITERIOS_POR_NIVEL = {
    "Lion Stars": [
        "Comunicação oral",
        "Compreensão oral",
        "Interesse pelo processo de aprendizagem",
        "Colaboração com colegas",
        "Engajamento nas atividades de sala"
    ],
    "Junior": [
        "Comunicação oral",
        "Compreensão oral",
        "Comunicação escrita",
        "Compreensão escrita",
        "Interesse pelo processo de aprendizagem",
        "Colaboração com colegas",
        "Engajamento nas atividades de sala"
    ],
    "Adultos": [
        "Produção oral",
        "Produção escrita",
        "Progress Check"
    ]
}

# Mapeamento de critérios -> placeholders nos modelos Word
MAPEAMENTO_CHAVES = {
    # Lion Stars
    "Comunicação oral": "Comunicacao",
    "Compreensão oral": "Compreensao",
    "Interesse pelo processo de aprendizagem": "Interesse",
    "Colaboração com colegas": "Colaboracao",
    "Engajamento nas atividades de sala": "Engajamento",

    # Junior
    "Comunicação escrita": "ComunicacaoE",
    "Compreensão escrita": "CompreensaoEscrita",

    # Adultos
    "Produção oral": "ProducaoOral",
    "Produção escrita": "ProducaoE",
    "Progress Check": "ProgressCheck"
}

OPCOES = ["A", "B", "C", "D"]

LEGENDA = """(A) = Atingiu plenamente (100–90%)
(B) = Atingiu satisfatoriamente (89–75%)
(C) = Atingiu parcialmente (74–60%)
(D) (N/A) = Ainda não atingiu / Não há evidências (59% ou menos)"""

# Faixas de valores (mínimo, máximo) por letra
NOTAS_RANGE = {
    "A": (90, 100),
    "B": (75, 89),
    "C": (60, 74),
    "D": (0, 59)
}

class BoletimApp:
    def __init__(self, root, alunos, professor):
        self.root = root
        self.alunos = alunos
        self.professor = professor
        self.index = 0
        self.resultados = []

        self.root.title("Sistema de Boletins")
        self.root.geometry("650x600")

        self.label_nome = tk.Label(root, text="", font=("Arial", 14, "bold"))
        self.label_nome.pack(pady=10)

        self.frame_criterios = tk.Frame(root)
        self.frame_criterios.pack(pady=10, fill="x")

        self.combos = {}

        self.btn = tk.Button(root, text="Salvar e próximo", command=self.salvar)
        self.btn.pack(pady=20)

        self.label_legenda = tk.Label(root, text=LEGENDA, justify="left", font=("Arial", 10), anchor="w")
        self.label_legenda.pack(pady=10, fill="x")

        self.mostrar_aluno()

    def criar_campos(self, criterios):
        for widget in self.frame_criterios.winfo_children():
            widget.destroy()
        self.combos.clear()

        for criterio in criterios:
            frame = tk.Frame(self.frame_criterios)
            frame.pack(pady=5, fill="x")

            lbl = tk.Label(frame, text=criterio, width=40, anchor="w")
            lbl.pack(side="left")

            cb = ttk.Combobox(frame, values=OPCOES, state="readonly", width=5)
            cb.pack(side="left", padx=10)
            self.combos[criterio] = cb

    def mostrar_aluno(self):
        if self.index < len(self.alunos):
            aluno = self.alunos[self.index]
            nome = aluno.get("Nome", "")
            turma = aluno.get("Turma", "")
            nivel = aluno.get("Nivel", "")
            self.label_nome.config(text=f"Aluno: {nome}  |  Turma: {turma}  |  Nível: {nivel}")

            criterios = CRITERIOS_POR_NIVEL.get(nivel, CRITERIOS_POR_NIVEL["Lion Stars"])
            self.criar_campos(criterios)

        else:
            messagebox.showinfo("Fim", "Todos os alunos foram avaliados!")
            df = pd.DataFrame(self.resultados)
            df.to_excel("resultados_boletim.xlsx", index=False)
            self.root.quit()

    def salvar(self):
        aluno_row = self.alunos[self.index]
        nivel = aluno_row.get("Nivel", "")
        notas = {
            "Aluno": aluno_row.get("Nome", ""),
            "Turma": aluno_row.get("Turma", ""),
            "Nivel": nivel,
            "Professor": self.professor
        }

        valores_min = []
        valores_max = []

        for crit, cb in self.combos.items():
            val = cb.get()
            if val == "":
                messagebox.showwarning("Atenção", f"Selecione uma nota para '{crit}'")
                return
            chave = MAPEAMENTO_CHAVES.get(crit, crit)
            notas[chave] = val

            if nivel == "Adultos":
                min_val, max_val = NOTAS_RANGE[val]
                valores_min.append(min_val)
                valores_max.append(max_val)

        # Se for Adultos → gera nota sugerida como intervalo e pede nota manual
        if nivel == "Adultos" and valores_min:
            media_min = sum(valores_min) / len(valores_min)
            media_max = sum(valores_max) / len(valores_max)
            notas["NotaSugerida"] = f"{int(media_min)} - {int(media_max)}"

            nota_final = simpledialog.askinteger(
                "Nota Final",
                f"Nota sugerida para {aluno_row.get('Nome', '')}: {int(media_min)} – {int(media_max)}\nDigite a nota final:",
                minvalue=0, maxvalue=100
            )
            notas["Nota"] = nota_final if nota_final is not None else int((media_min + media_max) / 2)

        self.resultados.append(notas)
        self.index += 1
        self.mostrar_aluno()

if __name__ == "__main__":
    df_alunos = pd.read_excel(ARQUIVO_ALUNOS)

    df_alunos.columns = [c.strip() for c in df_alunos.columns]
    falta = [c for c in ["Nome", "Turma"] if c not in df_alunos.columns]
    if falta:
        raise ValueError(f"A planilha '{ARQUIVO_ALUNOS}' precisa das colunas: {', '.join(falta)}")
    if "Nivel" not in df_alunos.columns:
        df_alunos["Nivel"] = ""

    ordem = ["Lion Stars", "Junior", "Adultos"]
    df_alunos["Nivel"] = pd.Categorical(df_alunos["Nivel"], categories=ordem, ordered=True)
    df_alunos = df_alunos.sort_values(by=["Nivel", "Turma", "Nome"])

    lista_alunos = df_alunos.to_dict(orient="records")

    root = tk.Tk()
    root.withdraw()
    professor = simpledialog.askstring("Professor", "Digite o nome do professor:")
    root.deiconify()

    app = BoletimApp(root, lista_alunos, professor or "")
    root.mainloop()
