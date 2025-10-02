import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd

# ====== CONFIGURAÇÕES ======
ARQUIVO_ALUNOS = "alunos.xlsx"

CRITERIOS = [
    "Comunicação oral",
    "Compreensão oral",
    "Interesse pelo processo de aprendizagem",
    "Colaboração com colegas",
    "Engajamento nas atividades de sala"
]

OPCOES = ["A", "B", "C", "D"]

LEGENDA = """(A) = Atingiu plenamente (100–90%)
(B) = Atingiu satisfatoriamente (89–75%)
(C) = Atingiu parcialmente (74–60%)
(D) (N/A) = Ainda não atingiu / Não há evidências (59% ou menos)"""

class BoletimApp:
    def __init__(self, root, alunos, professor):
        self.root = root
        self.alunos = alunos            # lista de dicts: Nome, Turma, (Nivel opcional)
        self.professor = professor
        self.index = 0
        self.resultados = []

        self.root.title("Sistema de Boletins")
        self.root.geometry("620x520")

        self.label_nome = tk.Label(root, text="", font=("Arial", 14, "bold"))
        self.label_nome.pack(pady=10)

        self.combos = {}
        for criterio in CRITERIOS:
            frame = tk.Frame(root)
            frame.pack(pady=5, fill="x")

            lbl = tk.Label(frame, text=criterio, width=40, anchor="w")
            lbl.pack(side="left")

            cb = ttk.Combobox(frame, values=OPCOES, state="readonly", width=5)
            cb.pack(side="left", padx=10)
            self.combos[criterio] = cb

        self.btn = tk.Button(root, text="Salvar e próximo", command=self.salvar)
        self.btn.pack(pady=20)

        self.label_legenda = tk.Label(root, text=LEGENDA, justify="left", font=("Arial", 10), anchor="w")
        self.label_legenda.pack(pady=10, fill="x")

        self.mostrar_aluno()

    def mostrar_aluno(self):
        if self.index < len(self.alunos):
            aluno = self.alunos[self.index]
            nome = aluno.get("Nome", "")
            turma = aluno.get("Turma", "")
            nivel = aluno.get("Nivel", "")
            # Mostra também o nível para conferência
            self.label_nome.config(text=f"Aluno: {nome}  |  Turma: {turma}  |  Nível: {nivel}")
            for cb in self.combos.values():
                cb.set("")
        else:
            messagebox.showinfo("Fim", "Todos os alunos foram avaliados!")
            df = pd.DataFrame(self.resultados)
            df.to_excel("resultados_boletim.xlsx", index=False)
            self.root.quit()

    def salvar(self):
        aluno_row = self.alunos[self.index]
        notas = {
            "Aluno": aluno_row.get("Nome", ""),
            "Turma": aluno_row.get("Turma", ""),
            "Nivel": aluno_row.get("Nivel", ""),      # <- agora vem da planilha
            "Professor": self.professor,
            "Comunicacao": self.combos["Comunicação oral"].get(),
            "Compreensao": self.combos["Compreensão oral"].get(),
            "Interesse": self.combos["Interesse pelo processo de aprendizagem"].get(),
            "Colaboracao": self.combos["Colaboração com colegas"].get(),
            "Engajamento": self.combos["Engajamento nas atividades de sala"].get()
        }
        # validação rápida: todas as notas escolhidas
        for crit, val in list(notas.items())[-5:]:
            if val == "":
                messagebox.showwarning("Atenção", f"Selecione uma nota para '{crit}'")
                return

        self.resultados.append(notas)
        self.index += 1
        self.mostrar_aluno()

if __name__ == "__main__":
    # Lê a planilha de alunos (espera colunas: Nome, Turma, Nivel)
    df_alunos = pd.read_excel(ARQUIVO_ALUNOS)

    # Normaliza nomes de colunas comuns (tira espaços e capitaliza)
    df_alunos.columns = [c.strip() for c in df_alunos.columns]
    falta = [c for c in ["Nome", "Turma"] if c not in df_alunos.columns]
    if falta:
        raise ValueError(f"A planilha '{ARQUIVO_ALUNOS}' precisa das colunas: {', '.join(falta)}")
    if "Nivel" not in df_alunos.columns:
        df_alunos["Nivel"] = ""  # opcional: pode deixar em branco se não tiver

    lista_alunos = df_alunos.to_dict(orient="records")

    root = tk.Tk()
    root.withdraw()
    professor = simpledialog.askstring("Professor", "Digite o nome do professor:")
    root.deiconify()

    app = BoletimApp(root, lista_alunos, professor or "")
    root.mainloop()
