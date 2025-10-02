import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd

ARQUIVO_ALUNOS = "alunos.xlsx"

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

MAPEAMENTO_CHAVES = {
    "Comunicação oral": "Comunicacao",
    "Compreensão oral": "Compreensao",
    "Interesse pelo processo de aprendizagem": "Interesse",
    "Colaboração com colegas": "Colaboracao",
    "Engajamento nas atividades de sala": "Engajamento",
    "Comunicação escrita": "ComunicacaoE",
    "Compreensão escrita": "CompreensaoEscrita",
    "Produção oral": "ProducaoOral",
    "Produção escrita": "ProducaoE",
    "Progress Check": "ProgressCheck"
}

OPCOES = ["A", "B", "C", "D"]

LEGENDA = """(A) = Atingiu plenamente (100–90%)
(B) = Atingiu satisfatoriamente (89–75%)
(C) = Atingiu parcialmente (74–60%)
(D) (N/A) = Ainda não atingiu / Não há evidências (59% ou menos)"""

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
        self.resultados = []
        self.alunos_filtrados = []
        self.index = 0

        self.root.title("Sistema de Boletins")
        self.root.geometry("780x650")

        # seleção de nível
        frame_sel = tk.Frame(root)
        frame_sel.pack(pady=10)

        tk.Label(frame_sel, text="Nível:").pack(side="left", padx=5)
        self.combo_nivel = ttk.Combobox(frame_sel, values=list(CRITERIOS_POR_NIVEL.keys()), state="readonly", width=20)
        self.combo_nivel.pack(side="left", padx=5)

        self.btn_carregar = tk.Button(frame_sel, text="Carregar turma", command=self.carregar_turma)
        self.btn_carregar.pack(side="left", padx=10)

        # seleção de aluno específico
        tk.Label(frame_sel, text="Aluno específico:").pack(side="left", padx=5)
        self.combo_aluno_especifico = ttk.Combobox(frame_sel, values=[], state="readonly", width=30)
        self.combo_aluno_especifico.pack(side="left", padx=5)
        self.btn_carregar_aluno = tk.Button(frame_sel, text="Carregar aluno", command=self.carregar_aluno_especifico)
        self.btn_carregar_aluno.pack(side="left", padx=10)

        self.label_nome = tk.Label(root, text="", font=("Arial", 14, "bold"))
        self.label_nome.pack(pady=10)

        self.frame_criterios = tk.Frame(root)
        self.frame_criterios.pack(pady=10, fill="x")

        self.combos = {}

        # botões de navegação
        frame_botoes = tk.Frame(root)
        frame_botoes.pack(pady=15)
        self.btn_voltar = tk.Button(frame_botoes, text="Voltar", command=self.voltar)
        self.btn_voltar.pack(side="left", padx=10)
        self.btn_pular = tk.Button(frame_botoes, text="Pular", command=self.pular)
        self.btn_pular.pack(side="left", padx=10)
        self.btn_salvar = tk.Button(frame_botoes, text="Salvar & Próximo", command=self.salvar)
        self.btn_salvar.pack(side="left", padx=10)

        self.label_legenda = tk.Label(root, text=LEGENDA, justify="left", font=("Arial", 10), anchor="w")
        self.label_legenda.pack(pady=10, fill="x")

    def carregar_turma(self):
        nivel = self.combo_nivel.get()
        if not nivel:
            messagebox.showwarning("Atenção", "Selecione um nível.")
            return

        self.alunos_filtrados = [a for a in self.alunos if a.get("Nivel", "") == nivel]
        if not self.alunos_filtrados:
            messagebox.showinfo("Info", f"Não há alunos para {nivel}.")
            return

        # preencher combobox de aluno específico
        self.combo_aluno_especifico["values"] = [a["Nome"] for a in self.alunos_filtrados]

        self.index = 0
        self.mostrar_aluno()

    def carregar_aluno_especifico(self):
        nome = self.combo_aluno_especifico.get()
        nivel = self.combo_nivel.get()
        if not nome or not nivel:
            messagebox.showwarning("Atenção", "Selecione um nível e um aluno.")
            return

        aluno = next((a for a in self.alunos if a["Nome"] == nome and a["Nivel"] == nivel), None)
        if not aluno:
            messagebox.showerror("Erro", "Aluno não encontrado.")
            return

        self.aluno_atual = aluno
        criterios = CRITERIOS_POR_NIVEL.get(nivel, [])
        aluno_salvo = self.buscar_resultado_salvo(aluno)
        self.label_nome.config(text=f"Aluno: {aluno['Nome']}  |  Turma: {aluno.get('Turma','')}  |  Nível: {nivel}")
        self.criar_campos(criterios, aluno_existente=aluno_salvo)

    def criar_campos(self, criterios, aluno_existente=None):
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

            # preencher se já existir nota lançada
            if aluno_existente:
                chave = MAPEAMENTO_CHAVES.get(criterio, criterio)
                if chave in aluno_existente:
                    cb.set(aluno_existente[chave])

            self.combos[criterio] = cb

    def mostrar_aluno(self):
        if self.index < 0: self.index = 0
        if self.index >= len(self.alunos_filtrados):
            messagebox.showinfo("Fim", "Todos os alunos desta turma foram avaliados!")
            return

        aluno = self.alunos_filtrados[self.index]
        self.aluno_atual = aluno
        nome, turma, nivel = aluno.get("Nome", ""), aluno.get("Turma", ""), aluno.get("Nivel", "")
        aluno_salvo = self.buscar_resultado_salvo(aluno)
        self.label_nome.config(text=f"Aluno: {nome}  |  Turma: {turma}  |  Nível: {nivel}")
        criterios = CRITERIOS_POR_NIVEL.get(nivel, [])
        self.criar_campos(criterios, aluno_existente=aluno_salvo)

    def buscar_resultado_salvo(self, aluno):
        return next(
            (r for r in self.resultados if r["Aluno"] == aluno["Nome"] and r["Turma"] == aluno["Turma"]),
            None
        )

    def salvar(self):
        if not hasattr(self, "aluno_atual"):
            return
        aluno_row = self.aluno_atual
        nivel = aluno_row.get("Nivel", "")
        notas = {
            "Aluno": aluno_row.get("Nome", ""),
            "Turma": aluno_row.get("Turma", ""),
            "Nivel": nivel,
            "Professor": self.professor
        }
        valores_min, valores_max = [], []
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
        if nivel == "Adultos" and valores_min:
            media_min = sum(valores_min)/len(valores_min)
            media_max = sum(valores_max)/len(valores_max)
            notas["NotaSugerida"] = f"{int(media_min)} - {int(media_max)}"
            nota_final = simpledialog.askinteger(
                "Nota Final",
                f"Nota sugerida para {aluno_row.get('Nome', '')}: {int(media_min)}–{int(media_max)}\nDigite a nota final:",
                minvalue=0, maxvalue=100
            )
            notas["Nota"] = nota_final if nota_final is not None else int((media_min+media_max)/2)

        # substituir se já existe
        self.resultados = [r for r in self.resultados if not (r["Aluno"] == aluno_row["Nome"] and r["Turma"] == aluno_row["Turma"])]
        self.resultados.append(notas)

        df = pd.DataFrame(self.resultados)
        df.to_excel("resultados_boletim.xlsx", index=False)

        self.index += 1
        self.mostrar_aluno()

    def voltar(self):
        self.index -= 1
        self.mostrar_aluno()

    def pular(self):
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
