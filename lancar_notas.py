import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import os, sys

# ================== SUPORTE A CAMINHOS (.py e .exe) ==================
def app_dir() -> str:
    if getattr(sys, "_MEIPASS", None):            # PyInstaller (onefile)
        return sys._MEIPASS
    if getattr(sys, "frozen", False):             # PyInstaller (onedir)
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def resource_path(*paths):
    base = app_dir()
    return os.path.join(base, *paths)

# importa a função de gerar PDFs
from gerar_pdfs import gerar_boletins

# ====== ARQUIVOS ======
BASE = app_dir()
ARQUIVO_ALUNOS     = os.path.join(BASE, "alunos.xlsx")
ARQUIVO_RESULTADOS = os.path.join(BASE, "resultados_boletim.xlsx")

# ====== LISTAS DE SUBNÍVEIS ======
ANTIGO_SUBNIVEIS = [
    "High Resolution 4", "High Resolution 5", "High Resolution 6",
    "Basic 5", "Basic 6",
    "New Plus Adult 3",
]

ADULTOS_SUBNIVEIS = [
    "Express Pack 1", "Express Pack 2", "Express Pack 3",
    "Inter Teens 1", "Inter Teens 2", "Inter Teens 3",
    "Teen League 1", "Teen League 2", "Teen League 3", "Teen League 4",
]

# ====== CRITÉRIOS POR NÍVEL ======
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
    ],
    # ANTIGO: 4 conceitos A–D + 3 notas numéricas (WB1, WB2, CP)
    "Antigo": [
        "Compreensão oral",
        "Compreensão escrita",
        "Produção oral",
        "Produção escrita",
    ],
}

MAPEAMENTO_CHAVES = {
    "Comunicação oral": "Comunicacao",
    "Compreensão oral": "Compreensao",          # Lion/Junior
    "Interesse pelo processo de aprendizagem": "Interesse",
    "Colaboração com colegas": "Colaboracao",
    "Engajamento nas atividades de sala": "Engajamento",
    "Comunicação escrita": "ComunicacaoE",
    "Compreensão escrita": "CompreensaoEscrita",
    "Produção oral": "ProducaoOral",            # Adultos
    "Produção escrita": "ProducaoE",
    "Progress Check": "ProgressCheck",

    # Antigo → placeholders do modelo
    "Compreensão oral__ANTIGO": "CompreensaoO",
    "Compreensão escrita__ANTIGO": "CompreensaoE",
    "Produção oral__ANTIGO": "ProducaoO",
    "Produção escrita__ANTIGO": "ProducaoE",
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
    def __init__(self, root, alunos, professor, resultados_existentes=None):
        self.root = root
        self.alunos = alunos
        self.professor = professor
        self.resultados = resultados_existentes or []
        self.alunos_filtrados = []
        self.index = 0

        self.root.title("Sistema de Boletins")
        self.root.geometry("980x860")

        # ===== Barra de seleção =====
        frame_sel = tk.Frame(root)
        frame_sel.pack(pady=10)

        tk.Label(frame_sel, text="Nível:").pack(side="left", padx=5)
        self.combo_nivel = ttk.Combobox(frame_sel, values=list(CRITERIOS_POR_NIVEL.keys()), state="readonly", width=20)
        self.combo_nivel.pack(side="left", padx=5)

        self.btn_carregar = tk.Button(frame_sel, text="Carregar turma", command=self.carregar_turma)
        self.btn_carregar.pack(side="left", padx=10)

        # --- filtro de turma ---
        tk.Label(frame_sel, text="Turma:").pack(side="left", padx=(10,5))
        self.combo_turma = ttk.Combobox(frame_sel, values=[], state="readonly", width=18)
        self.combo_turma.pack(side="left", padx=5)

        # seleção de aluno específico
        tk.Label(frame_sel, text="Aluno:").pack(side="left", padx=(10,5))
        self.combo_aluno_especifico = ttk.Combobox(frame_sel, values=[], state="readonly", width=30)
        self.combo_aluno_especifico.pack(side="left", padx=5)
        self.btn_carregar_aluno = tk.Button(frame_sel, text="Carregar aluno", command=self.carregar_aluno_especifico)
        self.btn_carregar_aluno.pack(side="left", padx=10)

        # botão gerar PDFs
        self.btn_gerar_pdfs = tk.Button(root, text="📄 Gerar PDFs", bg="lightblue", command=self.gerar_pdfs)
        self.btn_gerar_pdfs.pack(pady=10)

        self.label_nome = tk.Label(root, text="", font=("Arial", 14, "bold"))
        self.label_nome.pack(pady=10)

        self.frame_criterios = tk.Frame(root)
        self.frame_criterios.pack(pady=10, fill="x")

        # Campos extras / estado
        self.entry_wb1 = None
        self.entry_wb2 = None
        self.entry_cp  = None
        self.label_media = None

        # ANTIGO: rótulo por turma + override por aluno
        self.rotulo_antigo_por_turma = {}                 # {"SEG-QUA1910": "High Resolution 5", ...}
        self.combo_rotulo_turma = None
        self.combo_rotulo_antigo = None
        self.var_override_rotulo = tk.BooleanVar(value=False)

        # ADULTOS: subnível único
        self.combo_subadultos = None

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

    # ================== AÇÕES DE TELA ==================
    def gerar_pdfs(self):
        try:
            gerar_boletins()
            messagebox.showinfo("Sucesso", "Boletins gerados em 'boletins_pdf'")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar PDFs: {e}")

    def carregar_turma(self):
        nivel = self.combo_nivel.get()
        if not nivel:
            messagebox.showwarning("Atenção", "Selecione um nível.")
            return

        # lista de turmas para o nível
        turmas_nivel = sorted({ a.get("Turma","") for a in self.alunos if a.get("Nivel","") == nivel })
        self.combo_turma["values"] = turmas_nivel

        turma_escolhida = self.combo_turma.get().strip()
        if len(turmas_nivel) > 1 and not turma_escolhida:
            messagebox.showinfo("Seleção de Turma",
                "Há várias turmas para este nível. Selecione a turma no campo 'Turma:' e clique em 'Carregar turma'.")
            return
        if not turma_escolhida and len(turmas_nivel) == 1:
            turma_escolhida = turmas_nivel[0]
            self.combo_turma.set(turma_escolhida)

        # filtra por Nível + Turma
        self.alunos_filtrados = [
            a for a in self.alunos
            if a.get("Nivel","") == nivel and a.get("Turma","") == turma_escolhida
        ]
        if not self.alunos_filtrados:
            messagebox.showinfo("Info", f"Não há alunos para {nivel} / {turma_escolhida}.")
            return

        # ANTIGO: rótulo da turma
        if nivel == "Antigo":
            if self.combo_rotulo_turma is None:
                barra = tk.Frame(self.root)
                barra.pack(pady=(0, 6))
                tk.Label(barra, text="Subnível da TURMA (Antigo):").pack(side="left")
                self.combo_rotulo_turma = ttk.Combobox(barra, values=ANTIGO_SUBNIVEIS, width=28)
                self.combo_rotulo_turma.pack(side="left", padx=8)
                tk.Button(barra, text="Aplicar a todos desta turma", command=self._aplicar_rotulo_turma)\
                    .pack(side="left", padx=6)

            rotulo_salvo = self.rotulo_antigo_por_turma.get(turma_escolhida, "")
            self.combo_rotulo_turma.set(rotulo_salvo or ANTIGO_SUBNIVEIS[0])

        # nomes no combo de aluno
        self.combo_aluno_especifico["values"] = [a["Nome"] for a in self.alunos_filtrados]
        self.index = 0
        self.mostrar_aluno()

    def _aplicar_rotulo_turma(self):
        """Salva o subnível escolhido para a turma atualmente selecionada (Antigo)."""
        turma = self.combo_turma.get().strip()
        if not turma:
            messagebox.showwarning("Atenção", "Selecione a Turma primeiro.")
            return
        rotulo = (self.combo_rotulo_turma.get() or "").strip() if self.combo_rotulo_turma else ""
        if not rotulo:
            messagebox.showwarning("Atenção", "Informe um subnível (ex.: High Resolution 5).")
            return
        if rotulo not in ANTIGO_SUBNIVEIS:
            messagebox.showwarning("Atenção", "Escolha um subnível válido dos antigos.")
            return
        self.rotulo_antigo_por_turma[turma] = rotulo
        messagebox.showinfo("OK", f"Subnível definido para a turma {turma}: {rotulo}")

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

    def _preencher_num(self, widget, valor):
        try:
            if valor is None or pd.isna(valor) or str(valor).strip() == "":
                return
            widget.delete(0, tk.END)
            widget.insert(0, str(int(float(valor))))
        except:
            widget.delete(0, tk.END)
            widget.insert(0, str(valor))

    def criar_campos(self, criterios, aluno_existente=None):
        for widget in self.frame_criterios.winfo_children():
            widget.destroy()
        self.combos.clear()

        # reset extras
        self.entry_wb1 = None
        self.entry_wb2 = None
        self.entry_cp  = None
        self.label_media = None
        self.combo_rotulo_antigo = None
        self.var_override_rotulo.set(False)
        self.combo_subadultos = None

        nivel_atual = getattr(self, "aluno_atual", {}).get("Nivel", "") if hasattr(self, "aluno_atual") else ""

        # Adultos → subnível único (lista fixa)
        if nivel_atual == "Adultos":
            frame_sub = tk.LabelFrame(self.frame_criterios, text="Subnível (Adultos)", padx=10, pady=10)
            frame_sub.pack(pady=10, fill="x")

            tk.Label(frame_sub, text="Selecione o subnível:").grid(row=0, column=0, sticky="w", padx=5)
            self.combo_subadultos = ttk.Combobox(frame_sub, values=ADULTOS_SUBNIVEIS, state="readonly", width=25)
            self.combo_subadultos.grid(row=0, column=1, padx=10)

            if aluno_existente:
                salvo = str(aluno_existente.get("Nivel", "")).strip()
                if salvo in ADULTOS_SUBNIVEIS:
                    self.combo_subadultos.set(salvo)

            tk.Label(self.frame_criterios, text="").pack(pady=10)
        else:
            self.combo_subadultos = None

        # Campos de conceitos A–D
        for criterio in criterios:
            frame = tk.Frame(self.frame_criterios)
            frame.pack(pady=8, fill="x")
            lbl = tk.Label(frame, text=criterio, width=40, anchor="w", font=("Arial", 11))
            lbl.pack(side="left")
            cb = ttk.Combobox(frame, values=OPCOES, state="readonly", width=5)
            cb.pack(side="left", padx=10)

            if aluno_existente:
                if nivel_atual == "Antigo":
                    chave = MAPEAMENTO_CHAVES.get(f"{criterio}__ANTIGO")
                else:
                    chave = MAPEAMENTO_CHAVES.get(criterio, criterio)
                if chave in aluno_existente:
                    cb.set(aluno_existente[chave])

            self.combos[(nivel_atual, criterio)] = cb

        # ANTIGO: Notas numéricas e rótulo por aluno (override opcional)
        if nivel_atual == "Antigo":
            box = tk.LabelFrame(self.frame_criterios, text="Notas Numéricas (0–100)", padx=10, pady=10)
            box.pack(pady=12, fill="x")

            tk.Label(box, text="WritingBit1").grid(row=0, column=0, sticky="w")
            self.entry_wb1 = tk.Entry(box, width=7, justify="center")
            self.entry_wb1.grid(row=0, column=1, padx=8)

            tk.Label(box, text="WritingBit2").grid(row=0, column=2, sticky="w")
            self.entry_wb2 = tk.Entry(box, width=7, justify="center")
            self.entry_wb2.grid(row=0, column=3, padx=8)

            tk.Label(box, text="CheckPoint").grid(row=0, column=4, sticky="w")
            self.entry_cp = tk.Entry(box, width=7, justify="center")
            self.entry_cp.grid(row=0, column=5, padx=8)

            self.label_media = tk.Label(box, text="Média (Nota): —", font=("Arial", 10, "italic"))
            self.label_media.grid(row=1, column=0, columnspan=6, sticky="w", pady=(10,0))

            if aluno_existente:
                self._preencher_num(self.entry_wb1, aluno_existente.get("WritingBit1"))
                self._preencher_num(self.entry_wb2, aluno_existente.get("WritingBit2"))
                self._preencher_num(self.entry_cp,  aluno_existente.get("CheckPoint"))
                if aluno_existente.get("Nota") not in (None, "", float("nan")):
                    self.label_media.config(text=f"Média (Nota): {aluno_existente.get('Nota')}")

            # ---- Override por aluno (opcional) ----
            box_rot = tk.LabelFrame(self.frame_criterios, text="Subnível (override por aluno - opcional)", padx=10, pady=8)
            box_rot.pack(pady=6, fill="x")

            linha = tk.Frame(box_rot)
            linha.pack(fill="x")
            tk.Checkbutton(linha, text="Usar subnível personalizado apenas para ESTE aluno",
                           variable=self.var_override_rotulo).pack(side="left")

            self.combo_rotulo_antigo = ttk.Combobox(box_rot, values=ANTIGO_SUBNIVEIS, width=28)
            self.combo_rotulo_antigo.pack(anchor="w", pady=(8,0))

            # Prefill: (1) salvo por aluno, (2) rótulo da turma
            rotulo_existente = ""
            if aluno_existente:
                rotulo_existente = str(aluno_existente.get("Nivel", "")).strip()
            turma_atual = getattr(self, "aluno_atual", {}).get("Turma", "")
            rotulo_turma  = self.rotulo_antigo_por_turma.get(turma_atual, "")
            if rotulo_existente:
                self.combo_rotulo_antigo.set(rotulo_existente)
            elif rotulo_turma:
                self.combo_rotulo_antigo.set(rotulo_turma)

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

    def _pega_num(self, entry, nome_campo):
        val = entry.get().strip()
        if val == "":
            raise ValueError(f"Preencha {nome_campo} (0–100).")
        try:
            n = int(float(val))
        except:
            raise ValueError(f"{nome_campo} precisa ser número.")
        if not (0 <= n <= 100):
            raise ValueError(f"{nome_campo} deve estar entre 0 e 100.")
        return n

    def salvar(self):
        if not hasattr(self, "aluno_atual"):
            return
        aluno_row = self.aluno_atual
        nivel = aluno_row.get("Nivel", "")
        notas = {
            "Aluno": aluno_row.get("Nome", ""),
            "Turma": aluno_row.get("Turma", ""),
            "Nivel": nivel,           # pode ser sobrescrito abaixo
            "Professor": self.professor
        }
        valores_min, valores_max = [], []

        # Preenche conceitos A–D
        for (nivel_ctx, crit), cb in self.combos.items():
            if nivel_ctx != nivel:
                continue
            val = cb.get()
            if val == "":
                messagebox.showwarning("Atenção", f"Selecione uma nota para '{crit}'")
                return

            if nivel == "Antigo":
                chave = MAPEAMENTO_CHAVES.get(f"{crit}__ANTIGO")
            else:
                chave = MAPEAMENTO_CHAVES.get(crit, crit)
            notas[chave] = val

            if nivel == "Adultos":
                min_val, max_val = NOTAS_RANGE[val]
                valores_min.append(min_val)
                valores_max.append(max_val)

        # Regras específicas por nível
        if nivel == "Adultos":
            subnivel = self.combo_subadultos.get().strip() if self.combo_subadultos else ""
            if not subnivel:
                messagebox.showwarning("Atenção", "Selecione o subnível de Adultos (Ex.: Express Pack 1, Inter Teens 2...).")
                return

            notas["Nivel"] = subnivel  # rótulo exato salvo

            if valores_min:
                media_min = sum(valores_min)/len(valores_min)
                media_max = sum(valores_max)/len(valores_max)
                notas["NotaSugerida"] = f"{int(media_min)} - {int(media_max)}"
                nota_final = simpledialog.askinteger(
                    "Nota Final",
                    f"Nota sugerida para {aluno_row.get('Nome', '')}: {int(media_min)}–{int(media_max)}\nDigite a nota final:",
                    minvalue=0, maxvalue=100
                )
                notas["Nota"] = nota_final if nota_final is not None else int((media_min+media_max)/2)

        elif nivel == "Antigo":
            # notas numéricas
            try:
                wb1 = self._pega_num(self.entry_wb1, "WritingBit1")
                wb2 = self._pega_num(self.entry_wb2, "WritingBit2")
                cp  = self._pega_num(self.entry_cp,  "CheckPoint")
            except ValueError as e:
                messagebox.showwarning("Atenção", str(e))
                return
            media = round((wb1 + wb2 + cp) / 3)
            notas["WritingBit1"] = wb1
            notas["WritingBit2"] = wb2
            notas["CheckPoint"]  = cp
            notas["Nota"]        = media

            # rótulo do subnível (turma → override por aluno)
            turma_atual = aluno_row.get("Turma", "")
            rotulo_turma = self.rotulo_antigo_por_turma.get(turma_atual, "").strip()
            rotulo_aluno = (self.combo_rotulo_antigo.get() or "").strip() if self.combo_rotulo_antigo else ""

            if self.var_override_rotulo.get():
                rotulo = rotulo_aluno
            else:
                rotulo = rotulo_turma or rotulo_aluno

            if not rotulo or rotulo not in ANTIGO_SUBNIVEIS:
                messagebox.showwarning("Atenção",
                    "Defina um subnível válido para a turma (botão 'Aplicar a todos desta turma') "
                    "ou marque override e escolha para este aluno.")
                return

            # cache para próxima navegação
            if rotulo and turma_atual:
                self.rotulo_antigo_por_turma[turma_atual] = rotulo

            notas["Nivel"]  = rotulo     # ex.: "High Resolution 5"
            notas["Modelo"] = "Antigo"   # indica qual template usar

        # Salva/atualiza linha deste aluno+turma
        self.resultados = [r for r in self.resultados if not (r["Aluno"] == aluno_row["Nome"] and r["Turma"] == aluno_row["Turma"])]
        self.resultados.append(notas)

        df = pd.DataFrame(self.resultados)
        df.to_excel(ARQUIVO_RESULTADOS, index=False)

        self.index += 1
        self.mostrar_aluno()

    def voltar(self):
        self.index -= 1
        self.mostrar_aluno()

    def pular(self):
        self.index += 1
        self.mostrar_aluno()

# ================== MAIN ==================
if __name__ == "__main__":
    if not os.path.exists(ARQUIVO_ALUNOS):
        messagebox.showerror("Arquivo não encontrado",
            f"Não encontrei {ARQUIVO_ALUNOS}.\n"
            "Deixe o alunos.xlsx na mesma pasta do aplicativo.")
        raise SystemExit(1)

    df_alunos = pd.read_excel(ARQUIVO_ALUNOS)
    df_alunos.columns = [c.strip() for c in df_alunos.columns]
    falta = [c for c in ["Nome", "Turma"] if c not in df_alunos.columns]
    if falta:
        raise ValueError(f"A planilha '{ARQUIVO_ALUNOS}' precisa das colunas: {', '.join(falta)}")
    if "Nivel" not in df_alunos.columns:
        df_alunos["Nivel"] = ""

    # Ordenação padrão
    ordem = ["Lion Stars", "Junior", "Adultos", "Antigo"]
    df_alunos["Nivel"] = pd.Categorical(df_alunos["Nivel"], categories=ordem, ordered=True)
    df_alunos = df_alunos.sort_values(by=["Nivel", "Turma", "Nome"])
    lista_alunos = df_alunos.to_dict(orient="records")

    resultados_existentes = []
    if os.path.exists(ARQUIVO_RESULTADOS):
        try:
            resultados_existentes = pd.read_excel(ARQUIVO_RESULTADOS).to_dict(orient="records")
            print(f"✅ {len(resultados_existentes)} registros carregados de {ARQUIVO_RESULTADOS}")
        except Exception as e:
            print(f"⚠️ Não foi possível carregar resultados anteriores: {e}")

    root = tk.Tk()
    root.withdraw()
    professor = simpledialog.askstring("Professor", "Digite o nome do professor:")
    root.deiconify()
    app = BoletimApp(root, lista_alunos, professor or "", resultados_existentes)
    root.mainloop()
