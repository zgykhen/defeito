# paragem_defeito.py
# Paragem ao DEFEITO / Tally sheet
# Registers quality control readings per banco (bank) with two control points:
# Controlo Funcional and Controlo Aspeto. Each can be OK or NOK;
# if NOK, operator selects multiple defect reasons (checkbox-style multi-select).

import os
import csv
import sqlite3
import logging
import configparser
from uuid import uuid4
from datetime import datetime, date
from typing import List, Tuple, Optional

import tkinter as tk
from tkinter import ttk, messagebox

from app_paths import APP_DIR, CONFIG_FILENAME
from db_utils import db_path, db_connect

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
LOG_PREFIX = "log_paragem"
DB_TABLE = "paragens"

CORES = {
    "azul":         "#0024D3",
    "azul_claro":   "#00A9EB",
    "cinza_claro":  "#8C8C8C",
    "cinza_escuro": "#575757",
    "branco":       "#FFFFFF",
    "fundo":        "#F5F5F5",
    "painel_titulo":"#E8F0FE",
    "verde":        "#2E7D32",
    "vermelho":     "#C62828",
    "laranja":      "#E65100",
}

# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------

def _ler_config() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    cfg_path = os.path.join(APP_DIR, CONFIG_FILENAME)
    cfg.read(cfg_path, encoding="utf-8")
    return cfg


def carregar_sessao_dropdowns() -> Tuple[List[str], List[str]]:
    cfg = _ler_config()
    projetos_raw = cfg.get("dropdowns", "projetos_linhas", fallback="Picking")
    turnos_raw   = cfg.get("dropdowns", "turnos", fallback="A,B,C")
    projetos = [p.strip() for p in projetos_raw.split(",") if p.strip()]
    turnos   = [t.strip() for t in turnos_raw.split(",") if t.strip()]
    if not projetos:
        projetos = ["Picking"]
    if not turnos:
        turnos = ["A", "B", "C"]
    return projetos, turnos


def carregar_postos() -> List[str]:
    cfg = _ler_config()
    raw = cfg.get("paragem", "postos", fallback="Controlo Final,Duplo Controlo")
    postos = [p.strip() for p in raw.split(",") if p.strip()]
    return postos if postos else ["Controlo Final", "Duplo Controlo"]


def carregar_defeitos_funcionais() -> List[str]:
    cfg = _ler_config()
    fallback = ",".join(f"Teste Funcional {i}" for i in range(1, 11))
    raw = cfg.get("paragem", "defeitos_funcionais", fallback=fallback)
    items = [d.strip() for d in raw.split(",") if d.strip()]
    return items if items else [f"Teste Funcional {i}" for i in range(1, 11)]


def carregar_defeitos_aspeto() -> List[str]:
    cfg = _ler_config()
    fallback = ",".join(f"Teste Aspeto {i}" for i in range(1, 11))
    raw = cfg.get("paragem", "defeitos_aspeto", fallback=fallback)
    items = [d.strip() for d in raw.split(",") if d.strip()]
    return items if items else [f"Teste Aspeto {i}" for i in range(1, 11)]


def carregar_caminhos() -> Tuple[str, str]:
    cfg = _ler_config()
    log_dir = cfg.get("paths", "log", fallback=".").strip()
    db_dir  = cfg.get("paths", "db",  fallback=".").strip()
    if log_dir in (".", ""):
        log_dir = APP_DIR
    if db_dir in (".", ""):
        db_dir = APP_DIR
    return log_dir, db_dir


def carregar_caminho_logo() -> str:
    cfg = _ler_config()
    logo = cfg.get("paths", "logo", fallback="logo.png").strip()
    if not os.path.isabs(logo):
        logo = os.path.join(APP_DIR, logo)
    return logo


# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------

def _db_init_paragens(con: sqlite3.Connection) -> None:
    cur = con.cursor()
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS {DB_TABLE} (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            ts                  TEXT NOT NULL,
            operador            TEXT NOT NULL,
            projeto             TEXT NOT NULL,
            turno               TEXT,
            id_banco            TEXT NOT NULL,
            posto               TEXT,
            controlo_funcional  TEXT,
            defeitos_funcionais TEXT,
            cf_outro            TEXT,
            controlo_aspeto     TEXT,
            defeitos_aspeto     TEXT,
            ca_outro            TEXT,
            sessao_id           TEXT NOT NULL
        )
    """)
    # Add any missing columns via ALTER TABLE (idempotent)
    existing_cols = {row[1] for row in cur.execute(f"PRAGMA table_info({DB_TABLE})")}
    desired_cols = {
        "ts":                  "TEXT NOT NULL DEFAULT ''",
        "operador":            "TEXT NOT NULL DEFAULT ''",
        "projeto":             "TEXT NOT NULL DEFAULT ''",
        "turno":               "TEXT",
        "id_banco":            "TEXT NOT NULL DEFAULT ''",
        "posto":               "TEXT",
        "controlo_funcional":  "TEXT",
        "defeitos_funcionais": "TEXT",
        "cf_outro":            "TEXT",
        "controlo_aspeto":     "TEXT",
        "defeitos_aspeto":     "TEXT",
        "ca_outro":            "TEXT",
        "sessao_id":           "TEXT NOT NULL DEFAULT ''",
    }
    for col, col_def in desired_cols.items():
        if col not in existing_cols:
            try:
                cur.execute(f"ALTER TABLE {DB_TABLE} ADD COLUMN {col} {col_def}")
            except sqlite3.OperationalError:
                pass
    # Indexes
    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_paragens_ts ON {DB_TABLE}(ts)")
    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_paragens_sessao_id ON {DB_TABLE}(sessao_id)")
    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_paragens_id_banco ON {DB_TABLE}(id_banco)")
    con.commit()


# ---------------------------------------------------------------------------
# DefeitosBtnGroup widget
# ---------------------------------------------------------------------------

class DefeitosBtnGroup(tk.Frame):
    """Multi-select toggle button group for defect selection."""

    COLS = 5

    def __init__(self, master, defeitos: List[str], cores: dict, **kwargs):
        super().__init__(master, bg=cores["branco"], **kwargs)
        self._cores = cores
        self._defeitos = defeitos
        self._selected = [False] * len(defeitos)
        self._btns: List[tk.Button] = []
        self._outro_ativo = False
        self._entry_outro: Optional[tk.Entry] = None
        self._btn_outro: Optional[tk.Button] = None
        self._build()

    def _build(self):
        branco      = self._cores["branco"]
        cinza_escuro= self._cores["cinza_escuro"]
        laranja     = self._cores["laranja"]

        frame_grid = tk.Frame(self, bg=branco)
        frame_grid.pack(fill=tk.X, padx=4, pady=(4, 2))

        for idx, nome in enumerate(self._defeitos):
            row = idx // self.COLS
            col = idx % self.COLS
            btn = tk.Button(
                frame_grid,
                text=nome,
                width=16,
                font=("Segoe UI", 9),
                wraplength=120,
                relief=tk.RAISED,
                bg=branco,
                fg=cinza_escuro,
                cursor="hand2",
                command=lambda i=idx: self._toggle(i),
            )
            btn.grid(row=row, column=col, padx=3, pady=3, sticky="ew")
            self._btns.append(btn)

        for c in range(self.COLS):
            frame_grid.columnconfigure(c, weight=1)

        # Outro row
        frame_outro = tk.Frame(self, bg=branco)
        frame_outro.pack(fill=tk.X, padx=4, pady=(0, 4))

        self._btn_outro = tk.Button(
            frame_outro,
            text="Outro...",
            width=10,
            font=("Segoe UI", 9),
            bg=branco,
            fg=cinza_escuro,
            cursor="hand2",
            relief=tk.RAISED,
            command=self._toggle_outro,
        )
        self._btn_outro.pack(side=tk.LEFT, padx=(0, 6))

        self._entry_outro = tk.Entry(
            frame_outro,
            width=50,
            font=("Segoe UI", 9),
            state=tk.DISABLED,
            bg="#EEEEEE",
        )
        self._entry_outro.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _toggle(self, idx: int):
        self._selected[idx] = not self._selected[idx]
        self._refresh_btn(idx)

    def _refresh_btn(self, idx: int):
        branco       = self._cores["branco"]
        cinza_escuro = self._cores["cinza_escuro"]
        laranja      = self._cores["laranja"]
        btn = self._btns[idx]
        if self._selected[idx]:
            btn.configure(bg=laranja, fg=branco, relief=tk.SUNKEN)
        else:
            btn.configure(bg=branco, fg=cinza_escuro, relief=tk.RAISED)

    def _toggle_outro(self):
        self._outro_ativo = not self._outro_ativo
        branco  = self._cores["branco"]
        laranja = self._cores["laranja"]
        cinza_escuro = self._cores["cinza_escuro"]
        if self._outro_ativo:
            self._btn_outro.configure(bg=laranja, fg=branco, relief=tk.SUNKEN)
            self._entry_outro.configure(state=tk.NORMAL, bg=branco)
            self._entry_outro.focus_set()
        else:
            self._btn_outro.configure(bg=branco, fg=cinza_escuro, relief=tk.RAISED)
            self._entry_outro.delete(0, tk.END)
            self._entry_outro.configure(state=tk.DISABLED, bg="#EEEEEE")

    # Public API

    def get_defeitos_str(self) -> str:
        return ",".join(
            self._defeitos[i]
            for i, sel in enumerate(self._selected)
            if sel
        )

    def get_outro(self) -> str:
        if self._outro_ativo and self._entry_outro:
            return self._entry_outro.get().strip()
        return ""

    def tem_selecao(self) -> bool:
        if any(self._selected):
            return True
        if self._outro_ativo:
            txt = self._entry_outro.get().strip() if self._entry_outro else ""
            return bool(txt)
        return False

    def outro_ativo_sem_texto(self) -> bool:
        if not self._outro_ativo:
            return False
        txt = self._entry_outro.get().strip() if self._entry_outro else ""
        return not bool(txt)

    def reset(self):
        self._selected = [False] * len(self._defeitos)
        for i in range(len(self._btns)):
            self._refresh_btn(i)
        if self._outro_ativo:
            self._outro_ativo = False
            branco       = self._cores["branco"]
            cinza_escuro = self._cores["cinza_escuro"]
            self._btn_outro.configure(bg=branco, fg=cinza_escuro, relief=tk.RAISED)
            self._entry_outro.delete(0, tk.END)
            self._entry_outro.configure(state=tk.DISABLED, bg="#EEEEEE")


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class ParagemDefeitoApp:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Paragem ao DEFEITO / Tally sheet")
        self.root.state("zoomed")
        self.root.minsize(1100, 900)
        self.root.configure(bg=CORES["fundo"])
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Session state
        self.sessao_iniciada: bool = False
        self.operador: str = ""
        self.projeto: str = ""
        self.turno: str = ""
        self.inicio_sessao: Optional[datetime] = None
        self.sessao_id: str = ""
        self.logfile: str = ""

        self.ultimas_leituras: List[tuple] = []
        # Each tuple: (id, id_banco, posto, ts, cf, defeitos_func_str, cf_outro, ca, defeitos_asp_str, ca_outro)

        self.db_con: Optional[sqlite3.Connection] = None
        self.db_path_atual: str = ""
        self.log_dir: str = APP_DIR
        self.db_dir: str = APP_DIR

        # Load config
        self.defeitos_funcionais = carregar_defeitos_funcionais()
        self.defeitos_aspeto     = carregar_defeitos_aspeto()
        self.postos_list         = carregar_postos()

        # Control vars
        self.var_cf = tk.StringVar(value="")
        self.var_ca = tk.StringVar(value="")

        self._timer_duracao = None

        self._construir_interface()

    # ------------------------------------------------------------------
    # Close handler
    # ------------------------------------------------------------------

    def _on_close(self):
        if self.sessao_iniciada:
            if not messagebox.askyesno(
                "Fechar",
                "Existe uma sessão ativa. Deseja mesmo fechar a aplicação?\n"
                "A sessão será terminada.",
                parent=self.root,
            ):
                return
            self._terminar_sessao()
        self.root.destroy()

    # ------------------------------------------------------------------
    # Interface construction
    # ------------------------------------------------------------------

    def _construir_interface(self):
        # Header
        header = tk.Frame(self.root, bg=CORES["azul"], height=64)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        logo_path = carregar_caminho_logo()
        self._logo_img = None
        if os.path.isfile(logo_path):
            try:
                self._logo_img = tk.PhotoImage(file=logo_path)
                tk.Label(header, image=self._logo_img, bg=CORES["azul"]).pack(
                    side=tk.LEFT, padx=16, pady=12)
            except tk.TclError:
                self._logo_img = None

        if self._logo_img is None:
            tk.Label(header, text="FORVIA", font=("Segoe UI", 18, "bold"),
                     fg=CORES["branco"], bg=CORES["azul"]).pack(side=tk.LEFT, padx=16, pady=12)

        tk.Label(
            header,
            text="Paragem ao DEFEITO / Tally sheet",
            font=("Segoe UI", 14, "bold"),
            bg=CORES["azul"],
            fg=CORES["branco"],
        ).pack(side=tk.LEFT, padx=20)

        # Main area
        main = tk.Frame(self.root, bg=CORES["fundo"])
        main.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        left = tk.Frame(main, bg=CORES["fundo"])
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right = tk.Frame(main, bg=CORES["fundo"], width=300)
        right.pack(side=tk.RIGHT, fill=tk.Y)
        right.pack_propagate(False)

        self._construir_sessao(left)
        self._construir_leitura(left)
        self._construir_ultimas(left)
        self._construir_resumo(right)
        self._construir_footer()

    def _painel_titulo(self, parent, texto: str):
        f = tk.Frame(parent, bg=CORES["painel_titulo"], pady=4)
        f.pack(fill=tk.X, pady=(8, 0))
        tk.Label(
            f, text=texto, font=("Segoe UI", 10, "bold"),
            bg=CORES["painel_titulo"], fg=CORES["cinza_escuro"], padx=8
        ).pack(side=tk.LEFT)
        return f

    # ------------------------------------------------------------------
    # Session panel
    # ------------------------------------------------------------------

    def _construir_sessao(self, parent):
        self._painel_titulo(parent, "Sessão")

        outer = tk.Frame(parent, bg=CORES["branco"], relief=tk.GROOVE, bd=1)
        outer.pack(fill=tk.X, pady=(0, 4))

        frame_sessao = tk.Frame(outer, bg=CORES["branco"], padx=12, pady=8)
        frame_sessao.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        projetos, turnos = carregar_sessao_dropdowns()

        # Row 0 – Operador
        tk.Label(frame_sessao, text="Operador:", bg=CORES["branco"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        self.entry_operador = tk.Entry(frame_sessao, width=18, font=("Segoe UI", 9))
        self.entry_operador.grid(row=0, column=1, sticky="w", padx=(6, 0), pady=2)

        # Row 1 – Projeto/Linha
        tk.Label(frame_sessao, text="Projeto/Linha:", bg=CORES["branco"],
                 font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=2)
        self.combo_projeto = ttk.Combobox(
            frame_sessao, width=18, state="readonly",
            values=projetos, font=("Segoe UI", 9)
        )
        if projetos:
            self.combo_projeto.current(0)
        self.combo_projeto.grid(row=1, column=1, sticky="w", padx=(6, 0), pady=2)

        # Row 2 – Turno
        tk.Label(frame_sessao, text="Turno:", bg=CORES["branco"],
                 font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=2)
        self.combo_turno = ttk.Combobox(
            frame_sessao, width=18, state="readonly",
            values=turnos, font=("Segoe UI", 9)
        )
        if turnos:
            self.combo_turno.current(0)
        self.combo_turno.grid(row=2, column=1, sticky="w", padx=(6, 0), pady=2)

        # Row 3 – Status
        self.label_sessao = tk.Label(
            frame_sessao, text="Sem sessão ativa.",
            bg=CORES["branco"], fg=CORES["cinza_claro"], font=("Segoe UI", 9)
        )
        self.label_sessao.grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # Row 4 – Button
        self.btn_iniciar = tk.Button(
            frame_sessao, text="Iniciar sessão",
            font=("Segoe UI", 9, "bold"),
            bg=CORES["azul"], fg=CORES["branco"],
            relief=tk.FLAT, padx=12, pady=4,
            cursor="hand2",
            command=self._iniciar_sessao,
        )
        self.btn_iniciar.grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Right – Clock
        frame_hora = tk.Frame(outer, bg=CORES["azul"], width=160, padx=10, pady=8)
        frame_hora.pack(side=tk.RIGHT, fill=tk.Y)
        frame_hora.pack_propagate(False)

        self.label_hora = tk.Label(
            frame_hora, text="--:--:--",
            font=("Segoe UI", 22, "bold"),
            bg=CORES["azul"], fg=CORES["branco"]
        )
        self.label_hora.pack()

        self.label_inicio_sessao = tk.Label(
            frame_hora, text="",
            font=("Segoe UI", 9),
            bg=CORES["azul"], fg=CORES["branco"]
        )
        self.label_inicio_sessao.pack()

        self._atualizar_hora()

    # ------------------------------------------------------------------
    # Reading panel
    # ------------------------------------------------------------------

    def _construir_leitura(self, parent):
        self._painel_titulo(parent, "Leitura")

        frame_leitura = tk.Frame(parent, bg=CORES["branco"], relief=tk.GROOVE,
                                 bd=1, padx=12, pady=10)
        frame_leitura.pack(fill=tk.X, pady=(0, 4))

        # Row 1 – ID Banco + Posto
        row1 = tk.Frame(frame_leitura, bg=CORES["branco"])
        row1.pack(fill=tk.X, pady=(0, 6))

        tk.Label(row1, text="ID Banco:", bg=CORES["branco"],
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.entry_id_banco = tk.Entry(row1, width=22, font=("Segoe UI", 12))
        self.entry_id_banco.pack(side=tk.LEFT, padx=(4, 20))
        self.entry_id_banco.bind("<Return>", lambda e: self._registar_leitura())

        tk.Label(row1, text="Posto:", bg=CORES["branco"],
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.combo_posto = ttk.Combobox(
            row1, width=20, state="readonly",
            values=self.postos_list, font=("Segoe UI", 9)
        )
        if self.postos_list:
            self.combo_posto.current(0)
        self.combo_posto.pack(side=tk.LEFT, padx=(4, 0))

        # Separator
        tk.Frame(frame_leitura, height=1, bg=CORES["cinza_claro"]).pack(
            fill=tk.X, pady=4)

        # Controlo Funcional
        self._construir_controlo(
            frame_leitura,
            "Controlo Funcional",
            self.var_cf,
            self.defeitos_funcionais,
            self._set_cf,
            ("btn_cf_ok", "btn_cf_nok"),
            "frame_cf_defeitos",
            "group_cf",
        )

        # Separator
        tk.Frame(frame_leitura, height=1, bg=CORES["cinza_claro"]).pack(
            fill=tk.X, pady=4)

        # Controlo Aspeto
        self._construir_controlo(
            frame_leitura,
            "Controlo Aspeto",
            self.var_ca,
            self.defeitos_aspeto,
            self._set_ca,
            ("btn_ca_ok", "btn_ca_nok"),
            "frame_ca_defeitos",
            "group_ca",
        )

        # Separator
        tk.Frame(frame_leitura, height=1, bg=CORES["cinza_claro"]).pack(
            fill=tk.X, pady=4)

        # Registar button
        self.btn_registar = tk.Button(
            frame_leitura,
            text="  REGISTAR  ",
            font=("Segoe UI", 12, "bold"),
            bg=CORES["verde"], fg=CORES["branco"],
            relief=tk.FLAT,
            padx=20, pady=8,
            cursor="hand2",
            command=self._registar_leitura,
        )
        self.btn_registar.pack(pady=4)

    def _construir_controlo(
        self, parent, titulo, var, defeitos, set_fn,
        btn_attr, panel_attr, group_attr
    ):
        frame_section = tk.Frame(parent, bg=CORES["branco"])
        frame_section.pack(fill=tk.X, pady=2)

        frame_header = tk.Frame(frame_section, bg=CORES["branco"])
        frame_header.pack(fill=tk.X)

        tk.Label(
            frame_header, text=titulo,
            font=("Segoe UI", 10, "bold"),
            bg=CORES["branco"], fg=CORES["cinza_escuro"],
            width=20, anchor="w"
        ).pack(side=tk.LEFT, padx=(0, 12))

        btn_ok = tk.Button(
            frame_header,
            text="  OK  ",
            font=("Segoe UI", 10, "bold"),
            bg=CORES["cinza_claro"], fg=CORES["branco"],
            bd=2, relief=tk.RAISED,
            cursor="hand2",
            command=lambda: set_fn("OK"),
        )
        btn_ok.pack(side=tk.LEFT, padx=(0, 6))

        btn_nok = tk.Button(
            frame_header,
            text="  NOK  ",
            font=("Segoe UI", 10, "bold"),
            bg=CORES["cinza_claro"], fg=CORES["branco"],
            bd=2, relief=tk.RAISED,
            cursor="hand2",
            command=lambda: set_fn("NOK"),
        )
        btn_nok.pack(side=tk.LEFT)

        panel = DefeitosBtnGroup(frame_section, defeitos, CORES)
        # Do NOT pack yet — hidden initially

        setattr(self, btn_attr[0], btn_ok)
        setattr(self, btn_attr[1], btn_nok)
        setattr(self, panel_attr, panel)
        setattr(self, group_attr, panel)

    # ------------------------------------------------------------------
    # Control state setters
    # ------------------------------------------------------------------

    def _set_cf(self, valor: str):
        self.var_cf.set(valor)
        if valor == "OK":
            self.btn_cf_ok.configure(bg=CORES["verde"], relief=tk.SUNKEN)
            self.btn_cf_nok.configure(bg=CORES["cinza_claro"], relief=tk.RAISED)
            self.frame_cf_defeitos.pack_forget()
            self.group_cf.reset()
        elif valor == "NOK":
            self.btn_cf_nok.configure(bg=CORES["vermelho"], relief=tk.SUNKEN)
            self.btn_cf_ok.configure(bg=CORES["cinza_claro"], relief=tk.RAISED)
            self.frame_cf_defeitos.pack(fill=tk.X, pady=(0, 4))

    def _set_ca(self, valor: str):
        self.var_ca.set(valor)
        if valor == "OK":
            self.btn_ca_ok.configure(bg=CORES["verde"], relief=tk.SUNKEN)
            self.btn_ca_nok.configure(bg=CORES["cinza_claro"], relief=tk.RAISED)
            self.frame_ca_defeitos.pack_forget()
            self.group_ca.reset()
        elif valor == "NOK":
            self.btn_ca_nok.configure(bg=CORES["vermelho"], relief=tk.SUNKEN)
            self.btn_ca_ok.configure(bg=CORES["cinza_claro"], relief=tk.RAISED)
            self.frame_ca_defeitos.pack(fill=tk.X, pady=(0, 4))

    # ------------------------------------------------------------------
    # Form reset
    # ------------------------------------------------------------------

    def _resetar_form(self):
        self.entry_id_banco.delete(0, tk.END)

        # Reset CF
        self.var_cf.set("")
        self.btn_cf_ok.configure(bg=CORES["cinza_claro"], relief=tk.RAISED)
        self.btn_cf_nok.configure(bg=CORES["cinza_claro"], relief=tk.RAISED)
        self.frame_cf_defeitos.pack_forget()
        self.group_cf.reset()

        # Reset CA
        self.var_ca.set("")
        self.btn_ca_ok.configure(bg=CORES["cinza_claro"], relief=tk.RAISED)
        self.btn_ca_nok.configure(bg=CORES["cinza_claro"], relief=tk.RAISED)
        self.frame_ca_defeitos.pack_forget()
        self.group_ca.reset()

    # ------------------------------------------------------------------
    # Clock / duration timers
    # ------------------------------------------------------------------

    def _atualizar_hora(self):
        agora = datetime.now()
        self.label_hora.configure(text=agora.strftime("%H:%M:%S"))
        if self.sessao_iniciada and self.inicio_sessao:
            self.label_inicio_sessao.configure(
                text=f"Início: {self.inicio_sessao.strftime('%H:%M:%S')}"
            )
        else:
            self.label_inicio_sessao.configure(text="")
        self.root.after(1000, self._atualizar_hora)

    def _atualizar_duracao(self):
        if not self.sessao_iniciada:
            return
        if self.inicio_sessao:
            delta = datetime.now() - self.inicio_sessao
            total_seg = int(delta.total_seconds())
            h = total_seg // 3600
            m = (total_seg % 3600) // 60
            s = total_seg % 60
            self.label_duracao.configure(
                text=f"Duração: {h:02d}:{m:02d}:{s:02d}"
            )
        self._timer_duracao = self.root.after(1000, self._atualizar_duracao)

    # ------------------------------------------------------------------
    # Session management
    # ------------------------------------------------------------------

    def _iniciar_sessao(self):
        operador = self.entry_operador.get().strip()
        projeto  = self.combo_projeto.get().strip()
        if not operador:
            messagebox.showwarning("Atenção", "Por favor, insira o nome do operador.",
                                   parent=self.root)
            self.entry_operador.focus_set()
            return
        if not projeto:
            messagebox.showwarning("Atenção", "Por favor, selecione o projeto/linha.",
                                   parent=self.root)
            return

        # Close existing connection
        if self.db_con:
            try:
                self.db_con.close()
            except Exception:
                pass
            self.db_con = None

        self.operador       = operador
        self.projeto        = projeto
        self.turno          = self.combo_turno.get().strip()
        self.inicio_sessao  = datetime.now()
        self.sessao_id      = uuid4().hex

        # Paths
        log_dir, db_dir = carregar_caminhos()
        self.log_dir = log_dir
        self.db_dir  = db_dir
        os.makedirs(log_dir, exist_ok=True)
        os.makedirs(db_dir,  exist_ok=True)

        self.logfile = os.path.join(
            log_dir,
            f"{LOG_PREFIX}_{self.inicio_sessao.strftime('%Y-%m-%d')}.csv"
        )

        # DB
        self.db_path_atual = db_path(db_dir)
        self.db_con = db_connect(self.db_path_atual)
        _db_init_paragens(self.db_con)

        # State
        self.ultimas_leituras = []
        self.sessao_iniciada  = True
        self._atualizar_ultimas()

        # Disable session inputs
        self.entry_operador.configure(state=tk.DISABLED)
        self.combo_projeto.configure(state=tk.DISABLED)
        self.combo_turno.configure(state=tk.DISABLED)
        self.btn_iniciar.configure(state=tk.DISABLED)

        self.label_sessao.configure(
            text=f"Sessão ativa — {self.operador} | {self.projeto} | {self.turno}",
            fg=CORES["verde"]
        )
        self.label_inicio_sessao.configure(
            text=f"Início: {self.inicio_sessao.strftime('%H:%M:%S')}"
        )

        self._resetar_form()
        self._atualizar_resumo()
        self._atualizar_duracao()
        self.entry_id_banco.focus_set()

    def _terminar_sessao(self):
        if not self.sessao_iniciada:
            return

        # Export CSV of session
        try:
            self._exportar_csv_do_dia()
        except Exception as e:
            messagebox.showwarning("Exportação", f"Não foi possível exportar CSV:\n{e}",
                                   parent=self.root)

        self.sessao_iniciada = False

        # Cancel duration timer
        if self._timer_duracao:
            self.root.after_cancel(self._timer_duracao)
            self._timer_duracao = None

        # Re-enable UI
        self.entry_operador.configure(state=tk.NORMAL)
        self.combo_projeto.configure(state="readonly")
        self.combo_turno.configure(state="readonly")
        self.btn_iniciar.configure(state=tk.NORMAL)

        self.label_sessao.configure(text="Sem sessão ativa.", fg=CORES["cinza_claro"])
        self.label_duracao.configure(text="Duração: --:--:--")
        self.label_inicio_sessao.configure(text="")

        # Close DB
        if self.db_con:
            try:
                self.db_con.close()
            except Exception:
                pass
            self.db_con = None

        self._resetar_form()
        self._atualizar_resumo()

        messagebox.showinfo(
            "Sessão terminada",
            f"Sessão terminada.\nOperador: {self.operador}\n"
            f"Registos: {self.label_total['text']}",
            parent=self.root
        )
        self.entry_id_banco.focus_set()

    # ------------------------------------------------------------------
    # Reading registration
    # ------------------------------------------------------------------

    def _registar_leitura(self):
        if not self.sessao_iniciada:
            messagebox.showwarning("Atenção", "Inicie uma sessão primeiro.",
                                   parent=self.root)
            return

        id_banco = self.entry_id_banco.get().strip().upper()
        if not id_banco:
            messagebox.showwarning("Atenção", "Por favor, insira o ID do banco.",
                                   parent=self.root)
            self.entry_id_banco.focus_set()
            return

        posto = self.combo_posto.get().strip()
        if not posto:
            messagebox.showwarning("Atenção", "Por favor, selecione o posto.",
                                   parent=self.root)
            return

        cf = self.var_cf.get()
        if not cf:
            messagebox.showwarning("Atenção", "Por favor, selecione OK ou NOK para o Controlo Funcional.",
                                   parent=self.root)
            return

        if cf == "NOK":
            if self.group_cf.outro_ativo_sem_texto():
                messagebox.showwarning(
                    "Atenção",
                    "O campo 'Outro' do Controlo Funcional está ativo mas sem texto.\n"
                    "Preencha o texto ou desative 'Outro'.",
                    parent=self.root
                )
                return
            if not self.group_cf.tem_selecao():
                messagebox.showwarning(
                    "Atenção",
                    "Controlo Funcional NOK: selecione pelo menos um defeito.",
                    parent=self.root
                )
                return

        ca = self.var_ca.get()
        if not ca:
            messagebox.showwarning("Atenção", "Por favor, selecione OK ou NOK para o Controlo Aspeto.",
                                   parent=self.root)
            return

        if ca == "NOK":
            if self.group_ca.outro_ativo_sem_texto():
                messagebox.showwarning(
                    "Atenção",
                    "O campo 'Outro' do Controlo Aspeto está ativo mas sem texto.\n"
                    "Preencha o texto ou desative 'Outro'.",
                    parent=self.root
                )
                return
            if not self.group_ca.tem_selecao():
                messagebox.showwarning(
                    "Atenção",
                    "Controlo Aspeto NOK: selecione pelo menos um defeito.",
                    parent=self.root
                )
                return

        defeitos_func = self.group_cf.get_defeitos_str() if cf == "NOK" else ""
        cf_outro      = self.group_cf.get_outro()        if cf == "NOK" else ""
        defeitos_asp  = self.group_ca.get_defeitos_str() if ca == "NOK" else ""
        ca_outro      = self.group_ca.get_outro()        if ca == "NOK" else ""

        self._registar_item(id_banco, posto, cf, defeitos_func, cf_outro, ca, defeitos_asp, ca_outro)
        self._resetar_form()
        self._atualizar_ultimas()
        self._atualizar_resumo()
        self.entry_id_banco.focus_set()

    def _registar_item(
        self, id_banco: str, posto: str,
        cf: str, defeitos_func: str, cf_outro: str,
        ca: str, defeitos_asp: str, ca_outro: str
    ):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # DB INSERT
        row_id = None
        if self.db_con:
            try:
                cur = self.db_con.cursor()
                cur.execute(
                    f"""INSERT INTO {DB_TABLE}
                        (ts, operador, projeto, turno, id_banco, posto,
                         controlo_funcional, defeitos_funcionais, cf_outro,
                         controlo_aspeto, defeitos_aspeto, ca_outro, sessao_id)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (timestamp, self.operador, self.projeto, self.turno,
                     id_banco, posto, cf, defeitos_func, cf_outro,
                     ca, defeitos_asp, ca_outro, self.sessao_id)
                )
                self.db_con.commit()
                row_id = cur.lastrowid
            except Exception as e:
                messagebox.showerror("Erro DB", f"Erro ao gravar na base de dados:\n{e}",
                                     parent=self.root)

        # Prepend to list
        self.ultimas_leituras.insert(
            0,
            (row_id, id_banco, posto, timestamp, cf, defeitos_func, cf_outro,
             ca, defeitos_asp, ca_outro)
        )

        # CSV log
        file_is_new = not os.path.isfile(self.logfile)
        try:
            with open(self.logfile, "a", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                if file_is_new:
                    writer.writerow([
                        "Data", "Operador", "Projeto", "Turno",
                        "ID Banco", "Posto",
                        "Controlo Funcional", "Defeitos Funcionais", "CF Outro",
                        "Controlo Aspeto", "Defeitos Aspeto", "CA Outro"
                    ])
                writer.writerow([
                    timestamp, self.operador, self.projeto, self.turno,
                    id_banco, posto,
                    cf, defeitos_func, cf_outro,
                    ca, defeitos_asp, ca_outro
                ])
                f.flush()
                os.fsync(f.fileno())
        except Exception as e:
            messagebox.showwarning("Aviso CSV", f"Não foi possível gravar no CSV:\n{e}",
                                   parent=self.root)

    # ------------------------------------------------------------------
    # Ultimas leituras panel
    # ------------------------------------------------------------------

    def _construir_ultimas(self, parent):
        self._painel_titulo(parent, "Leituras da sessão")

        frame = tk.Frame(parent, bg=CORES["branco"], relief=tk.GROOVE, bd=1)
        frame.pack(fill=tk.X, pady=(0, 4))

        scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
        self.list_ultimas = tk.Listbox(
            frame,
            height=8,
            font=("Consolas", 9),
            yscrollcommand=scrollbar.set,
            selectmode=tk.SINGLE,
            activestyle="dotbox",
        )
        scrollbar.configure(command=self.list_ultimas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.list_ultimas.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_elim = tk.Button(
            frame,
            text="Eliminar leitura selecionada",
            font=("Segoe UI", 9),
            bg=CORES["vermelho"], fg=CORES["branco"],
            relief=tk.FLAT, cursor="hand2",
            padx=8, pady=4,
            command=self._eliminar_leitura,
        )
        btn_elim.pack(side=tk.BOTTOM, anchor="w", padx=8, pady=6)

    def _atualizar_ultimas(self):
        self.list_ultimas.delete(0, tk.END)
        total = len(self.ultimas_leituras)
        for ordem_rev, item in enumerate(self.ultimas_leituras):
            (row_id, id_banco, posto, ts, cf, defeitos_func_str,
             cf_outro, ca, defeitos_asp_str, ca_outro) = item
            ordem = total - ordem_rev
            hora = ts[11:19] if len(ts) >= 19 else ts

            # CF detail
            cf_parts = []
            if cf == "NOK":
                if defeitos_func_str:
                    cf_parts.append(defeitos_func_str)
                if cf_outro:
                    cf_parts.append(f"Outro:{cf_outro}")
            cf_detail = f" [{', '.join(cf_parts)}]" if cf_parts else ""

            # CA detail
            ca_parts = []
            if ca == "NOK":
                if defeitos_asp_str:
                    ca_parts.append(defeitos_asp_str)
                if ca_outro:
                    ca_parts.append(f"Outro:{ca_outro}")
            ca_detail = f" [{', '.join(ca_parts)}]" if ca_parts else ""

            linha = (
                f"{hora}  {ordem}. {id_banco} | {posto} | "
                f"CF:{cf}{cf_detail} | CA:{ca}{ca_detail}"
            )
            self.list_ultimas.insert(tk.END, linha)

    # ------------------------------------------------------------------
    # Resumo panel
    # ------------------------------------------------------------------

    def _construir_resumo(self, parent):
        self._painel_titulo(parent, "Resumo da sessão")

        frame_txt = tk.Frame(parent, bg=CORES["branco"], relief=tk.GROOVE, bd=1)
        frame_txt.pack(fill=tk.BOTH, expand=True, pady=(0, 4))

        scrollbar_r = tk.Scrollbar(frame_txt, orient=tk.VERTICAL)
        self.text_resumo = tk.Text(
            frame_txt,
            width=34,
            height=20,
            font=("Consolas", 10),
            state=tk.DISABLED,
            wrap=tk.WORD,
            yscrollcommand=scrollbar_r.set,
        )
        scrollbar_r.configure(command=self.text_resumo.yview)
        scrollbar_r.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_resumo.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._painel_titulo(parent, "Relatórios")
        btn_csv = tk.Button(
            parent,
            text="  Abrir exportação CSV  ",
            font=("Segoe UI", 9),
            bg=CORES["azul_claro"], fg=CORES["branco"],
            relief=tk.FLAT, cursor="hand2",
            padx=8, pady=4,
            command=self._abrir_janela_exportacao_csv_db,
        )
        btn_csv.pack(pady=6, anchor="w", padx=4)

    def _atualizar_resumo(self):
        n_bancos = len(self.ultimas_leituras)

        cf_ok  = sum(1 for item in self.ultimas_leituras if item[4] == "OK")
        cf_nok = sum(1 for item in self.ultimas_leituras if item[4] == "NOK")
        ca_ok  = sum(1 for item in self.ultimas_leituras if item[7] == "OK")
        ca_nok = sum(1 for item in self.ultimas_leituras if item[7] == "NOK")

        # Count defects
        cf_contagem: dict = {}
        ca_contagem: dict = {}
        for item in self.ultimas_leituras:
            (_, _, _, _, cf, defeitos_func_str, cf_outro,
             ca, defeitos_asp_str, ca_outro) = item
            if cf == "NOK":
                for d in defeitos_func_str.split(","):
                    d = d.strip()
                    if d:
                        cf_contagem[d] = cf_contagem.get(d, 0) + 1
                if cf_outro:
                    chave = f"Outro: {cf_outro}"
                    cf_contagem[chave] = cf_contagem.get(chave, 0) + 1
            if ca == "NOK":
                for d in defeitos_asp_str.split(","):
                    d = d.strip()
                    if d:
                        ca_contagem[d] = ca_contagem.get(d, 0) + 1
                if ca_outro:
                    chave = f"Outro: {ca_outro}"
                    ca_contagem[chave] = ca_contagem.get(chave, 0) + 1

        lines = []
        lines.append(f"Bancos registados: {n_bancos}")
        lines.append("")
        lines.append("Controlo Funcional")
        lines.append(f"  OK: {cf_ok}   NOK: {cf_nok}")
        for d, cnt in sorted(cf_contagem.items(), key=lambda x: -x[1]):
            lines.append(f"    {d}: {cnt}")
        lines.append("")
        lines.append("Controlo Aspeto")
        lines.append(f"  OK: {ca_ok}   NOK: {ca_nok}")
        for d, cnt in sorted(ca_contagem.items(), key=lambda x: -x[1]):
            lines.append(f"    {d}: {cnt}")

        self.text_resumo.configure(state=tk.NORMAL)
        self.text_resumo.delete("1.0", tk.END)
        self.text_resumo.insert(tk.END, "\n".join(lines))
        self.text_resumo.configure(state=tk.DISABLED)

        # Footer labels
        total_dia = self._total_do_dia()
        self.label_total.configure(text=f"Total do dia: {total_dia}")
        self.label_leituras_sessao.configure(text=f"Sessão: {n_bancos}")

    # ------------------------------------------------------------------
    # Footer
    # ------------------------------------------------------------------

    def _construir_footer(self):
        footer = tk.Frame(self.root, bg=CORES["branco"], relief=tk.GROOVE, bd=1, pady=4)
        footer.pack(fill=tk.X, side=tk.BOTTOM)

        self.label_total = tk.Label(
            footer, text="Total do dia: 0",
            font=("Segoe UI", 9), bg=CORES["branco"], fg=CORES["cinza_escuro"]
        )
        self.label_total.pack(side=tk.LEFT, padx=12)

        self.label_leituras_sessao = tk.Label(
            footer, text="Sessão: 0",
            font=("Segoe UI", 9), bg=CORES["branco"], fg=CORES["cinza_escuro"]
        )
        self.label_leituras_sessao.pack(side=tk.LEFT, padx=12)

        self.label_duracao = tk.Label(
            footer, text="Duração: --:--:--",
            font=("Segoe UI", 9), bg=CORES["branco"], fg=CORES["cinza_escuro"]
        )
        self.label_duracao.pack(side=tk.LEFT, padx=12)

        self.btn_terminar = tk.Button(
            footer,
            text="Terminar sessão",
            font=("Segoe UI", 9, "bold"),
            bg=CORES["vermelho"], fg=CORES["branco"],
            relief=tk.FLAT, cursor="hand2",
            padx=10, pady=4,
            command=self._terminar_sessao,
        )
        self.btn_terminar.pack(side=tk.RIGHT, padx=12)

        footer2 = tk.Frame(self.root, bg=CORES["fundo"])
        footer2.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(
            footer2,
            text="Desenvolvido por Bruno Santos - 2026 - v1.0",
            font=("Segoe UI", 8),
            bg=CORES["fundo"], fg=CORES["cinza_claro"]
        ).pack(side=tk.RIGHT, padx=8, pady=2)

    # ------------------------------------------------------------------
    # Total do dia
    # ------------------------------------------------------------------

    def _total_do_dia(self) -> int:
        hoje = date.today().strftime("%Y-%m-%d")
        # Try direct DB
        if self.db_con:
            try:
                cur = self.db_con.cursor()
                cur.execute(
                    f"SELECT COUNT(*) FROM {DB_TABLE} WHERE ts LIKE ?",
                    (f"{hoje}%",)
                )
                row = cur.fetchone()
                return row[0] if row else 0
            except Exception:
                pass
        # Try to open DB from config
        try:
            _, db_dir = carregar_caminhos()
            dp = db_path(db_dir)
            if os.path.isfile(dp):
                con = db_connect(dp)
                cur = con.cursor()
                cur.execute(
                    f"SELECT COUNT(*) FROM {DB_TABLE} WHERE ts LIKE ?",
                    (f"{hoje}%",)
                )
                row = cur.fetchone()
                con.close()
                return row[0] if row else 0
        except Exception:
            pass
        # Fallback: count from session list
        return sum(
            1 for item in self.ultimas_leituras
            if item[3].startswith(hoje)
        )

    # ------------------------------------------------------------------
    # Delete reading
    # ------------------------------------------------------------------

    def _eliminar_leitura(self):
        if not self.sessao_iniciada:
            messagebox.showwarning("Atenção", "Sem sessão ativa.", parent=self.root)
            return
        sel = self.list_ultimas.curselection()
        if not sel:
            messagebox.showwarning("Atenção", "Selecione uma leitura para eliminar.",
                                   parent=self.root)
            return
        idx = sel[0]
        if idx >= len(self.ultimas_leituras):
            return
        item = self.ultimas_leituras[idx]
        (row_id, id_banco, posto, ts, cf, defeitos_func_str,
         cf_outro, ca, defeitos_asp_str, ca_outro) = item

        if not messagebox.askyesno(
            "Confirmar eliminação",
            f"Eliminar leitura:\n{id_banco} | {posto} | {ts}?",
            parent=self.root
        ):
            return

        # DB DELETE
        if self.db_con and row_id is not None:
            try:
                self.db_con.execute(f"DELETE FROM {DB_TABLE} WHERE id=?", (row_id,))
                self.db_con.commit()
            except Exception as e:
                messagebox.showerror("Erro DB", f"Erro ao eliminar:\n{e}", parent=self.root)

        # Remove from list
        self.ultimas_leituras.pop(idx)

        # Best-effort: rewrite CSV removing matching row
        if self.logfile and os.path.isfile(self.logfile):
            try:
                with open(self.logfile, "r", encoding="utf-8-sig", newline="") as f:
                    rows = list(csv.reader(f, delimiter=";"))
                new_rows = []
                for row in rows:
                    # Header row or non-matching data rows are kept
                    if len(row) < 2:
                        new_rows.append(row)
                        continue
                    # row[0] = Data (timestamp), row[4] = ID Banco
                    if row[0] == ts and row[4] == id_banco:
                        continue  # skip this row (first match)
                    new_rows.append(row)
                with open(self.logfile, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f, delimiter=";")
                    writer.writerows(new_rows)
                    f.flush()
                    os.fsync(f.fileno())
            except Exception:
                pass

        self._atualizar_ultimas()
        self._atualizar_resumo()

    # ------------------------------------------------------------------
    # DB connection for reports
    # ------------------------------------------------------------------

    def _obter_conexao_db_relatorio(self) -> Tuple[sqlite3.Connection, bool]:
        """Returns (connection, should_close). If reusing existing, should_close=False."""
        if self.db_con:
            return self.db_con, False
        _, db_dir = carregar_caminhos()
        dp = db_path(db_dir)
        if not os.path.isfile(dp):
            raise FileNotFoundError(f"Base de dados não encontrada: {dp}")
        con = db_connect(dp)
        return con, True

    # ------------------------------------------------------------------
    # CSV report generation
    # ------------------------------------------------------------------

    def _gerar_relatorio_csv_db(
        self, data_ini: str, data_fim: str
    ) -> Tuple[str, str, int]:
        """
        Generate two CSV files (detalhe + totais) in log_dir.
        Returns (detalhe_path, totais_path, total_rows).
        """
        con, should_close = self._obter_conexao_db_relatorio()
        try:
            cur = con.cursor()
            cur.execute(
                f"""SELECT id, ts, operador, projeto, turno, id_banco, posto,
                           controlo_funcional, defeitos_funcionais, cf_outro,
                           controlo_aspeto, defeitos_aspeto, ca_outro
                    FROM {DB_TABLE}
                    WHERE ts >= ? AND ts <= ?
                    ORDER BY ts""",
                (f"{data_ini} 00:00:00", f"{data_fim} 23:59:59")
            )
            rows = cur.fetchall()
        finally:
            if should_close:
                con.close()

        total_rows = len(rows)
        defeitos_func_list = carregar_defeitos_funcionais()
        defeitos_asp_list  = carregar_defeitos_aspeto()

        # File name
        if data_ini == data_fim:
            base_name = f"relatorio_paragem_{data_ini}"
        else:
            base_name = f"relatorio_paragem_{data_ini}_a_{data_fim}"

        log_dir = self.log_dir if self.log_dir else APP_DIR
        os.makedirs(log_dir, exist_ok=True)

        detalhe_path = os.path.join(log_dir, f"{base_name}_detalhe.csv")
        totais_path  = os.path.join(log_dir, f"{base_name}_totais.csv")

        # ---- Detalhe CSV (WIDE FORMAT) ----
        header_detalhe = (
            ["Data", "Operador", "Projeto", "Turno", "ID Banco", "Posto", "Controlo Funcional"]
            + [f"CF: {d}" for d in defeitos_func_list]
            + ["CF Outro", "Controlo Aspeto"]
            + [f"CA: {d}" for d in defeitos_asp_list]
            + ["CA Outro"]
        )

        with open(detalhe_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(header_detalhe)
            for row in rows:
                (_, ts, operador, projeto, turno, id_banco, posto,
                 cf, def_func_str, cf_outro,
                 ca, def_asp_str, ca_outro) = row

                def_func_set = set(
                    d.strip() for d in (def_func_str or "").split(",") if d.strip()
                )
                def_asp_set = set(
                    d.strip() for d in (def_asp_str or "").split(",") if d.strip()
                )

                cf_cols = ["X" if d in def_func_set else "" for d in defeitos_func_list]
                ca_cols = ["X" if d in def_asp_set  else "" for d in defeitos_asp_list]

                writer.writerow(
                    [ts, operador, projeto, turno, id_banco, posto, cf or ""]
                    + cf_cols
                    + [cf_outro or "", ca or ""]
                    + ca_cols
                    + [ca_outro or ""]
                )

        # ---- Totais CSV ----
        total_bancos = total_rows
        cf_nok_total = sum(1 for r in rows if r[7] == "NOK")
        ca_nok_total = sum(1 for r in rows if r[10] == "NOK")

        cf_contagem: dict = {}
        ca_contagem: dict = {}
        for row in rows:
            def_func_str = row[8] or ""
            cf_outro_txt = row[9] or ""
            def_asp_str  = row[11] or ""
            ca_outro_txt = row[12] or ""
            if row[7] == "NOK":
                for d in def_func_str.split(","):
                    d = d.strip()
                    if d:
                        cf_contagem[d] = cf_contagem.get(d, 0) + 1
                if cf_outro_txt:
                    chave = f"Outro: {cf_outro_txt}"
                    cf_contagem[chave] = cf_contagem.get(chave, 0) + 1
            if row[10] == "NOK":
                for d in def_asp_str.split(","):
                    d = d.strip()
                    if d:
                        ca_contagem[d] = ca_contagem.get(d, 0) + 1
                if ca_outro_txt:
                    chave = f"Outro: {ca_outro_txt}"
                    ca_contagem[chave] = ca_contagem.get(chave, 0) + 1

        with open(totais_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["Tipo", "Defeito", "Total"])
            writer.writerow(["RESUMO", "Total Bancos", total_bancos])
            writer.writerow(["RESUMO", "CF NOK", cf_nok_total])
            writer.writerow(["RESUMO", "CA NOK", ca_nok_total])
            writer.writerow([])
            writer.writerow(["Controlo Funcional", "Defeito", "Ocorrências"])
            for d, cnt in sorted(cf_contagem.items(), key=lambda x: -x[1]):
                writer.writerow(["CF", d, cnt])
            writer.writerow([])
            writer.writerow(["Controlo Aspeto", "Defeito", "Ocorrências"])
            for d, cnt in sorted(ca_contagem.items(), key=lambda x: -x[1]):
                writer.writerow(["CA", d, cnt])

        return detalhe_path, totais_path, total_rows

    # ------------------------------------------------------------------
    # Export today
    # ------------------------------------------------------------------

    def _exportar_csv_do_dia(self):
        hoje = date.today().strftime("%Y-%m-%d")
        try:
            detalhe, totais, n = self._gerar_relatorio_csv_db(hoje, hoje)
            messagebox.showinfo(
                "Exportação concluída",
                f"Exportados {n} registos.\n\nDetalhe:\n{detalhe}\n\nTotais:\n{totais}",
                parent=self.root
            )
        except FileNotFoundError as e:
            messagebox.showwarning("Exportação", str(e), parent=self.root)
        except Exception as e:
            messagebox.showerror("Erro exportação", str(e), parent=self.root)

    # ------------------------------------------------------------------
    # Export dialog
    # ------------------------------------------------------------------

    def _abrir_janela_exportacao_csv_db(self):
        win = tk.Toplevel(self.root)
        win.title("Exportar CSV")
        win.resizable(False, False)
        win.grab_set()
        win.configure(bg=CORES["fundo"])

        win.bind("<Escape>", lambda e: win.destroy())

        tk.Label(win, text="Exportar relatório CSV",
                 font=("Segoe UI", 11, "bold"),
                 bg=CORES["fundo"], fg=CORES["cinza_escuro"]).pack(pady=(14, 6), padx=20)

        # Radio mode
        var_modo = tk.StringVar(value="dia")
        frame_radio = tk.Frame(win, bg=CORES["fundo"])
        frame_radio.pack(padx=20, pady=4, anchor="w")
        tk.Radiobutton(
            frame_radio, text="Um dia", variable=var_modo, value="dia",
            bg=CORES["fundo"], font=("Segoe UI", 9),
            command=lambda: _on_modo()
        ).pack(side=tk.LEFT, padx=(0, 20))
        tk.Radiobutton(
            frame_radio, text="Intervalo de dias", variable=var_modo, value="intervalo",
            bg=CORES["fundo"], font=("Segoe UI", 9),
            command=lambda: _on_modo()
        ).pack(side=tk.LEFT)

        hoje_str = date.today().strftime("%Y-%m-%d")

        # Date entries
        frame_datas = tk.Frame(win, bg=CORES["fundo"])
        frame_datas.pack(padx=20, pady=4, anchor="w")

        tk.Label(frame_datas, text="Data inicial (AAAA-MM-DD):",
                 bg=CORES["fundo"], font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=4)
        entry_ini = tk.Entry(frame_datas, width=16, font=("Segoe UI", 9))
        entry_ini.insert(0, hoje_str)
        entry_ini.grid(row=0, column=1, padx=(8, 0), pady=4)

        tk.Label(frame_datas, text="Data final (AAAA-MM-DD):",
                 bg=CORES["fundo"], font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=4)
        entry_fim = tk.Entry(frame_datas, width=16, font=("Segoe UI", 9),
                             state=tk.DISABLED)
        entry_fim.grid(row=1, column=1, padx=(8, 0), pady=4)

        def _on_modo():
            if var_modo.get() == "intervalo":
                entry_fim.configure(state=tk.NORMAL)
                if not entry_fim.get().strip():
                    entry_fim.insert(0, hoje_str)
            else:
                entry_fim.configure(state=tk.DISABLED)

        lbl_resultado = tk.Label(win, text="", bg=CORES["fundo"],
                                 font=("Segoe UI", 9), fg=CORES["cinza_escuro"],
                                 wraplength=340, justify=tk.LEFT)
        lbl_resultado.pack(padx=20, pady=4, anchor="w")

        def _exportar():
            d_ini = entry_ini.get().strip()
            if var_modo.get() == "dia":
                d_fim = d_ini
            else:
                d_fim = entry_fim.get().strip()
                if not d_fim:
                    d_fim = d_ini
            try:
                detalhe, totais, n = self._gerar_relatorio_csv_db(d_ini, d_fim)
                lbl_resultado.configure(
                    fg=CORES["verde"],
                    text=f"OK — {n} registos exportados.\nDetalhe: {detalhe}\nTotais: {totais}"
                )
            except Exception as e:
                lbl_resultado.configure(fg=CORES["vermelho"], text=f"Erro: {e}")

        frame_btns = tk.Frame(win, bg=CORES["fundo"])
        frame_btns.pack(pady=10, padx=20, anchor="e")

        btn_exp = tk.Button(
            frame_btns, text="Exportar",
            font=("Segoe UI", 9, "bold"),
            bg=CORES["azul"], fg=CORES["branco"],
            relief=tk.FLAT, padx=12, pady=4, cursor="hand2",
            command=_exportar,
        )
        btn_exp.pack(side=tk.LEFT, padx=(0, 8))

        btn_fechar = tk.Button(
            frame_btns, text="Fechar",
            font=("Segoe UI", 9),
            bg=CORES["cinza_claro"], fg=CORES["branco"],
            relief=tk.FLAT, padx=12, pady=4, cursor="hand2",
            command=win.destroy,
        )
        btn_fechar.pack(side=tk.LEFT)

        win.bind("<Return>", lambda e: _exportar())

    # ------------------------------------------------------------------
    # Main loop
    # ------------------------------------------------------------------

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = ParagemDefeitoApp()
    app.run()
