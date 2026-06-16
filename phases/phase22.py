# corporate_culture_research.py
from __future__ import annotations
import tkinter as tk
import os
from tkinter import messagebox
from dataclasses import dataclass
from ui.ui_styles import S, FONTS, PRIMARY_BG
from ui.ui_components import clear_frame, create_submit_button, show_incomplete_warning
from utils.session_manager import save_session, load_session, mark_session_complete
import utils.write_assessments_to_excel as write_assessments_to_excel

# ---------- Data ----------
@dataclass(frozen=True)
class Group:
    id: int
    name: str
    desc: str
    stmts: list[str]

GROUPS = [
    Group(1, "Mensgerichte cultuur",
          "Typerend voor zorginstellingen en servicegerichte organisaties.",
          [
              "Er bestaan gemeenschappelijke waarden en doelstellingen.",
              "Onderlinge samenhang (wij-gevoel). Participatieve instelling, teamwork, het beste uit elkaar halen.",
              "Klanten worden als partners beschouwd.",
              "Regels en procedures zijn ondergeschikt aan het gevoel een team te zijn.",
          ]),
    Group(2, "Innovatieve cultuur",
          "Typerend voor softwarebedrijven, luchtvaart, ruimtevaart e.d.",
          [
              "Snel reageren op snel veranderende omstandigheden.",
              "Centraal staan innovatie en vernieuwing zoals de ontwikkeling van nieuwe diensten en producten. Voorbereiding op de toekomst.",
              "Medewerkers worden gestimuleerd om creatief en flexibel te zijn.",
              "De organisatie is flexibel en kan snel een nieuwe vorm aannemen.",
          ]),
    Group(3, "Beheersgerichte cultuur",
          "Typerend voor overheids- of onderwijsinstellingen.",
          [
              "Het centraal staan van procedures en regels.",
              "De leidinggevenden coördineren en organiseren.",
              "Langetermijndoelen zijn stabiliteit (in stand houden van de organisatie), efficiëntie en voorspelbaarheid.",
              "De organisatie kenmerkt zich door trage besluitvormingsprocessen",
          ]),
    Group(4, "Resultaatgerichte cultuur",
          "Richting externe transacties; concurreren met gelijkaardige organisaties.",
          [
              "Centrale waarden zijn productiviteit, resultaten, winst en taakgerichtheid.",
              "De externe positionering wordt benadrukt om de concurrentiepositie te versterken.",
              "Er is een duidelijk doel aanwezig en er wordt gebruikgemaakt van een agressieve strategie.",
              "Leidinggevenden zijn veeleisend en leggen de nadruk op de marktleider zijn.",
          ]),
]

LIKERT = ["1", "2", "3", "4", "5"]


# ---------- Helpers ----------
def scrollable(parent: tk.Widget) -> tuple[tk.Frame, tk.Canvas, tk.Frame]:
    wrap = tk.Frame(parent, bg=S["bg"])
    wrap.pack(fill="both", expand=True)

    c = tk.Canvas(wrap, bg=S["bg"], highlightthickness=0)
    sb = tk.Scrollbar(wrap, orient="vertical", command=c.yview)
    c.configure(yscrollcommand=sb.set)
    c.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    inner = tk.Frame(c, bg=S["bg"])
    win = c.create_window((0, 0), window=inner, anchor="nw")

    inner.bind("<Configure>", lambda _e: c.configure(scrollregion=c.bbox("all")))
    c.bind("<Configure>", lambda e: c.itemconfig(win, width=e.width))
    c.bind_all("<MouseWheel>", lambda e: c.yview_scroll(-int(e.delta / 120), "units"))

    return wrap, c, inner

def make_likert_buttons(parent: tk.Widget, var: tk.StringVar, bg: str) -> tk.Frame:
    f = tk.Frame(parent, bg=bg)
    buttons = {}

    def refresh():
        cur = var.get()
        for v, b in buttons.items():
            if v == cur:
                b.config(relief="sunken", bg=S["btn_on"], fg="white")
            else:
                b.config(relief="raised", bg=S["btn"], fg="black")

    def choose(v: str):
        var.set(v)
        refresh()

    for i, v in enumerate(LIKERT):
        b = tk.Button(
            f,
            text=v,
            width=3,
            bd=1,
            font=S["f"],
            cursor="hand2",
            command=lambda x=v: choose(x),
            bg=S["btn"]
        )
        b.grid(row=0, column=i, padx=3)
        buttons[v] = b

    refresh()
    return f

# ---------- Page ----------
class Culture22Page(tk.Frame):
    def __init__(self, parent: tk.Widget, navigate):
        super().__init__(parent, bg=S["bg"])
        self.navigate = navigate
        self.vars: dict[tuple[int, int], tk.StringVar] = {}
        self.sub_lbl: dict[int, tk.Label] = {}
        
        # Get client context from parent frame
        self.parent_frame = parent
        self.client = getattr(parent, "current_assessment_client", None)
        
        self.build()

    def subtotal(self, gid: int) -> int:
        tot = 0
        for i in range(1, 5):
            v = self.vars[(gid, i)].get().strip()
            if v.isdigit():
                tot += int(v)
        return tot

    def update_subtotal(self, gid: int):
        self.sub_lbl[gid].config(text=str(self.subtotal(gid)))

    def build(self):
        # Load saved session if it exists
        if self.client:
            saved_session = load_session(str(self.client["id"]), self.client["name"], "phase2.2")
            self.saved_answers = saved_session.get("answers", {}) if saved_session else {}
        else:
            self.saved_answers = {}

        _, _, inner = scrollable(self)

        tk.Label(inner, text="Fase 2.2 – Cultuur | Beoordeling van stellingen (1–4)",
                 bg=S["bg"], font=S["f_title"], anchor="w").pack(fill="x", padx=20, pady=(15, 5))
        tk.Label(inner, text="Beoordeel elke stelling: 1=oneens … 5=volledig eens.",
                 bg=S["bg"], font=S["f_sub"], anchor="w").pack(fill="x", padx=20, pady=(0, 12))

        outer = tk.Frame(inner, bg=S["bg"])
        outer.pack(fill="both", expand=True, padx=20)

        left = tk.Frame(outer, bg=S["bg"])
        right = tk.Frame(outer, bg=S["bg"])
        left.pack(side="left", fill="both", expand=True)
        right.pack(side="right", fill="y", padx=(15, 0))

        # header
        h = tk.Frame(left, bg=S["dark"])
        h.pack(fill="x")
        tk.Label(h, text="Cultuur", bg=S["dark"], fg="white", font=S["f_b"], width=8, anchor="w", padx=10)\
            .grid(row=0, column=0, sticky="w")


        # tk.Label(h, text="Aspecten: stellingen", bg=S["dark"], fg="white", font=S["f_b"], anchor="w", padx=10)\
        #     .grid(row=0, column=1, sticky="w")

        body = tk.Frame(left, bg=S["bg"])
        body.pack(fill="both", expand=True)

        row_counter = 0
        for g in GROUPS:
            grp = tk.Frame(body, bg=S["yellow"])
            grp.pack(fill="x", pady=(8, 2))
            tk.Label(grp, text=str(g.id), bg=S["yellow"], font=S["f_b"], width=8)\
                .grid(row=0, column=0)
            tk.Label(grp, text=g.name, bg=S["yellow"], font=S["f_b"], anchor="w", padx=10)\
                .grid(row=0, column=1, sticky="w")

            for i, stmt in enumerate(g.stmts, start=1):
                row_counter += 1
                bg = S["odd"] if row_counter % 2 else S["even"]

                r = tk.Frame(body, bg=bg)
                r.pack(fill="x", pady=1)
                r.grid_columnconfigure(1, weight=1)

                tk.Label(r, text="", bg=bg, width=8).grid(row=0, column=0, sticky="w")
                tk.Label(
                    r,
                    text=f"{i}. {stmt}",
                    bg=bg,
                    font=S["f"],
                    anchor="w",
                    justify="left",
                    wraplength=680,
                    padx=10
                ).grid(row=0, column=1, sticky="w")

                v = tk.StringVar(value="")
                
                # Try to restore saved value
                saved_key = str((g.id, i))
                if saved_key in self.saved_answers:
                    saved_value = self.saved_answers[saved_key]
                    if saved_value:
                        v.set(saved_value)
                
                self.vars[(g.id, i)] = v
                v.trace_add("write", lambda *_a, gid=g.id: self.update_subtotal(gid))
                v.trace_add("write", lambda *_a: self.auto_save())

                make_likert_buttons(r, v, bg).grid(row=0, column=2, sticky="e", padx=(10, 25))
                tk.Label(r, text="", bg=bg, width=8).grid(row=0, column=3, sticky="e")

            # subtotal row
            sub = tk.Frame(body, bg=S["yellow"])
            sub.pack(fill="x", pady=(2, 6))
            sub.grid_columnconfigure(1, weight=1)

            tk.Label(sub, text="", bg=S["yellow"], width=8).grid(row=0, column=0)
            tk.Label(sub, text="Totaal score (4 stellingen)", bg=S["yellow"], font=S["f_b"],
                     anchor="w", padx=10).grid(row=0, column=1, sticky="w")
            tk.Label(sub, text="", bg=S["yellow"], width=22).grid(row=0, column=2)

            lbl = tk.Label(sub, text="0", bg=S["yellow"], font=S["f_b"], width=8, anchor="e", padx=10)
            lbl.grid(row=0, column=3, sticky="e")
            self.sub_lbl[g.id] = lbl
            self.update_subtotal(g.id)

        # right descriptions
        tk.Label(right, text="Cultuur – omschrijving", bg=S["bg"], font=S["f_b"], anchor="w")\
            .pack(fill="x", pady=(0, 8))
        for g in GROUPS:
            box = tk.Frame(right, bg="#f5f5f5", bd=1, relief="solid")
            box.pack(fill="x", pady=6)
            tk.Label(box, text=g.name, bg=S["yellow"], font=S["f_b"], pady=6).pack(fill="x")
            tk.Label(
                box,
                text=g.desc,
                bg="#f5f5f5",
                font=S["f"],
                wraplength=360,
                justify="left",
                anchor="w",
                padx=8,
                pady=8
            ).pack(fill="x")

        # submit
        btn_row = tk.Frame(inner, bg=S["bg"])
        btn_row.pack(fill="x", padx=20, pady=(12, 20))

        btn_submit = create_submit_button(
            btn_row,
            text="Opslaan en verder",
            command=self.submit
        )
        btn_submit.pack(side="right")


    def submit(self):
        """
        Collects Likert scale values for Phase 2.2 and writes them
        to the Excel sheet according to PHASE_2_2_MAPPING.
        """

        # check for missing answers
        missing = [
            (gid, i)
            for gid in range(1, 5)
            for i in range(1, 5)
            if not self.vars[(gid, i)].get().strip()
        ]
        if missing:
            show_incomplete_warning(len(missing), item_name="stellingen")
            return

        # prepare answers dict {(group_id, stmt_index): value}
        answers_to_excel = {k: v.get() for k, v in self.vars.items()}

        # get existing Excel file
        root = self.winfo_toplevel()
        excel_path = getattr(root, "results_excel_path", None)
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror(
                "Geen Excel-bestand",
                "Er is nog geen resultatenbestand aangemaakt (fase 1.1)."
            )
            return

        # write to Excel
        saved_path = write_assessments_to_excel.write_phase_2_2_to_excel(
            answers=answers_to_excel,
            excel_path=excel_path
        )

        if not saved_path:
            messagebox.showerror(
                "Fout",
                "Kon de antwoorden niet opslaan in Excel."
            )
            return

        # save path
        root.results_excel_path = saved_path

        # calculate totals per group
        self.cultuur_totals = {g.id: self.subtotal(g.id) for g in GROUPS}

        messagebox.showinfo(
            "Opgeslagen",
            f"Antwoorden succesvol opgeslagen in:\n{saved_path}"
        )

        # Mark session as complete
        if self.client:
            mark_session_complete(str(self.client["id"]), self.client["name"], "phase2.2")

        # next phase
        self.navigate("phase2.3")

    def auto_save(self):
        """Auto-save current answers to session file."""
        if not self.client:
            return
        
        # Collect current answers
        answers_to_save = {}
        for (gid, i), var in self.vars.items():
            try:
                value = var.get().strip()
                if value:
                    answers_to_save[str((gid, i))] = value
            except Exception:
                pass
        
        # Save to session
        save_session(
            str(self.client["id"]),
            self.client["name"],
            "phase2.2",
            0,
            answers_to_save,
            16  # 4 groups × 4 statements each
        )


# -------------------- BUILDER FUNCTION --------------------
def build_cultuur_page(parent_frame: tk.Frame, navigate) -> None:
    clear_frame(parent_frame)
    page = Culture22Page(parent_frame, navigate)
    page.pack(fill="both", expand=True)