import tkinter as tk
from tkinter import messagebox
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.write_assessments_to_excel import write_assessment_answers_to_excel
from utils.session_manager import save_session, load_session, mark_session_complete
from ui.ui_components import clear_frame, add_nav_buttons
from ui.ui_styles import (
    DARK_GREY,
    LIGHT_GREY,
    TEXT_FONT,
    FONTS,
    PRIMARY_BG,
    SECONDARY_BG,
    ACCENT_PRIMARY_COLOR,
    BUTTON_PADX,
    BUTTON_PADY,
    BUTTON_CONTAINER_PADX,
    LIKERT_SELECTED_BG,
    LIKERT_DEFAULT_BG,
)

# --------- Likert-scale: 1 t/m 5 ---------
LIKERT_OPTIONS = [
    ("1", "Oneens"),
    ("2", "Deels oneens"),
    ("3", "Neutraal"),
    ("4", "Deels eens"),
    ("5", "Volledig eens"),
]

# --------- Big Five items (1 t/m 50) ---------
BIG_FIVE_ITEMS = [
    (1, "Ik ben het middelpunt van het feest."),
    (2, "Ik voel me weinig bezorgd om anderen."),
    (3, "Ik ben altijd voorbereid."),
    (4, "Ik raak niet snel gestrest."),
    (5, "Ik heb een rijk vocabulaire."),
    (6, "Ik praat niet veel."),
    (7, "Ik ben oprecht geïnteresseerd in anderen."),
    (8, "Ik vergeet met regelmaat mijn spullen."),
    (9, "Ik ben meestal ontspannen."),
    (10, "Ik heb moeite om abstracte ideeën te begrijpen."),
    (11, "Ik voel me comfortabel bij mensen."),
    (12, "Ik beledig mensen"),
    (13, "Ik besteed aandacht aan details."),
    (14, "Ik maak me niet vaak zorgen."),
    (15, "Ik heb een levendige verbeelding."),
    (16, "Ik blijf liever op de achtergrond."),
    (17, "Ik sympathiseer met de gevoelens van anderen."),
    (18, "Ik maak al snel een chaos van een situatie."),
    (19, "Ik voel me vaak droevig."),
    (20, "Ik ben niet geïnteresseerd in abstracte ideeën."),
    (21, "Ik begin gesprekken."),
    (22, "Ik ben niet geïnteresseerd in de problemen van anderen."),
    (23, "Ik krijg een klus meteen gedaan."),
    (24, "Ik raak niet snel afgeleid."),
    (25, "Ik heb uitstekende ideeën."),
    (26, "Ik heb over het algemeen weinig te zeggen."),
    (27, "Ik heb een zacht hart."),
    (28, "Ik vergeet vaak om objecten op de juiste plaats terug te zetten."),
    (29, "Ik raak niet gemakkelijk overstuur."),
    (30, "Ik heb geen goede verbeelding."),
    (31, "Ik praat met veel verschillende mensen op een feest."),
    (32, "Ik ben niet echt geïnteresseerd in anderen."),
    (33, "Ik ben gesteld op orde."),
    (34, "Mijn humeur verandert niet vaak."),
    (35, "Ik begrijp zaken snel."),
    (36, "Ik houd er niet van de aandacht op mezelf te vestigen."),
    (37, "Ik neem de tijd voor anderen."),
    (38, "Ik onttrek me van mijn plichten."),
    (39, "Ik ervaar vrijwel geen stemmingswisselingen."),
    (40, "Ik gebruik moeilijke woorden."),
    (41, "Ik heb er geen probleem mee om het centrum van de aandacht te zijn."),
    (42, "Ik voel de emoties van anderen."),
    (43, "Ik volg graag schema's."),
    (44, "Ik raak niet snel geïrriteerd."),
    (45, "Ik besteed tijd aan het reflecteren op zaken."),
    (46, "Ik ben stil rondom vreemden."),
    (47, "Ik laat mensen zich op hun gemak voelen."),
    (48, "Ik stel hoge eisen aan de kwaliteit van mijn werk."),
    (49, "Ik voel me zelden droevig."),
    (50, "Ik zit vol met ideeën."),
]

# --- Table column headers ---

def get_assessment_results(parent_frame: tk.Frame) -> dict[int, int | None]:
    """
    Haal alle ingevulde antwoorden op uit parent_frame.assessment_vars.
    Geeft een dict terug: {nummer: int(waarde) of None als leeg}
    """
    results: dict[int, int | None] = {}
    vars_dict = getattr(parent_frame, "assessment_vars", {})

    for nummer, var in vars_dict.items():
        value = var.get().strip()
        results[nummer] = int(value) if value else None

    return results


def make_likert_row(parent: tk.Frame, nummer: int, stelling: str, var: tk.StringVar) -> None:
    """Eén rij: nummer + stelling + 1..5 knoppen (rechts uitgelijnd)."""

    row = tk.Frame(parent, bg=LIGHT_GREY)
    row.pack(fill="x", pady=1)

    # make sure that column 2 (position) stretches, so that column 3 is always on the right
    row.grid_columnconfigure(1, weight=1)

    # Number (yellow)
    num_label = tk.Label(
        row,
        text=str(nummer),
        width=4,
        bg=ACCENT_PRIMARY_COLOR,
        fg="black",
        font=FONTS["small_bold"],
        anchor="c",
    )
    num_label.grid(row=0, column=0, padx=(10, 10), sticky="w")

    # Statement
    stmt_label = tk.Label(
        row,
        text=stelling,
        bg=LIGHT_GREY,
        fg="black",
        font=TEXT_FONT,
        anchor="w",
        justify="left",
        wraplength=650,
    )
    stmt_label.grid(row=0, column=1, sticky="w", padx=10, pady=8)

    # Likert buttons (right)
    likert_frame = tk.Frame(row, bg=LIGHT_GREY)
    likert_frame.grid(row=0, column=2, padx=20, pady=8, sticky="e")

    buttons: list[tuple[str, tk.Label]] = []

    def update_buttons():
        current = var.get()
        for value, widget in buttons:
            if current == value:
                widget.config(relief="sunken", bg=LIKERT_SELECTED_BG, fg="white")
            else:
                widget.config(relief="raised", bg=LIKERT_DEFAULT_BG, fg="black")

    for col, (value, _label_text) in enumerate(LIKERT_OPTIONS):
        btn = tk.Label(
            likert_frame,
            text=value,
            width=3,
            bd=1,
            relief="raised",
            bg=LIKERT_DEFAULT_BG,
            fg="black",
            font=FONTS["small"],
        )
        btn.grid(row=0, column=col, padx=4)

        def on_click(event, v=value):
            var.set(v)
            update_buttons()

        btn.bind("<Button-1>", on_click)
        buttons.append((value, btn))

    update_buttons()


# -------------------- Page builder --------------------
def build_assessments_page(parent_frame: tk.Frame, navigate) -> None:
    """
    Bouwt de Big Five 1.1 vragenlijst in het rechter scherm.
    navigate is de router uit app.py, dus je kan: navigate("phase2.0") etc.
    """
    clear_frame(parent_frame)
    # Ensure client context is present
    client = getattr(parent_frame, "current_assessment_client", None)
    if client is None:
        messagebox.showerror("Fout", "Geen actieve client gevonden. Keer terug naar het dashboard en selecteer een client.")
        return

    # ---------- fixed header (non-scrolling) ----------
    header_frame = tk.Frame(parent_frame, bg="white")
    header_frame.pack(fill="x")

    # ---------- scrollable container ----------
    container = tk.Frame(parent_frame, bg="white")
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container, bg="white", highlightthickness=0)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    scroll_frame = tk.Frame(canvas, bg="white")
    window_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    def on_frame_configure(_event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_canvas_resize(event):
        canvas.itemconfig(window_id, width=event.width)

    scroll_frame.bind("<Configure>", on_frame_configure)
    canvas.bind("<Configure>", on_canvas_resize)

    # Mousewheel (Windows)
    def on_mousewheel(event):
        canvas.yview_scroll(-int(event.delta / 120), "units")

    canvas.bind_all("<MouseWheel>", on_mousewheel)

    # ---------- title & explanation (fixed) ----------
    title = tk.Label(
        header_frame,
        text="Fase 1.1 – Big Five persoonlijkheidsdimensies",
        bg="white",
        fg="black",
        font=FONTS["title"],
        anchor="w",
    )
    title.pack(fill="x", padx=20, pady=(15, 5))

    subtitle = tk.Label(
        header_frame,
        text=(
            "Beoordeel elke stelling in het algemeen voor jezelf.\n"
            "1 = oneens · 2 = deels oneens · 3 = neutraal · 4 = deels eens · 5 = volledig eens."
        ),
        bg="white",
        fg="black",
        font=FONTS["small"],
        anchor="w",
        justify="left",
    )
    subtitle.pack(fill="x", padx=20, pady=(0, 15))

    # ---------- column headings (fixed) ----------
    header = tk.Frame(header_frame, bg=DARK_GREY)
    header.pack(fill="x")

    tk.Label(
        header,
        text="Nummer",
        bg=DARK_GREY,
        fg="white",
        font=FONTS["small_bold"],
        width=8,
        anchor="w",
        padx=10,
    ).grid(row=0, column=0, sticky="w")

    tk.Label(
        header,
        text="Stelling",
        bg=DARK_GREY,
        fg="white",
        font=FONTS["small_bold"],
        anchor="w",
        padx=10,
    ).grid(row=0, column=1, sticky="w")

    tk.Label(
        header,
        text="",
        bg=DARK_GREY,
        fg="white",
        font=FONTS["small_bold"],
        anchor="e",
        padx=40,
    ).grid(row=0, column=2, sticky="e")

    # ---------- all items ----------
    parent_frame.assessment_vars = {}

    questions_container = tk.Frame(scroll_frame, bg="white")
    questions_container.pack(fill="both", expand=True)

    # Load saved session if it exists
    client = getattr(parent_frame, "current_assessment_client", None)
    saved_session = load_session(str(client["id"]), client["name"], "phase1.1")
    saved_answers = saved_session.get("answers", {}) if saved_session else {}

    # Helper function to auto-save answers
    def auto_save_answers():
        """Save current answers to session file."""
        results = get_assessment_results(parent_frame)
        # Convert keys to strings for JSON serialization
        answers_to_save = {str(k): v for k, v in results.items()}
        save_session(
            str(client["id"]),
            client["name"],
            "phase1.1",
            0,
            answers_to_save,
            len(BIG_FIVE_ITEMS)
        )

    for nummer, tekst in BIG_FIVE_ITEMS:
        var = tk.StringVar(value="")
        # Pre-populate with saved answer if it exists
        if str(nummer) in saved_answers:
            saved_value = saved_answers[str(nummer)]
            if saved_value:
                var.set(str(saved_value))
        parent_frame.assessment_vars[nummer] = var
        # Add trace callback to auto-save on any change
        var.trace_add("write", lambda *args: auto_save_answers())
        make_likert_row(questions_container, nummer, tekst, var)

    # spacer
    tk.Frame(scroll_frame, bg="white", height=10).pack(fill="x")

    # ---------- Save and continue button ----------
    def on_submit():
        results = get_assessment_results(parent_frame)

        missing = [nr for nr, v in results.items() if v is None]
        if missing:
            messagebox.showwarning(
                "Onvolledige vragenlijst",
                f"Er zijn nog {len(missing)} stellingen niet ingevuld."
            )
            return

        # Save results and generate Excel file in client folder
        parent_frame.assessment_results = results

        root = parent_frame.winfo_toplevel()
        client = getattr(parent_frame, "current_assessment_client", None)
        if not client:
            messagebox.showerror("Fout", "Geen actieve client gevonden. Keer terug naar het dashboard en selecteer een client.")
            return
        safe_name = "_".join(client["name"].split())
        folder_name = f"{client['id']}_{safe_name}"
        client_dir = os.path.join("clients", folder_name)
        os.makedirs(client_dir, exist_ok=True)
        excel_path = os.path.join(client_dir, f"{safe_name}_assessment_results.xlsx")
        save_path = write_assessment_answers_to_excel(results, excel_path=excel_path)
        if save_path:
            root.results_excel_path = save_path
            messagebox.showinfo("Succes", f"Antwoorden opgeslagen naar:\n{save_path}")
        else:
            messagebox.showwarning("Waarschuwing", "Antwoorden opgeslagen maar Excel-bestand kon niet worden gegenereerd.")

        # Mark session as complete
        mark_session_complete(str(client["id"]), client["name"], "phase1.1")

        # -> go to phase 2.0
        navigate("phase2.0")

    btn_frame = tk.Frame(scroll_frame, bg="white")
    btn_frame.pack(fill="x", pady=(5, 20))

    add_nav_buttons(
        btn_frame,
        submit_command=on_submit,
        submit_text="Opslaan en verder",
        skip_side="left",
        submit_side="right",
        padx=BUTTON_CONTAINER_PADX
    )

    # extra space at the bottom
    tk.Frame(scroll_frame, bg="white", height=30).pack(fill="x")