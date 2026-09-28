import tkinter as tk
from tkinter import messagebox
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from ui.ui_components import clear_frame, create_submit_button
from utils.write_assessments_to_excel import write_phase_2_3_to_excel
from utils.session_manager import save_session, load_session, mark_session_complete
from ui.ui_styles import S, FONTS, PRIMARY_BG, CARD_LIGHT_BG

# ==================== Job Characteristics Model Data ====================
JOB_CHARACTERISTICS_MODEL = [
    (1, "Taakvaardigheid",
        "De mate waarin een baan verschillende activiteiten vereist, waarbij de werknemer een verscheidenheid aan vaardigheden en talenten moet ontwikkelen. Werknemers kunnen meer betekenis ervaren in banen die verschillende vaardigheden en capaciteiten vereisen dan wanneer de banen elementair en routinematig zijn.",
        "Hoe belangrijk vind je het dat je werk bestaat uit verschillende soorten activiteiten waarbij je diverse vaardigheden en talenten kunt inzetten en/of ontwikkelen? En wat zou voor jou een ideale verhouding zijn tussen afwisselende taken en meer routinematige werkzaamheden?"),

    (2, "Taakidentiteit",
        "De mate waarin de functie vereist dat de functiehouders een werkstuk identificeren en voltooien met een zichtbaar resultaat. Werknemers ervaren meer zingeving in een baan wanneer ze betrokken zijn bij het hele proces in plaats van alleen verantwoordelijk te zijn voor een deel van het werk.",
        "Hoe belangrijk vind je het om in je werk betrokken te zijn bij een volledig proces of werkstuk, van begin tot eind, met een duidelijk zichtbaar resultaat? En in hoeverre zou je idealiter verantwoordelijk willen zijn voor het hele werkproces versus alleen een deel ervan?"),

    (3, "Taakbetekenis",
        "De mate waarin de baan het leven van anderen beïnvloedt. De invloed kan zowel in de directe organisatie als in de externe omgeving zijn. Werknemers ervaren meer zingeving in een baan die het psychisch of fysiek welzijn van anderen aanzienlijk verbetert, dan een baan die een beperkt effect heeft op iemand anders.",
        "Hoe belangrijk vind je het dat je werk een merkbare invloed heeft op het leven of welzijn van anderen, binnen en/of buiten de organisatie? En in welke mate zou je idealiter willen dat jouw werk impact heeft op anderen?"),

    (4, "Autonomie",
        "De mate waarin de baan de werknemer aanzienlijke vrijheid, onafhankelijkheid en discretie biedt om het werk te plannen en de procedures in de baan te bepalen. Voor banen met een hoge mate van autonomie hangen de resultaten van het werk af van de eigen inspanningen, initiatieven en beslissingen van de werknemers. In plaats van in opdracht van een manager of een handleiding met werkprocedures. In dergelijke gevallen ervaren de werknemers een grotere persoonlijke verantwoordelijkheid voor hun eigen successen en mislukkingen op het werk.",
        "Hoe belangrijk vind je het om in je werk zelf te kunnen bepalen hoe en wanneer je taken uitvoert, zonder dat alles strak voorgeschreven is? En hoeveel vrijheid zou je idealiter willen hebben in het plannen en uitvoeren van je werk?"),

    (5, "Feedback",
        "De mate waarin de werknemer kennis heeft van resultaten. Dit is duidelijke, specifieke, gedetailleerde, bruikbare informatie over de effectiviteit van zijn of haar werkprestaties. Wanneer werknemers duidelijke, bruikbare informatie over hun werkprestaties ontvangen, hebben ze een betere algemene kennis van het effect van hun werkactiviteiten en welke specifieke acties ze moeten ondernemen (indien aanwezig) om hun productiviteit te verbeteren.",
        "Hoe belangrijk vind je het om duidelijke en bruikbare feedback te krijgen over hoe goed je je werk doet? En in welke mate zou je idealiter op de hoogte willen zijn van het effect van je werk en van punten waarop je jezelf kunt verbeteren?"),
]

# ==================== Build Function ====================
def build_job_characteristics_models_page(parent_frame: tk.Frame, navigate=None) -> None:
    """Build the Job Characteristics Models page with text input fields."""
    
    clear_frame(parent_frame)
    
    # Get client context and load session
    client = getattr(parent_frame, "current_assessment_client", None)
    if client:
        saved_session = load_session(str(client["id"]), client["name"], "phase2.3")
        saved_answers = saved_session.get("answers", {}) if saved_session else {}
    else:
        saved_answers = {}
    
    # =================== Scrollable container ===================
    container = tk.Frame(parent_frame, bg=S["bg"])
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container, bg=S["bg"], highlightthickness=0)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    scroll_frame = tk.Frame(canvas, bg=S["bg"])
    window_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    def on_frame_configure(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_canvas_resize(event):
        canvas.itemconfig(window_id, width=event.width)

    scroll_frame.bind("<Configure>", on_frame_configure)
    canvas.bind("<Configure>", on_canvas_resize)

    def on_mousewheel(event):
        canvas.yview_scroll(-int(event.delta / 120), "units")

    canvas.bind_all("<MouseWheel>", on_mousewheel)

    # =================== Headers ===================
    title = tk.Label(
        scroll_frame,
        text="Fase 2.3 – Werk karakteristieken modellen",
        bg=S["bg"],
        fg="black",
        font=FONTS["medium_bold"],
        anchor="w",
    )
    title.pack(fill="x", padx=20, pady=(15, 5))

    subtitle = tk.Label(
        scroll_frame,
        text="Geef je eigen antwoord op elke vraag. Typ je gedachten, gevoelens en ervaringen in het tekstvak.",
        bg=S["bg"],
        fg="#555555",
        font=S["f_sub"],
        anchor="w",
        justify="left",
        wraplength=700,
    )
    subtitle.pack(fill="x", padx=20, pady=(0, 20))

    # =================== Dictionary to store text entries ===================
    text_entries = {}

    # =================== Build question cards ===================
    for idx, (q_num, characteristic, description, question) in enumerate(JOB_CHARACTERISTICS_MODEL):
        # Outer container for header (above) and card (below)
        outer = tk.Frame(scroll_frame, bg=S["bg"])
        outer.pack(fill="x", padx=15, pady=8)

        # Header above the box: yellow background with an outline
        header_frame = tk.Frame(outer, bg=S["yellow"], relief="solid", bd=1)
        header_frame.pack(fill="x", pady=(0, 4))

        header = tk.Label(
            header_frame,
            text=f"{q_num}. {characteristic}",
            bg=S["yellow"],
            fg="white",
            font=S["f_b"],
            anchor="w",
            wraplength=700,
            justify="left",
        )
        header.pack(fill="x", padx=12, pady=6)

        # Card background (content box)
        card = tk.Frame(outer, bg=S["odd"] if idx % 2 == 0 else S["even"], relief="solid", bd=1)
        card.pack(fill="x")

        # Description
        desc_label = tk.Label(
            card,
            text=description,
            bg=card.cget("bg"),
            fg="#000000",
            font=FONTS["medium"],
            anchor="w",
            wraplength=700,
            justify="left",
        )
        desc_label.pack(fill="x", padx=12, pady=(0, 10))

        # Question
        question_label = tk.Label(
            card,
            text=question,
            bg=card.cget("bg"),
            fg="black",
            font=FONTS["medium"],
            anchor="w",
            wraplength=700,
            justify="left",
        )
        question_label.pack(fill="x", padx=12, pady=(0, 8))

        # Text input box (multi-line)
        text_box = tk.Text(
            card,
            height=4,
            width=80,
            font=FONTS["medium"],
            bg="white",
            fg="black",
            relief="solid",
            bd=1,
            wrap="word",
            padx=8,
            pady=6,
        )
        text_box.pack(fill="x", padx=12, pady=(0, 12))
        
        # Restore saved answer if it exists
        if str(q_num) in saved_answers:
            saved_value = saved_answers[str(q_num)]
            if saved_value:
                text_box.insert("1.0", saved_value)
        
        # Add callback for auto-save
        def make_auto_save(q_num):
            def auto_save_on_change(*args):
                if client:
                    # Collect all current answers
                    answers_to_save = {}
                    for qn, tb in text_entries.items():
                        answer = tb.get("1.0", "end-1c").strip()
                        if answer:
                            answers_to_save[str(qn)] = answer
                    
                    # Save to session
                    save_session(
                        str(client["id"]),
                        client["name"],
                        "phase2.3",
                        0,
                        answers_to_save,
                        len(JOB_CHARACTERISTICS_MODEL)
                    )
            return auto_save_on_change
        
        text_box.bind("<KeyRelease>", make_auto_save(q_num))
        
        # Store reference to text box
        text_entries[q_num] = text_box

    # =================== Submit Button ===================
    button_frame = tk.Frame(scroll_frame, bg=S["bg"])
    button_frame.pack(fill="x", padx=20, pady=20)

    def on_submit():
        """Validate and submit answers."""
        missing = []
        for q_num, text_box in text_entries.items():
            answer = text_box.get("1.0", "end-1c").strip()
            if not answer:
                missing.append(q_num)

        if missing:
            messagebox.showwarning(
                "Onvolledige vragenlijst",
                f"Vraag(en) {missing} nog niet ingevuld. Vul alle vragen in alvorens verder te gaan."
            )
            return

        # save answers to dictionary
        answers = {q_num: text_box.get("1.0", "end-1c").strip() for q_num, text_box in text_entries.items()}
        
        # save to Excel file
        root = parent_frame.winfo_toplevel()
        excel_path = getattr(root, "results_excel_path", None)

        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror(
                "Geen Excel-bestand",
                "Er is nog geen resultatenbestand aangemaakt (fase 1.1)."
            )
            return

        result = write_phase_2_3_to_excel(
            answers,
            excel_path
        )

        if not result:
            messagebox.showerror(
                "Fout bij opslaan",
                "Er ging iets mis bij het opslaan van je antwoorden."
            )
            return
        
        messagebox.showinfo(
            "Succes",
            f"Je antwoorden zijn opgeslagen"
        )

        # Mark session as complete
        if client:
            mark_session_complete(str(client["id"]), client["name"], "phase2.3")

        # Navigate directly to client dashboard after completing all phases
        if navigate:
            navigate("client_dashboard")

    btn_submit = create_submit_button(
        button_frame,
        text="Opslaan en verder",
        command=on_submit
    )
    btn_submit.pack(side="right")