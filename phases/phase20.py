import tkinter as tk
from tkinter import messagebox
from utils import write_assessments_to_excel
from utils.session_manager import save_session, load_session, mark_session_complete
from ui.ui_components import clear_frame, create_back_button, create_submit_button
from ui.ui_styles import (
    S, FONTS, PRIMARY_BG, SECONDARY_BG, CARD_LIGHT_BG,
    TABLE_PADX, TABLE_PADY, BUTTON_PADX, BUTTON_PADY, BUTTON_CONTAINER_PADX
)

# ----------------------------- Phase 2.0 - Career anchors -----------------------------
# Each entry: (question number, anchor letter, statement text)
# Anchors: V = Moving upward, W = Feeling safe, X = Being free, Y = Finding balance, Z = Seeking challenge.

CAREER_STATEMENTS = [
    (1, "V", "V | Graag wil ik het voor mezelf en voor anderen dusdanig regelen dat succes verzekerd is."),
    (1, "X", "X | Ik houd me binnen een werksituatie het liefst bezig met mijn eigen zaken."),
    (2, "Y", "Y | Binnen het werk moet er tijd zijn voor zaken die jezelf belangrijk vindt en moet er gelegenheid zijn om zinvolle relaties te cultiveren."),
    (2, "V", "V | Vooruitkomen is voor mij belangrijker dan persoonlijke behoeften."),
    (3, "W", "W | Ik werk graag in een omgeving waar hard werken, loyaliteit en toewijding gewaardeerd wordt."),
    (3, "X", "X | Ik houd van een werksituatie waar ik mijn eigen doelen kan stellen en ze kan bereiken op mijn eigen manier en op mijn eigen tempo."),
    (4, "V", "V | Ik ben strijdlustig, kan goed analyseren en met mensen omgaan."),
    (4, "Y", "Y | Ik kan goed mijn evenwicht bewaren tussen de eisen van mijn werk en die van mijn privé-leven."),
    (5, "X", "X | Ik wil onafhankelijk werken."),
    (5, "W", "W | Ik houd ervan me een vertegenwoordiger te voelen van een groter geheel."),
    (6, "Z", "Z | Ik houd ervan als consultant of probleemoplosser te werken en me dusdanig te profileren door middel van een opwindend project."),
    (6, "V", "V | Ik houd ervan in een situatie te werken waarin ik de leiding heb en verantwoordelijk ben voor het bereiken van bepaalde doelen."),
    (7, "Y", "Y | Mijn echtgenoot/partner is net zo belangrijk voor mij als mijn loopbaan."),
    (7, "Z", "Z | Mijn echtgenoot/partner verdwijnt naar de achtergrond als ik midden in een zeer opwindend project zit."),
    (8, "X", "X | Het allerbelangrijkst voor mij is vrijheid."),
    (8, "Y", "Y | Het allerbelangrijkst voor mij is een doel in mijn leven."),
    (9, "W", "W | Ik ben bekwaam, loyaal, betrouwbaar en ik werk hard."),
    (9, "Z", "Z | Ik ben sociaal en in de omgang, een goede leider en een goede organisator."),
    (10, "X", "X | Ik ben onafhankelijk."),
    (10, "Y", "Y | Ik ben evenwichtig."),
    (11, "Z", "Z | Ik ben iemand die in actie komt door opwindende projecten."),
    (11, "Y", "Y | Ik ben iemand die graag met anderen werkt."),
    (12, "X", "X | Ik ben ambitieus en iemand die graag met anderen wedijvert."),
    (12, "W", "W | Ik ben iemand die een medewerker zijn met wie men kan rekenen."),
    (13, "Z", "Z | Ik voel zelfvertrouwen en ben in staat mezelf te redden."),
    (13, "V", "V | Ik heb veel fantasie en enthousiasme."),
    (14, "W", "W | Ik ben stabiel en vasthoudend."),
    (14, "X", "X | Ik ben onafhankelijk en in staat een eigen koers te bepalen."),
    (15, "Y", "Y | Ik ben iemand die goed kan plannen en coördineren."),
    (15, "Z", "Z | Ik ben iemand die situaties analyseert en creatieve, nieuwe oplossingen ontwikkelt."),
    (16, "V", "V | Ik ben een expert op mijn terrein."),
    (16, "W", "W | Ik ben een betrouwbare en degelijk persoon."),
    (17, "Y", "Y | Ik ben iemand die wil werken volgens vaststaande procedures."),
    (17, "X", "X | Ik ben iemand die probeert de doelen in het werk in overeenstemming te brengen met het persoonlijk nastreven."),
    (18, "Z", "Z | Een persoonlijk doel is om mijn eigen lot te bepalen."),
    (18, "Y", "Y | Een persoonlijk doel is om mijn werk te verweven met mijn privé-leven."),
    (19, "W", "W | Ik vind het belangrijk een veilige baan te hebben en het gevoel te hebben erbij te horen."),
    (19, "X", "X | Ik vind het belangrijk om tijd te kunnen besteden aan mijn privé-leven en hobby’s."),
    (20, "V", "V | Ik geef de voorkeur aan een carrière waarin veel promotiekansen voorhanden zijn."),
    (20, "Z", "Z | Ik geef de voorkeur aan om in staat gesteld te worden uitdagende problemen en taken aan te pakken."),
    (21, "Y", "Y | Ik ben graag in een werksituatie waar invloed uitgeoefend kan worden."),
    (21, "W", "W | Ik waardeer een baan waar je langere tijd kunt blijven werken en waar je gewaardeerd en geaccepteerd wordt."),
    (22, "V", "V | Ik denk dat de juiste mensen en goede vrienden maken belangrijk is om vooruit te komen."),
    (22, "Z", "Z | Ik denk dat het essentieel is om interessesgebieden te ontwikkelen."),
    (23, "Y", "Y | Voor mij geldt als basis het scheppen van een evenwicht tussen mijn privé-leven en mijn werk."),
    (23, "W", "W | Voor mij geldt als basis stabiliteit, waardering en een veilige plaats binnen mijn werksituatie."),
    (24, "X", "X | Ik denk dat ik graag een positie zou willen hebben met een maximum aan zelfstandigheid."),
    (24, "V", "V | Ik denk dat ik graag tot \"de kring van ingewijden\" zou willen behoren."),
    (25, "W", "W | Voor mij geldt als basis stabiliteit, waardering en een veilige plaats op het werk."),
    (25, "V", "V | Als basis geldt voor mij dat ik vooruit wil komen in de werkomgeving."),
    (26, "V", "V | Ik denk dat geld, macht en aanzien een belangrijke maatstaf zijn van een succesvolle loopbaan."),
    (26, "Y", "Y | Ik denk dat een loopbaan succesvol is als je evenveel tijd hebt voor het werk, het gezin en je eigen ontwikkeling."),
    (27, "Z", "Z | Ik wil liever uitblinken op mijn gebied."),
    (27, "W", "W | Ik wil liever beschouwd worden als betrouwbaar en loyaal."),
    (28, "W", "W | Ik geef de voorkeur aan het werken met een team op lange termijn en een hechte basis."),
    (28, "Z", "Z | Ik geef de voorkeur aan het werken met een taakgerichte of projectgroep op korte termijn basis en in een hoog tempo."),
    (29, "Z", "Z | Ik geef de voorkeur aan professionele ontwikkeling en permanente training."),
    (29, "X", "X | Ik geef de voorkeur aan professionele ontwikkeling om een expert te worden en om meer flexibiliteit en onafhankelijkheid te verkrijgen."),
    (30, "Y", "Y | Ik geef de voorkeur aan een werksituatie die een evenwicht garandeert tussen mijn privé-leven en mijn werk."),
    (30, "Z", "Z | Ik geef de voorkeur aan een werksituatie die opwindend is en mij stimuleert."),
]


CAREER_ANCHOR_DESCRIPTIONS = {
    "Omhoog komen": "Deze op opwaartse mobiliteit gerichte loopbaanoriëntatie wordt gewoonlijk geassocieerd met het vooruitkomen in een hiërarchische en/of statusgevoelige organisatie. Het verwerven van steeds meer invloed speelt in deze kaders een grote rol. Prestige en beloning nemen bij iedere opwaartse beweging toe.",
    "Veilig voelen": "Sommige personen hebben behoefte aan een veilige baan in een duidelijke organisatie die vooral gekenmerkt wordt door orde en rust. Zij geven de voorkeur aan een lang en vast dienstverband, erkenning en appreciatie door de werkgever. In ruil daarvoor bieden ze een loyale en toegewijde instelling en zijn ze niet bang om hard te werken. Onderling respect, wederkerigheid en loyaliteit karakteriseren de werkhouding.",
    "Vrij zijn": "Personen met deze loopbaanoriëntatie zijn er op uit hun grenzen te verkennen. De nadruk ligt bij hen meer op het verwerven van persoonlijke autonomie, ruimte en verantwoordelijkheid voor het bereiken van resultaten dan op gebondenheid, zekerheid en vaste regels. Men is bereid zeer hard te werken als daar gunstige voorwaarden tegenover staan in de sfeer van onafhankelijkheid en zelfcontrole. Interessant werk is belangrijk maar individuele vrijheid is het uiteindelijke doel.",
    "Balans vinden": "De meeste mensen streven naar evenwicht maar zelden vormt dit het basis uitgangspunt voor hun loopbaanbeslissingen. Sommige mensen zoeken echter een optimaal evenwicht tussen werk, privé-leven en zelfontwikkeling. Het werk vormt voor hen slechts één dimensie van hun totale levensvervulling. De aandacht voor werk en privé-leven kan verschillen afhankelijk van de situatie waarin personen met deze loopbaanoriëntatie zich bevinden. Individuen die de kwaliteit van het leven hoog in het vaandel dragen, vallen dikwijls in deze categorie.",
    "Uitdaging zoeken": "Deze loopbaanoriëntatie wordt gekenmerkt door de behoefte aan opwinding en uitdaging en een sterke betrokkenheid bij het werk. Men is er op gericht dichtbij het centrum van actie, avontuur en creativiteit te zijn en heeft er zeer veel moeite mee zich aan het werk los te maken. Een bureaucratische organisatie wordt als bijzonder remmend ervaren. Autonomie is belangrijk maar het belangrijkste is opwindend en uitdagend werk.",
}

# --------------- career anchors -> Excel column mapping ---------------
ANCHOR_TO_COLUMN = {
    "V": "C",
    "W": "D",
    "X": "E",
    "Y": "F",
    "Z": "G",
}

ANCHORS = ["V", "W", "X", "Y", "Z"]

def build_career_anchors_page(parent_frame: tk.Frame, navigate_to) -> None:
    """Show Phase 2.0 – Career Anchors in the right panel with two statements per question."""

    clear_frame(parent_frame)

    # --------------- Fixed header area ---------------

    title_frame = tk.Frame(parent_frame, bg="white")
    title_frame.pack(fill="x")

    title = tk.Label(
        title_frame,
        text="Fase 2.0 | Wat wil de cliënt? | Identificatie van de loopbaanwaarden",
        bg=S["bg"],
        fg="black",
        font=S["f_title"],
        anchor="w",
    )
    title.pack(fill="x", padx=20, pady=(20, 2))

    subtitle = tk.Label(
        title_frame,
        text=(
            "Beoordeling van stelling a.h.v. onderstaande criteria:\n"
            "Met welke stelling kan cliënt zich het sterkst identificeren?"
        ),
        bg=S["bg"],
        fg="black",
        font=S["f_sub"],
        anchor="w",
        justify="left",
    )
    subtitle.pack(fill="x", padx=20, pady=(0, 15))

    header_table = tk.Frame(parent_frame, bg=S["bg"])
    header_table.pack(fill="x", padx=TABLE_PADX)

    # columns
    for c in range(7):
        header_table.grid_columnconfigure(c, weight=0)
    header_table.grid_columnconfigure(1, weight=1)  # statement column stretches

    header_bg = S["header"]
    headers = [
        ("Nummer", 0),
        ("Stelling", 1),
        ("V", 2),
        ("W", 3),
        ("X", 4),
        ("Y", 5),
        ("Z", 6),
    ]

    for text, col in headers:
        lbl = tk.Label(
            header_table,
            text=text,
            bg=header_bg,
            fg="white",
            font=S["f_b"],
            padx=21, # space between text at teh navigation 
            anchor="w" if col <= 1 else "center",
        )
        lbl.grid(row=0, column=col, sticky="nsew")

    # --------------- Scrollable container ---------------

    container = tk.Frame(parent_frame, bg="white")
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container, bg="white", highlightthickness=0)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    scroll_frame = tk.Frame(canvas, bg="white")
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

    # =================== Table rows ===================
    table = tk.Frame(scroll_frame, bg=S["bg"])
    table.pack(fill="both", expand=True, padx=TABLE_PADX, pady=TABLE_PADY)

    # columns
    for c in range(7):
        table.grid_columnconfigure(c, weight=0)
    table.grid_columnconfigure(1, weight=1)  # statement column stretches

    # -------------------- Group statements per question --------------------
    from collections import defaultdict

    questions = defaultdict(list)
    for nummer, anker, tekst in CAREER_STATEMENTS:
        questions[nummer].append((anker, tekst))


    # -------------------- Prepare variables --------------------
    vraag_vars: dict[int, tk.StringVar] = {}
    vraag_buttons: dict[int, list[tuple[str, tk.Radiobutton]]] = {}

    # Load saved session if it exists
    client = getattr(parent_frame, "current_assessment_client", None)
    saved_session = load_session(str(client["id"]), client["name"], "phase2.0")
    saved_answers = saved_session.get("answers", {}) if saved_session else {}

    # Helper function to auto-save answers
    def auto_save_loopbaan():
        """Save current answers to session file."""
        answers_to_save = {str(k): v.get() for k, v in vraag_vars.items()}
        save_session(
            str(client["id"]),
            client["name"],
            "phase2.0",
            0,
            answers_to_save,
            len(questions)
        )

    for nummer in questions.keys():
        var = tk.StringVar(value="")
        # Pre-populate with saved answer if it exists
        if str(nummer) in saved_answers:
            saved_value = saved_answers[str(nummer)]
            if saved_value:
                var.set(saved_value)
        vraag_vars[nummer] = var
        # Add trace callback to auto-save on any change
        var.trace_add("write", lambda *args: auto_save_loopbaan())
        vraag_buttons[nummer] = []

    def update_row(question_number: int) -> None:
        """Show only a '1' in the selected box for this question."""
        current = vraag_vars[question_number].get()
        for code, btn in vraag_buttons[question_number]:
            if code == current:
                btn.config(text="1", font=FONTS["medium_bold"])
            else:
                btn.config(text="", font=FONTS["medium"])

    # -------------------- Build rows --------------------
    row_index = 0
    for nummer, stmts in sorted(questions.items()):
        row_bg = S["odd"] if row_index % 2 == 0 else S["even"]

        # Number (yellow block)
        num_label = tk.Label(
            table,
            text=str(nummer),
            width=4,
            bg=S["yellow"],
            fg="black",
            font=S["f_b"],
            anchor="c",
        )
        num_label.grid(row=row_index, column=0, rowspan=len(stmts), padx=(10, 10), pady=(2, 2), sticky="nsw")

        # Statements stacked vertically
        stmt_text = "\n".join([tekst for _, tekst in stmts])
        stmt_label = tk.Label(
            table,
            text=stmt_text,
            bg=row_bg,
            fg="black",
            font=S["f"],
            anchor="w",
            justify="left",
            padx=12,
            pady=6,
        )
        stmt_label.grid(row=row_index, column=1, rowspan=len(stmts), sticky="nsew", padx=(0, 10), pady=(2, 2))

        def update_wrap(event, label=stmt_label):
            label.configure(wraplength=event.width - 24)

        stmt_label.bind("<Configure>", update_wrap)

        # 5 anchor boxes (Radiobuttons)
        var = vraag_vars[nummer]
        for offset, code in enumerate(ANCHORS):
            col = 2 + offset
            cell = tk.Frame(table, bg=row_bg, bd=1, relief="solid")
            cell.grid(row=row_index, column=col, rowspan=len(stmts), padx=3, pady=(2, 2), sticky="nsew")

            rb = tk.Radiobutton(
                cell,
                variable=var,
                value=code,
                indicatoron=False,
                text="",
                width=4,
                font=FONTS["medium"],
                bg=row_bg,
                fg="black",
                activebackground=row_bg,
                activeforeground="black",
                selectcolor=row_bg,
                relief="flat",
                borderwidth=0,
                command=lambda q=nummer: update_row(q),
                cursor="hand2",
            )
            rb.pack(expand=True, fill="both")
            vraag_buttons[nummer].append((code, rb))

        row_index += len(stmts)

    # init: all empty
    for qnum in vraag_vars.keys():
        update_row(qnum)

    # save choices on parent_frame
    parent_frame.loopbaan_vars = vraag_vars

    # -------------------- Descriptions --------------------
    desc_frame = tk.Frame(scroll_frame, bg=S["bg"])
    desc_frame.pack(fill="x", padx=20, pady=(20, 20))

    tk.Label(
        desc_frame,
        text="Loopbaanankers – omschrijving",
        bg=S["bg"],
        fg="black",
        font=FONTS["medium_bold"],
        anchor="w",
    ).pack(fill="x", pady=(0, 5))

    for naam, tekst in CAREER_ANCHOR_DESCRIPTIONS.items():
        box = tk.Frame(desc_frame, bg=CARD_LIGHT_BG, bd=1, relief="solid")
        box.pack(fill="x", pady=4)

        tk.Label(
            box,
            text=naam,
            bg=S["yellow"],
            fg="black",
            font=S["f_b"],
            width=16,
            anchor="center",
            padx=4,
            pady=4,
        ).pack(side="left", fill="y")

        tk.Label(
            box,
            text=tekst,
            bg=CARD_LIGHT_BG,
            fg="black",
            font=S["f"],
            justify="left",
            wraplength=650,
            anchor="w",
            padx=8,
            pady=6,
        ).pack(side="left", fill="both", expand=True)

    # -------------------- Save and continue --------------------
    def on_submit_loopbaan():
        vraag_vars = parent_frame.loopbaan_vars

        # validate answers
        missing = [q for q, var in vraag_vars.items() if not var.get()]
        if missing:
            messagebox.showwarning(
                "Onvolledige vragenlijst",
                f"Er zijn nog {len(missing)} vragen niet ingevuld."
            )
            return

        # map to Excel columns
        excel_answers = {
            qnum: ANCHOR_TO_COLUMN[var.get()]
            for qnum, var in vraag_vars.items()
        }

        # write to Excel
        root = parent_frame.winfo_toplevel()

        if not hasattr(root, "results_excel_path"):
            messagebox.showerror(
                "Geen Excel-bestand",
                "Er is nog geen resultatenbestand aangemaakt (fase 1.1)."
            )
            return


        excel_path = getattr(root, "results_excel_path", None)

        save_path = write_assessments_to_excel.write_career_anchors_to_excel(
            excel_answers,
            excel_path
        )

        if not save_path:
            messagebox.showerror(
                "Fout",
                "Kon de antwoorden niet opslaan in Excel."
            )
            return

        # same excel file
        root.results_excel_path = save_path

        messagebox.showinfo(
            "Opgeslagen",
            f"Antwoorden opgeslagen in:\n{save_path}"
        )

        # Mark session as complete
        client = getattr(parent_frame, "current_assessment_client", None)
        mark_session_complete(str(client["id"]), client["name"], "phase2.0")

        # next phase
        navigate_to("phase2.1")

    btn_frame = tk.Frame(scroll_frame, bg=S["bg"])
    btn_frame.pack(fill="x", pady=(10, 30))

    from ui.ui_components import add_nav_buttons
    add_nav_buttons(
        btn_frame,
        submit_command=on_submit_loopbaan,
        submit_text="Opslaan en verder",
        skip_side="left",
        submit_side="right",
        padx=BUTTON_CONTAINER_PADX
        )

def show(parent_frame: tk.Frame):
    """Public function to display Phase 2.0 page."""
    build_career_anchors_page(parent_frame)