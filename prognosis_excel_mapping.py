"""
prognosis_excel_mapping.py
--------------------------
Stappen:
1. Copy template on submission
2. Map out answers to Excel
3. Write answers from questionnaire into Excel
4. Testing
"""

import json
import shutil
import os
from openpyxl import load_workbook


# ─────────────────────────────────────────────
# INSTELLINGEN 
# ─────────────────────────────────────────────

TEMPLATE_PATH = "Integratie_Prognose_Model_5.0_(2).xlsx"  # het originele template
OUTPUT_FOLDER = "clients"                           # map waar klant-excels komen

# ─────────────────────────────────────────────
# MAPPING
# (indeling gemaakt door Mette - niet aanpassen)
# ─────────────────────────────────────────────

PROGNOSIS_MAPPING = {

    "q1": {
        "depends_on": "gender",
        "man":   {"row": 3,   "options": 3},
        "woman": {"row": 20,  "options": 3},
    },

    "q3":        {"row": 37,  "options": 3},
    "q4":        {"row": 39,  "options": 12},
    "q5":        {"row": 41,  "options": 3},
    "q6":        {"row": 45,  "options": 5},
    "q7":        {"row": 48,  "options": 5},
    "q8":        {"row": 52,  "options": 4},
    "q9":        {"row": 68,  "options": 5},
    "q10":       {"row": 70,  "options": 3},
    "q11":       {"row": 74,  "options": 3},
    "q12":       {"row": 76,  "options": 5},
    "q13":       {"row": 78,  "options": 5},
    "q14":       {"row": 80,  "options": 3},
    "q15":       {"row": 82,  "options": 2},
    "q16":       {"row": 84,  "options": 2},
    "q17":       {"row": 87,  "options": 4},
    "q18":       {"row": 90,  "options": 3},
    "q19":       {"row": 93,  "options": 3},
    "q20":       {"row": 95,  "options": 3},
    "q21":       {"row": 97,  "options": 4},
    "q22":       {"row": 100, "options": 3},
    "q23":       {"row": 103, "options": 3},
    "q24":       {"row": 106, "options": 3},

    "q25":       {"row": 109, "options": 3},

    "q26":       {"row": 113, "options": 4},
    "q27":       {"row": 116, "options": 4},
    "q28":       {"row": 119, "options": 5},
    "q29":       {"row": 122, "options": 3},
    "q30":       {"row": 125, "options": 5},
    "q31":       {"row": 127, "options": 3},
    "q32":       {"row": 129, "options": 3},
    "q33":       {"row": 131, "options": 3},
    "q34":       {"row": 133, "options": 3},
    "q35":       {"row": 135, "options": 3},
    "q36":       {"row": 137, "options": 3},
    "q37":       {"row": 140, "options": 4},
    "q38":       {"row": 143, "options": 4},
    "q39":       {"row": 146, "options": 2},
    "q40":       {"row": 148, "options": 2},
    "q41":       {"row": 150, "options": 2},
    "q42":       {"row": 152, "options": 4},
    "q43":       {"row": 154, "options": 2},
    "q44":       {"row": 156, "options": 4},
    "q45":       {"row": 158, "options": 3},
    "q46":       {"row": 160, "options": 4},
}

# ─────────────────────────────────────────────
# STAP 1: COPY TEMPLATE
# ─────────────────────────────────────────────

def copy_template(output_path: str):
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    shutil.copy2(TEMPLATE_PATH, output_path)

# ─────────────────────────────────────────────
# STAP 2 & 3: MAP ANSWERS + WRITE TO EXCEL
# ─────────────────────────────────────────────


def fill_excel(file_path: str, answers: dict):
    """
    answers = {
        "q1":  0,    <- 0 = Man, 1 = Vrouw
        "q3":  1,    <- index van gekozen optie (0 = eerste keuze)
        "q4":  4,
        ...
    }
    """
    wb = load_workbook(file_path)
    ws = wb["Integratie Prognose Model"]  # Corrected sheet name

    gender = "man" if answers.get("q1", 0) == 0 else "woman"

    for q, config in PROGNOSIS_MAPPING.items():
        if q not in answers:
            continue

        if "depends_on" in config:
            config = config[gender]

        row = config["row"]

        if q == "q1":
            age_index = answers.get("q2")
            if age_index is None:
                option_index = answers[q]
            else:
                option_index = age_index
        else:
            option_index = answers[q]

        col = chr(ord("D") + option_index)   # index 0 -> D, 1 -> E, 2 -> F ...

        value = ws[f"{col}{row}"].value
        ws[f"Q{row}"] = value

    wb.save(file_path)


# ─────────────────────────────────────────────
# GECOMBINEERDE FUNCTIE - gebruik deze vanuit de app
# ─────────────────────────────────────────────

def process_submission(client_id: str, answers: dict) -> str:
    """
    Dit is de enige functie die de rest van de app hoeft aan te roepen.

    Wat het doet:
        1. Maakt een kopie van het template voor deze klant
        2. Schrijft de antwoorden naar die kopie
        3. Slaat de antwoorden op als JSON voor archivering
        4. Geeft het bestandspad terug zodat de site het kan opslaan/tonen

    Parameters:
        client_id  -- unieke ID van de klant vanuit het klantaccount
                      bijv. "klant_042" of "jan_de_vries"
        answers    -- dict met vraagnummer -> gekozen index
                      { "q1": 0, "q3": 1, "q4": 4, ... }

    Gebruik vanuit de rest van de app:

        from prognosis_excel_mapping import process_submission

        bestand = process_submission(
            client_id = client.id,
            answers   = form.answers
        )
        # -> "clients/klant_042/prognose.xlsx"

    Geeft terug:
        Pad naar het opgeslagen Excel-bestand (str)
    """
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template niet gevonden: {TEMPLATE_PATH}. Zorg dat het bestand in dezelfde map staat als het script.")
    
    client_folder = os.path.join(OUTPUT_FOLDER, client_id)
    os.makedirs(client_folder, exist_ok=True)
    safe_client_id = client_id.replace(" ", "_")
    output_path = os.path.join(client_folder, f"prognose_{safe_client_id}.xlsx")
    answers_path = os.path.join(client_folder, "answers.json")
    
    copy_template(output_path)
    fill_excel(output_path, answers)
    
    # Save answers as JSON for record-keeping
    with open(answers_path, 'w', encoding='utf-8') as f:
        json.dump(answers, f, indent=4, ensure_ascii=False)
    
    return output_path
# ─────────────────────────────────────────────
# STAP 4: TESTING
# ─────────────────────────────────────────────

def test_mapping():
    """
    Simuleert een volledig ingevuld formulier en controleert
    of alle Q-cellen correct worden ingevuld.

    Draai met:  python prognosis_excel_mapping.py
    """
    test_answers = {
        "q1":      0,   # Man
        "q3":      1,   # Westers
        "q4":      1,   # Utrecht
        "q5":      0,   # Autochtoon
        "q6":      2,   # Extraversie
        "q7":      3,   # Zakelijk Formeel
        "q8":      0,   # Geen problematiek
        "q9":      4,   # Geen beperkingen
        "q10":     2,   # Actief
        "q11":     0,   # WW
        "q12":     1,   # 3-6 maanden WW
        "q13":     2,   # Geen schulden
        "q14":     2,   # Geen overige inkomsten
        "q15":     0,   # Financiele steun beschikbaar
        "q16":     1,   # Normaal IQ
        "q17":     1,   # B1-B2 Nederlands
        "q18":     1,   # 1 vreemde taal
        "q19":     1,   # B1-B2 vreemde taal
        "q20":     1,   # Basis ICT
        "q21":     1,   # MBO
        "q22":     2,   # Goede MBO-perspectieven
        "q25":     0,   # Alleenstaand
        "q26":     2,   # Midden inkomen
        "q27":     1,   # MBO onderwijsniveau
        "q29":     1,   # Gemiddeld netwerk
        "q30":     2,   # Geen justitieproblemen
        "q31":     2,   # Geen thuisproblemen
        "q32":     2,   # Geen verslaving
        "q33":     2,   # Rijbewijs ja
        "q34":     2,   # Onbeperkte mobiliteit
        "q35":     0,   # Betaald -> Betaald
        "q37":     0,   # Homogene transitie
        "q38":     0,   # Volledig werkloos
        "q39":     0,   # Vrijwilligerswerk
        "q40":     1,   # 3-6 maanden werkloos
        "q41":     0,   # Structureel werk
        "q42":     1,   # 6-10 jaar ervaring
        "q43":     1,   # 5-10 sollicitaties/mnd
        "q44":     0,   # Jan-Mrt
        "q45":     0,   # Hoeveel sollicitaties per maand
        "q46":     0,   # Jan-Mrt
    }

    if not os.path.exists(TEMPLATE_PATH):
        print(f"[FOUT] Template niet gevonden: {TEMPLATE_PATH}")
        print(f"       Zorg dat '{TEMPLATE_PATH}' in dezelfde map staat.")
        return

    output = process_submission("test_klant", test_answers)

    wb = load_workbook(output, data_only=True)
    ws = wb["Integratie Prognose Model"]

    verwacht = {
        3: "q1",  37: "q3",  39: "q4",  41: "q5",  45: "q6",
        48: "q7", 52: "q8",  68: "q9",  70: "q10", 74: "q11",
        76: "q12", 78: "q13", 80: "q14", 82: "q15",
        84: "q16", 87: "q17", 90: "q18", 93: "q19", 95: "q20",
        97: "q21", 100: "q22", 103: "q23", 106: "q24", 109: "q25",
        113: "q26", 116: "q27", 119: "q28", 122: "q29", 125: "q30",
        127: "q31", 129: "q32", 131: "q33", 133: "q34", 135: "q35",
        137: "q36", 140: "q37", 143: "q38", 146: "q39", 148: "q40",
        150: "q41", 152: "q42", 154: "q43", 156: "q44", 158: "q45",
        160: "q46",
    }

    conditioneel = ["q23", "q24", "q28", "q36"]

    print("\n=== TEST RESULTATEN ===")
    fouten = 0
    for row, vraag in verwacht.items():
        val = ws[f"Q{row}"].value
        ok  = val is not None
        if not ok:
            fouten += 1
        print(f"  Q{row:>3} ({vraag:<8}): {'OK' if ok else 'LEEG'}  ->  {val}")

    print(f"\n  Overgeslagen (conditioneel): {', '.join(conditioneel)}")
    print(f"\n  {'GESLAAGD' if fouten == 0 else f'MISLUKT  --  {fouten} onverwacht lege cellen'}")
    print(f"  Opgeslagen als: {output}\n")


# ─────────────────────────────────────────────
# START
# ─────────────────────────────────────────────

if __name__ == "__main__":
    test_mapping()