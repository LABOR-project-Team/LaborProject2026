from openpyxl import load_workbook
from datetime import date, datetime
from pathlib import Path
from docx import Document


REQUIRED_SHEETS = {
    "Fase 1.1 | Big Five Dimensies",
    "Fase 2.0 | Loopbaanankers",
    "Fase 2.1 | Carriere Clusters",
    "Fase 2.2 | Cultuur analyse",
    "Fase 2.3 | J.C.M.",
    "Gegevens",
    "Rapport",
}

BIG_FIVE_SCORE_FORMULAS = {
    "Extraversie": (20, (("C4", 1), ("C9", -1), ("C14", 1), ("C19", -1), ("C24", 1), ("C29", -1), ("C34", 1), ("C39", -1), ("C44", 1), ("C49", -1))),
    "Altruïsme": (14, (("D5", -1), ("D10", 1), ("D15", -1), ("D20", 1), ("D25", -1), ("D30", 1), ("D35", -1), ("D40", 1), ("D45", 1), ("D50", 1))),
    "Consciëntieusheid": (14, (("E6", 1), ("E11", -1), ("E16", 1), ("E21", -1), ("E26", 1), ("E31", -1), ("E36", 1), ("E41", -1), ("E46", 1), ("E51", 1))),
    "Neuroticisme": (38, (("F7", -1), ("F12", 1), ("F17", -1), ("F22", 1), ("F27", -1), ("F32", -1), ("F37", -1), ("F42", -1), ("F47", -1), ("F52", -1))),
    "Openheid": (8, (("G8", 1), ("G13", -1), ("G18", 1), ("G23", -1), ("G28", 1), ("G33", -1), ("G38", 1), ("G43", 1), ("G48", 1), ("G53", 1))),
}


def _number(value):
    if value is None or value == "":
        return 0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value
    raise ValueError(f"Expected a numeric assessment value, got {value!r}")


def _get_big_five_scores(sheet):
    return {
        trait: base + sum(_number(sheet[cell].value) * factor for cell, factor in terms)
        for trait, (base, terms) in BIG_FIVE_SCORE_FORMULAS.items()
    }


def _get_big_five_interpretations(scores, data_sheet):
    description_rows = {
        "Extraversie": 25,
        "Altruïsme": 26,
        "Consciëntieusheid": 27,
        "Neuroticisme": 28,
        "Openheid": 29,
    }
    interpretations = {}
    for trait, score in scores.items():
        description_column = "C" if score <= 17 else "B" if score <= 26 else "A"
        description = data_sheet[f"{description_column}{description_rows[trait]}"].value
        if not description:
            raise ValueError(f"Geen score-interpretatie gevonden voor {trait}.")
        interpretations[trait] = description
    return interpretations


def _get_highest_anchors(sheet):
    anchor_names = {
        "V": "Omhoog komen",
        "W": "Veilig voelen",
        "X": "Vrij zijn",
        "Y": "Balans vinden",
        "Z": "Uitdaging zoeken",
    }
    scores = {
        code: sum(_number(sheet[f"{column}{row}"].value) for row in range(4, 100))
        for code, column in zip(anchor_names, "CDEFG")
    }
    return [anchor_names[code] for code, score in sorted(scores.items(), key=lambda item: item[1], reverse=True) if score > 0][:2]


def _get_highest_clusters(sheet):
    scores = []
    for cluster_id in range(1, 17):
        first_row = 4 + (cluster_id - 1) * 8
        name = sheet[f"I{first_row}"].value
        score = sum(_number(sheet[f"C{row}"].value) for row in range(first_row, first_row + 7))
        score += sum(_number(sheet[f"E{row}"].value) for row in range(first_row, first_row + 5))
        score += sum(_number(sheet[f"G{row}"].value) for row in range(first_row, first_row + 5))
        if name and score > 0:
            scores.append((str(name).strip(), score))
    return [name for name, _ in sorted(scores, key=lambda item: item[1], reverse=True)[:2]]


def _get_highest_cultures(sheet):
    scores = []
    for first_row, name_row in zip((2, 7, 12, 17), (2, 8, 13, 18)):
        name = sheet[f"E{name_row}"].value
        score = sum(_number(sheet[f"C{row}"].value) for row in range(first_row, first_row + 4))
        if name and score > 0:
            scores.append((str(name).strip(), score))
    return [name for name, _ in sorted(scores, key=lambda item: item[1], reverse=True)[:2]]


def _get_client_career_phase(client):
    value = client.get("Date of Birth")
    if isinstance(value, date):
        birth_date = value
    else:
        birth_date = None
        for date_format in ("%d-%m-%Y", "%d-%m-%y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                birth_date = datetime.strptime(str(value), date_format).date()
                break
            except ValueError:
                continue
    if birth_date is None:
        raise ValueError("De geboortedatum van de cliënt ontbreekt of heeft een ongeldig formaat.")

    today = date.today()
    age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
    if 15 <= age <= 24:
        return "Verkenning"
    if 25 <= age <= 44:
        return "Keuze"
    if 45 <= age <= 64:
        return "Onderhoud"
    if age >= 65:
        return "Terugtrekking"
    raise ValueError("De cliënt moet minimaal 15 jaar oud zijn om een loopbaanfase te bepalen.")


def _write_after_label(sheet, label, value, occurrence=1):
    matches = [
        cell
        for row in sheet.iter_rows()
        for cell in row
        if cell.value is not None and str(cell.value).strip() == label
    ]
    if len(matches) < occurrence:
        raise ValueError(f"Rapport-label niet gevonden: {label} (voorkomen {occurrence})")
    cell = matches[occurrence - 1]
    sheet.cell(row=cell.row, column=cell.column + 1, value=value)


def generate_report(excel_path, client=None):
    wb = load_workbook(excel_path)

    missing_sheets = REQUIRED_SHEETS.difference(wb.sheetnames)
    if missing_sheets:
        wb.close()
        raise ValueError(f"Dit is geen resultatenbestand. Ontbrekende tabbladen: {', '.join(sorted(missing_sheets))}")

    # Get all sheets
    rapport = wb["Rapport"]
    fase11 = wb["Fase 1.1 | Big Five Dimensies"]
    fase20 = wb["Fase 2.0 | Loopbaanankers"]
    fase21 = wb["Fase 2.1 | Carriere Clusters"]
    fase22 = wb["Fase 2.2 | Cultuur analyse"]
    fase23 = wb["Fase 2.3 | J.C.M."]
    gegevens = wb["Gegevens"]

    big_five_scores = _get_big_five_scores(fase11)
    big_five_interpretations = _get_big_five_interpretations(big_five_scores, gegevens)

    # ============================================
    # 2. LOOPBAANANKERS (Fase 2.0)
    # ============================================
    top_anchors = _get_highest_anchors(fase20)
    anchor_1 = top_anchors[0] if len(top_anchors) > 0 else ""
    anchor_2 = top_anchors[1] if len(top_anchors) > 1 else ""

    # ============================================
    # 3. CARRIERE CLUSTERS (Fase 2.1)
    # ============================================
    top_clusters = _get_highest_clusters(fase21)
    cluster_1 = top_clusters[0] if len(top_clusters) > 0 else ""
    cluster_2 = top_clusters[1] if len(top_clusters) > 1 else ""

    # ============================================
    # 4. CULTUUR ANALYSE (Fase 2.2)
    # ============================================
    top_cultures = _get_highest_cultures(fase22)
    culture_1 = top_cultures[0] if len(top_cultures) > 0 else ""
    culture_2 = top_cultures[1] if len(top_cultures) > 1 else ""

    # ============================================
    # 5. J.C.M. (Fase 2.3) - FIXED: Reading from column D (index 3)
    # ============================================
    def get_jcm_scores():
        """Get the JCM scores from Fase 2.3."""
        jcm_scores = {}

        # Look for the JCM scores in the sheet
        # The answers are in column D (index 3) - "Toelichting" column
        for row in fase23.iter_rows(min_row=1, max_row=50, values_only=False):
            if row[0].value:
                cell_value = str(row[0].value).strip()
                # Check if this is one of the JCM components
                if cell_value in ["Taakvaardigheid", "Taakidentiteit", "Taakbetekenis", "Autonomie", "Feedback"]:
                    # The answer is in column D (index 3)
                    if len(row) > 3 and row[3].value:
                        jcm_scores[cell_value] = str(row[3].value)
                    else:
                        jcm_scores[cell_value] = ""

        return jcm_scores

    jcm_scores = get_jcm_scores()

    if client is not None:
        _write_after_label(rapport, "Loopbaan fase *", _get_client_career_phase(client))

    for trait, interpretation in big_five_interpretations.items():
        _write_after_label(rapport, f"{trait} *", interpretation)

    for occurrence, value in enumerate((anchor_1, anchor_2), start=1):
        _write_after_label(rapport, f"Hoogste score {occurrence} *", value)
    for label, value in zip(("Hoogste score 1 *", "Hoogste score 2 *"), (cluster_1, cluster_2)):
        _write_after_label(rapport, label, value, occurrence=2)
    for label, value in zip(("Hoogste score 1 *", "Hoogste score 2 *"), (culture_1, culture_2)):
        _write_after_label(rapport, label, value, occurrence=3)

    for component, value in jcm_scores.items():
        _write_after_label(rapport, component, value)

    wb.save(excel_path)
    wb.close()
    return excel_path


def generate_word_report(excel_path, client=None):
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        if "Rapport" not in wb.sheetnames:
            raise ValueError("Het tabblad Rapport ontbreekt in het resultatenbestand.")
        rapport = wb["Rapport"]
        report_rows = [
            [cell.value for cell in row if cell.value is not None]
            for row in rapport.iter_rows()
        ]
    finally:
        wb.close()

    document = Document()
    document.add_heading("Loopbaanrapport", level=0)
    if client and client.get("name"):
        document.add_paragraph(client["name"], style="Subtitle")

    big_five_traits = set(BIG_FIVE_SCORE_FORMULAS)
    for values in report_rows:
        values = [str(value).strip() for value in values]
        if not values or values == ["* Selecteer optie in uitvouwmenu"]:
            continue

        if len(values) == 1:
            text = values[0]
            if text.startswith("Fase "):
                document.add_heading(text, level=1)
            else:
                document.add_paragraph(text)
            continue

        label, value = values[:2]
        label = label.rstrip("*").strip()
        if label in big_five_traits:
            document.add_heading(label, level=2)
            document.add_paragraph(value)
        else:
            paragraph = document.add_paragraph()
            paragraph.add_run(f"{label}: ").bold = True
            paragraph.add_run(value)

    word_path = Path(excel_path).with_name(f"{Path(excel_path).stem}_report.docx")
    document.save(word_path)
    return str(word_path)


# ============================================
# DEBUG FUNCTION
# ============================================
def debug_excel_structure(excel_path):
    """Print the structure of the Excel file to help find the right cells."""
    wb = load_workbook(excel_path)

    print("\n=== Fase 2.0 | Loopbaanankers ===")
    fase20 = wb["Fase 2.0 | Loopbaanankers"]
    for row_idx, row in enumerate(fase20.iter_rows(min_row=1, max_row=150, values_only=False), 1):
        if row[0].value and "Totaal score" in str(row[0].value):
            print(f"Row {row_idx}: A='{row[0].value}'")
            print(f"  D (V/Omhoog): {row[3].value if len(row) > 3 else 'EMPTY'}")
            print(f"  E (W/Veilig): {row[4].value if len(row) > 4 else 'EMPTY'}")
            print(f"  F (X/Vrij): {row[5].value if len(row) > 5 else 'EMPTY'}")
            print(f"  G (Y/Balans): {row[6].value if len(row) > 6 else 'EMPTY'}")
            print(f"  H (Z/Uitdaging): {row[7].value if len(row) > 7 else 'EMPTY'}")

    print("\n=== Fase 2.1 | Carriere Clusters ===")
    fase21 = wb["Fase 2.1 | Carriere Clusters"]
    for row_idx, row in enumerate(fase21.iter_rows(min_row=10, max_row=200, values_only=False), 1):
        if row[0].value is not None:
            try:
                cluster_num = int(row[0].value)
                if 1 <= cluster_num <= 16:
                    if len(row) > 8:
                        print(f"Cluster {cluster_num}: Score = {row[8].value if row[8].value else 0}")
            except (ValueError, TypeError):
                pass

    print("\n=== Fase 2.2 | Cultuur analyse ===")
    fase22 = wb["Fase 2.2 | Cultuur analyse"]
    for row_idx, row in enumerate(fase22.iter_rows(min_row=5, max_row=25, values_only=False), 1):
        if row[0].value is not None:
            try:
                culture_num = int(row[0].value)
                if 1 <= culture_num <= 4:
                    if len(row) > 4:
                        print(f"Cultuur {culture_num}: Score = {row[4].value if row[4].value else 0}")
            except (ValueError, TypeError):
                pass

    print("\n=== Fase 2.3 | J.C.M. ===")
    fase23 = wb["Fase 2.3 | J.C.M."]
    for row_idx, row in enumerate(fase23.iter_rows(min_row=1, max_row=50, values_only=False), 1):
        if row[0].value:
            val = str(row[0].value).strip()
            if val in ["Taakvaardigheid", "Taakidentiteit", "Taakbetekenis", "Autonomie", "Feedback"]:
                print(f"Row {row_idx}: {val}")
                if len(row) > 3:
                    print(f"  Answer (Column D): {row[3].value if row[3].value else 'EMPTY'}")
                if len(row) > 2:
                    print(f"  Question (Column C): {row[2].value if row[2].value else 'EMPTY'}")

    wb.close()