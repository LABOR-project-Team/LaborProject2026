from openpyxl import load_workbook
from docx import Document


def safe_float(value):
    """Converteert naar float, vervangt komma door punt."""
    if value is None: return 0.0
    s = str(value).strip().replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def generate_report(excel_path):
    wb = load_workbook(excel_path, data_only=True)
    doc = Document()
    doc.add_heading('Loopbaan Rapportage', 0)

    # 1. Big Five (Correcte drempels voor 5-40)
    # Low: 5-16, Medium: 17-28, High: 29-40
    fase11 = wb["Fase 1.1 | Big Five Dimensies"]
    ws_gegevens = wb["Gegevens"]
    doc.add_heading('Big Five Persoonlijkheid', level=1)

    dims = ["Extraversie", "Altruïsme", "Consciëntieusheid", "Neuroticisme", "Openheid"]
    for i, dim in enumerate(dims):
        # We zoeken de rij waar 'Score' staat en pakken kolom i+2
        score = 0
        for row in fase11.iter_rows(min_row=1, max_row=20):
            if row[0].value == "Score":
                score = safe_float(row[i + 1].value)
                break

        category = "High" if score > 28 else "Medium" if score > 16 else "Low"
        col_idx = {"High": 1, "Medium": 2, "Low": 3}[category]
        text = ws_gegevens.cell(row=25 + i, column=col_idx).value
        doc.add_paragraph(f"{dim}: {category}e score ({score})")
        doc.add_paragraph(str(text))

    # 2. Loopbaanankers
    doc.add_heading('Loopbaanankers', level=1)
    fase20 = wb["Fase 2.0 | Loopbaanankers"]
    # De data staat vaak in de kolom "Totaal score" (Kolom 8 of 9)
    # Pas de index aan als dit niet werkt
    for row in fase20.iter_rows(min_row=5, max_row=30):
        if row[0].value and isinstance(row[0].value, str):
            score = safe_float(row[7].value)  # Controleer of dit de juiste kolom is
            if score > 0:
                doc.add_paragraph(f"{row[0].value}: Score {score}")

    # 3. Carriere Clusters (Accurater)
    doc.add_heading('Carriere Clusters', level=1)
    fase21 = wb["Fase 2.1 | Carriere Clusters"]
    clusters = []
    for row in fase21.iter_rows(min_row=5, max_row=50):
        if row[0].value and str(row[0].value).strip().isdigit():
            # Kolom 8 is 'Totaal score'
            s = safe_float(row[7].value)
            if s > 0: clusters.append((f"Cluster {row[0].value}", s))
    for name, score in sorted(clusters, key=lambda x: x[1], reverse=True)[:3]:
        doc.add_paragraph(f"{name}: Score {score}")

    # 4. Cultuur Analyse
    doc.add_heading('Cultuur Analyse', level=1)
    fase22 = wb["Fase 2.2 | Cultuur analyse"]
    cults = []
    for row in fase22.iter_rows(min_row=5, max_row=20):
        if row[5].value:  # Kolom 5 bevat vaak de categorienaam
            s = safe_float(row[3].value)  # Kolom 3 is totaalscore
            if s > 0: cults.append((str(row[5].value), s))
    for name, score in cults:
        doc.add_paragraph(f"{name}: Score {score}")

    doc.save("Loopbaan_Rapport_V2.docx")
    print("Rapport gegenereerd.")