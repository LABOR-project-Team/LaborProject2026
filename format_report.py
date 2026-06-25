from openpyxl import load_workbook


def generate_report(excel_path):
    wb = load_workbook(excel_path)

    # Get all sheets
    rapport = wb["Rapport"]
    fase11 = wb["Fase 1.1 | Big Five Dimensies"]
    fase20 = wb["Fase 2.0 | Loopbaanankers"]
    fase21 = wb["Fase 2.1 | Carriere Clusters"]
    fase22 = wb["Fase 2.2 | Cultuur analyse"]
    fase23 = wb["Fase 2.3 | J.C.M."]

    # ============================================
    # 1. BIG FIVE SCORES
    # ============================================
    def get_score_from_sheet(sheet, score_name):
        column_mapping = {
            "Extraversie": 2,  # Column B
            "Altruïsme": 3,  # Column C
            "Consciëntieusheid": 4,  # Column D
            "Neuroticisme": 5,  # Column E
            "Openheid": 6  # Column F
        }

        col_idx = column_mapping.get(score_name)
        if not col_idx:
            return ""

        for row in sheet.iter_rows(min_row=1, max_row=100, values_only=False):
            if row[0].value == "Score":
                score_cell = row[col_idx - 1]
                return score_cell.value if score_cell.value is not None else ""

        return ""

    extraversie_score = get_score_from_sheet(fase11, "Extraversie")
    altruisme_score = get_score_from_sheet(fase11, "Altruïsme")
    conscientieusheid_score = get_score_from_sheet(fase11, "Consciëntieusheid")
    neuroticisme_score = get_score_from_sheet(fase11, "Neuroticisme")
    openheid_score = get_score_from_sheet(fase11, "Openheid")

    # ============================================
    # 2. LOOPBAANANKERS (Fase 2.0)
    # ============================================
    def get_highest_anchors():
        """Get the two highest scoring career anchors from Fase 2.0."""
        anchor_names = {
            "V": "Omhoog komen",
            "W": "Veilig voelen",
            "X": "Vrij zijn",
            "Y": "Balans vinden",
            "Z": "Uitdaging zoeken"
        }

        for row in fase20.iter_rows(min_row=1, max_row=150, values_only=False):
            if row[0].value and "Totaal score" in str(row[0].value):
                # The scores are in columns D, E, F, G, H (indices 3, 4, 5, 6, 7)
                scores = {
                    "V": row[3].value if len(row) > 3 and row[3].value is not None else 0,
                    "W": row[4].value if len(row) > 4 and row[4].value is not None else 0,
                    "X": row[5].value if len(row) > 5 and row[5].value is not None else 0,
                    "Y": row[6].value if len(row) > 6 and row[6].value is not None else 0,
                    "Z": row[7].value if len(row) > 7 and row[7].value is not None else 0
                }

                # Convert to int/float and sort
                sorted_anchors = []
                for anchor_key, score in scores.items():
                    try:
                        score_value = float(score) if score else 0
                        sorted_anchors.append((anchor_key, score_value))
                    except (ValueError, TypeError):
                        sorted_anchors.append((anchor_key, 0))

                # Sort by score (highest first)
                sorted_anchors.sort(key=lambda x: x[1], reverse=True)

                # Get top 2
                top_2 = []
                for anchor_key, score in sorted_anchors[:2]:
                    if score > 0:
                        top_2.append(anchor_names.get(anchor_key, anchor_key))

                return top_2
        return []

    # Get top 2 anchors
    top_anchors = get_highest_anchors()
    anchor_1 = top_anchors[0] if len(top_anchors) > 0 else ""
    anchor_2 = top_anchors[1] if len(top_anchors) > 1 else ""

    # ============================================
    # 3. CARRIERE CLUSTERS (Fase 2.1)
    # ============================================
    def get_highest_clusters():
        """Get the two highest scoring career clusters from Fase 2.1."""
        cluster_scores = {}
        cluster_names = {
            1: "Landbouw, voeding en natuurlijke grondstoffen",
            2: "Architectuur en constructie",
            3: "Kunst, audio-visuele technologie en communicatie",
            4: "Business Management en administratie",
            5: "Educatie en training",
            6: "Financiën",
            7: "Overheid en publieke administratie",
            8: "Gezondheidswetenschappen",
            9: "Hospitality en toerisme",
            10: "Humanitaire dienstverlening",
            11: "ICT",
            12: "Publieke veiligheid en zekerheid",
            13: "Fabricage",
            14: "Marketing, sales en service",
            15: "Wetenschap, technologie, engineering en mathematica",
            16: "Transport, distributie en logistiek"
        }

        # Look for cluster scores in column I (index 8) - Totaal score
        for row in fase21.iter_rows(min_row=10, max_row=200, values_only=False):
            if row[0].value is not None:
                try:
                    cluster_num = int(row[0].value)
                    if 1 <= cluster_num <= 16:
                        if len(row) > 8 and row[8].value is not None:
                            try:
                                score = float(row[8].value)
                                if score > 0:
                                    cluster_scores[cluster_num] = score
                            except (ValueError, TypeError):
                                pass
                except (ValueError, TypeError):
                    pass

        sorted_clusters = sorted(cluster_scores.items(), key=lambda x: x[1], reverse=True)

        top_2 = []
        for cluster_num, score in sorted_clusters[:2]:
            top_2.append(cluster_names.get(cluster_num, f"Cluster {cluster_num}"))

        return top_2

    top_clusters = get_highest_clusters()
    cluster_1 = top_clusters[0] if len(top_clusters) > 0 else ""
    cluster_2 = top_clusters[1] if len(top_clusters) > 1 else ""

    # ============================================
    # 4. CULTUUR ANALYSE (Fase 2.2)
    # ============================================
    def get_highest_cultures():
        """Get the two highest scoring cultures from Fase 2.2."""
        culture_scores = {}
        culture_names = {
            1: "Mensgerichte cultuur",
            2: "Innovatieve cultuur",
            3: "Beheersgerichte cultuur",
            4: "Resultaatgerichte cultuur"
        }

        for row in fase22.iter_rows(min_row=5, max_row=25, values_only=False):
            if row[0].value is not None:
                try:
                    culture_num = int(row[0].value)
                    if 1 <= culture_num <= 4:
                        if len(row) > 4 and row[4].value is not None:
                            try:
                                score = float(row[4].value)
                                if score > 0:
                                    culture_scores[culture_num] = score
                            except (ValueError, TypeError):
                                pass
                except (ValueError, TypeError):
                    pass

        sorted_cultures = sorted(culture_scores.items(), key=lambda x: x[1], reverse=True)

        top_2 = []
        for culture_num, score in sorted_cultures[:2]:
            top_2.append(culture_names.get(culture_num, f"Cultuur {culture_num}"))

        return top_2

    top_cultures = get_highest_cultures()
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

    # ============================================
    # WRITE TO RAPPORT SHEET
    # ============================================
    def write_after_label(label, value):
        """Find a cell with the label and write the value in the cell to its right."""
        for row in rapport.iter_rows():
            for cell in row:
                if cell.value and str(cell.value).strip() == label:
                    rapport.cell(row=cell.row, column=cell.column + 1).value = value
                    print(f"  Wrote '{value}' after label '{label}'")
                    return
        print(f"  Warning: Label '{label}' not found in Rapport sheet")

    # Write Big Five scores
    print("\nWriting Big Five scores...")
    write_after_label("Extraversie *", extraversie_score)
    write_after_label("Altruïsme *", altruisme_score)
    write_after_label("Consciëntieusheid *", conscientieusheid_score)
    write_after_label("Neuroticisme *", neuroticisme_score)
    write_after_label("Openheid *", openheid_score)

    # Write Loopbaanankers
    print("\nWriting Loopbaanankers...")
    write_after_label("Hoogste score 1 *", anchor_1)
    write_after_label("Hoogste score 2 *", anchor_2)

    # Write Carriere Clusters
    print("\nWriting Carriere Clusters...")
    cluster_labels_found = 0
    for row in rapport.iter_rows():
        for cell in row:
            if cell.value and "Hoogste score" in str(cell.value):
                cluster_labels_found += 1
                if cluster_labels_found == 1:
                    continue
                elif cluster_labels_found == 2:
                    rapport.cell(row=cell.row, column=cell.column + 1).value = cluster_1
                    print(f"  Wrote cluster 1: '{cluster_1}'")
                elif cluster_labels_found == 3:
                    rapport.cell(row=cell.row, column=cell.column + 1).value = cluster_2
                    print(f"  Wrote cluster 2: '{cluster_2}'")
                    break

    # Write Cultuur scores
    print("\nWriting Cultuur scores...")
    culture_labels_found = 0
    for row in rapport.iter_rows():
        for cell in row:
            if cell.value and "Hoogste score" in str(cell.value):
                culture_labels_found += 1
                if culture_labels_found <= 3:
                    continue
                elif culture_labels_found == 4:
                    rapport.cell(row=cell.row, column=cell.column + 1).value = culture_1
                    print(f"  Wrote culture 1: '{culture_1}'")
                elif culture_labels_found == 5:
                    rapport.cell(row=cell.row, column=cell.column + 1).value = culture_2
                    print(f"  Wrote culture 2: '{culture_2}'")
                    break

    # Write JCM scores
    print("\nWriting JCM scores...")
    for component, value in jcm_scores.items():
        write_after_label(component, value)

    # Save the workbook
    wb.save(excel_path)
    print(f"\n✅ Report successfully formatted and saved to {excel_path}")


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