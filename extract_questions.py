"""
extract_questions_fixed.py
Extracts all questionnaire data from the Excel file and converts to JSON format.
"""

import json
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any

# Configuration
EXCEL_FILE = "Loopbaan onderzoek 5.0.xlsx"
OUTPUT_DIR = "data"  # Where JSON files will be saved


def create_output_directory():
    """Create the output directory if it doesn't exist."""
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)


def extract_fase_1_0(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Extract Phase 1.0 data (Loopbaanfasen from Donald Super and Levinson)
    """
    donald_super = []
    levinson = []

    # Find the starting points by looking for specific text
    donald_start = None
    levinson_start = None

    for idx, row in df.iterrows():
        if isinstance(row.iloc[0], str) and "Loopbaanfasen van Donald Super" in row.iloc[0]:
            donald_start = idx + 2  # Skip header row
        if isinstance(row.iloc[0], str) and "Loopbaanfasen van Levinson en Levinson" in row.iloc[0]:
            levinson_start = idx + 2  # Skip header row
            break

    # Extract Donald Super phases
    if donald_start is not None:
        for idx in range(donald_start, min(donald_start + 8, len(df)), 2):
            if idx < len(df):
                row = df.iloc[idx]
                if pd.notna(row.iloc[0]):
                    phase_data = {
                        "leeftijd": str(row.iloc[0]) if pd.notna(row.iloc[0]) else "",
                        "fase": str(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
                        "omschrijving": str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                    }
                    donald_super.append(phase_data)

    # Extract Levinson phases
    if levinson_start is not None:
        for idx in range(levinson_start, len(df), 2):
            if idx < len(df):
                row = df.iloc[idx]
                if pd.notna(row.iloc[0]) and "Leeftijd" not in str(row.iloc[0]):
                    phase_data = {
                        "leeftijd": str(row.iloc[0]) if pd.notna(row.iloc[0]) else "",
                        "fase": str(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
                        "omschrijving": str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                    }
                    levinson.append(phase_data)

    return {
        "donald_super": donald_super,
        "levinson": levinson
    }


def extract_fase_1_1(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Extract Phase 1.1 data (Big Five Personality dimensions)
    """
    questions = []

    # Find the start of the questions - row with "Nummer" in column A
    start_row = None
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
            if "Nummer" in str(row.iloc[0]):
                start_row = idx + 1  # Start after the header
                break

    if start_row is None:
        return {"questions": [], "scoring_formulas": {}}

    # Extract questions - they start from the found row
    for idx in range(start_row, min(start_row + 50, len(df))):
        row = df.iloc[idx]

        # Check if this is a question row (has a number in column A)
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], (int, float)):
            question_num = int(row.iloc[0])

            # Get the question text from column B (index 1)
            question_text = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""

            # Clean up the text (remove trailing spaces)
            question_text = question_text.strip()

            # Determine which dimension this question belongs to
            # The dimension headers are in row 2: Extraversie (col 2), Altruïsme (col 3),
            # Conciëntieusheid (col 4), Neuroticisme (col 5), Openheid (col 6)
            # But the actual data has the values in the cells, not the headers
            # Based on the Excel structure, the first 10 questions map to specific dimensions
            # Let's use the mapping from the Excel formulas

            # Map question numbers to dimensions based on the pattern
            # From the Excel data:
            # Extraversie: 1, 6, 11, 16, 21, 26, 31, 36, 41, 46
            # Altruïsme: 2, 7, 12, 17, 22, 27, 32, 37, 42, 47
            # Conciëntieusheid: 3, 8, 13, 18, 23, 28, 33, 38, 43, 48
            # Neuroticisme: 4, 9, 14, 19, 24, 29, 34, 39, 44, 49
            # Openheid: 5, 10, 15, 20, 25, 30, 35, 40, 45, 50

            dimension = ""
            if question_num in [1, 6, 11, 16, 21, 26, 31, 36, 41, 46]:
                dimension = "extraversie"
            elif question_num in [2, 7, 12, 17, 22, 27, 32, 37, 42, 47]:
                dimension = "altruïsme"
            elif question_num in [3, 8, 13, 18, 23, 28, 33, 38, 43, 48]:
                dimension = "conciëntieusheid"
            elif question_num in [4, 9, 14, 19, 24, 29, 34, 39, 44, 49]:
                dimension = "neuroticisme"
            elif question_num in [5, 10, 15, 20, 25, 30, 35, 40, 45, 50]:
                dimension = "openheid"
            else:
                # Skip if no dimension found
                continue

            question = {
                "id": question_num,
                "text": question_text,
                "dimension": dimension,
                "reverse_scored": False  # Will be determined from formulas
            }
            questions.append(question)

    # Extract scoring formulas
    scoring_formulas = {}

    # Find the section with the scoring formulas
    scoring_start = None
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
            if "Score" in row.iloc[0] and "Omschrijving" in row.iloc[0]:
                scoring_start = idx + 1
                break

    if scoring_start is not None:
        scoring_categories = []
        for idx in range(scoring_start, min(scoring_start + 5, len(df))):
            row = df.iloc[idx]
            if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
                if "-" in row.iloc[0]:
                    score_range = row.iloc[0].split('-')
                    try:
                        min_score = int(score_range[0].strip())
                        max_score = int(score_range[1].strip())
                        category = {
                            "min_score": min_score,
                            "max_score": max_score,
                            "label": str(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
                            "formula": str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                        }
                        scoring_categories.append(category)
                    except (ValueError, IndexError):
                        continue
        scoring_formulas["categories"] = scoring_categories

    # Extract reverse-scored questions from the formulas
    reverse_scored = set()
    if scoring_start is not None:
        for idx in range(scoring_start, min(scoring_start + 5, len(df))):
            row = df.iloc[idx]
            if pd.notna(row.iloc[2]) and isinstance(row.iloc[2], str):
                formula = str(row.iloc[2])
                import re
                # Find all numbers in the formula that have a '-' before them
                reverse_matches = re.findall(r'-\s*\((\d+)\)', formula)
                for match in reverse_matches:
                    reverse_scored.add(int(match))

    # Update questions with reverse_scored flag
    for question in questions:
        if question["id"] in reverse_scored:
            question["reverse_scored"] = True

    # Add dimension descriptions
    dimension_descriptions = {
        "extraversie": "Mensen die hoog scoren op extraversie ervaren de behoefte om het gezelschap en de stimulering van anderen op te zoeken. Binnen groepsverband wil men veelal het middelpunt van aandacht zijn. De aandacht is dan ook extern gericht. Ze houden van opwinding, spanning en zijn energiek van aard. Ze zijn doorgaans enthousiast, assertief en optimistisch.",
        "altruïsme": "Altruïsme is de mate waarin iemand het belang van anderen boven zijn eigen belang stelt. Bij mensen die hoog scoren op altruïsme is de ander het onderwerp in de relatie. De relatie wordt dan ook vaak vanuit de ander beleefd. Ze verplaatsen zich daardoor makkelijk in een ander en bezien situaties vaak vanuit het doel van de ander. Altruïstische mensen zijn hulpvaardig, bescheiden, vriendelijk en geneigd tot samenwerken.",
        "conciëntieusheid": "Consciëntieusheid weerspiegelt de neiging om verantwoordelijk, georganiseerd, hardwerkend en doelgericht te zijn en zich te conformeren aan normen en regels. Mensen die hoog scoren op consciëntieusheid houden zich tevens vaak aan bepaalde levensrichtlijnen. Ze zullen bijvoorbeeld volharden bij een uitdaging en verantwoordelijkheid nemen voor problemen waarmee men geconfronteerd wordt.",
        "neuroticisme": "Personen die hoog scoren op neuroticisme ervaren meer dan gemiddeld gevoelens zoals angst, woede, frustratie, afgunst, jaloezie, schuld, depressiviteit en eenzaamheid. Mensen die neurotisch zijn, reageren dan ook slechter op stressoren, interpreteren gewone situaties eerder als bedreigend en ervaren kleine frustraties als uiterst ontmoedigend.",
        "openheid": "Personen met een hoge mate aan openheid hebben een versterkte waardering voor ongewone ideeën en het esthetische zoals kunst, architectuur en mode. Ze zijn meestal fantasierijk in plaats van praktisch. Veelal is men creatief, staat men open voor nieuwe en abstracte ideeën en staat men beter in contact met de eigen emotie."
    }
    scoring_formulas["dimension_descriptions"] = dimension_descriptions

    return {
        "questions": questions,
        "scoring_formulas": scoring_formulas
    }


def extract_fase_2_0(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Extract Phase 2.0 data (Loopbaanankers)
    """
    statements = []
    anchors = {}

    # Find the start of the statements - row with "Nummer" in column A
    start_row = None
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
            if "Nummer" in str(row.iloc[0]):
                start_row = idx + 1  # Start after the header
                break

    if start_row is None:
        return {"statements": [], "anchors": {}}

    # Extract statements
    current_statement = None
    current_id = None

    for idx in range(start_row, min(start_row + 60, len(df))):
        row = df.iloc[idx]

        # Check if this is a new statement (has a number in column A)
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], (int, float)):
            # Save previous statement if exists
            if current_statement is not None:
                statements.append(current_statement)

            current_id = int(row.iloc[0])

            # Get the statement text from column B (index 1)
            if pd.notna(row.iloc[1]):
                text = str(row.iloc[1]).strip()
                # Extract the anchor letter (V, W, X, Y, Z)
                anchor_letter = text[0] if text and text[0] in ['V', 'W', 'X', 'Y', 'Z'] else ''
                # Remove the anchor letter and colon
                clean_text = text[2:].strip() if len(text) > 2 and text[1] == ':' else text

                current_statement = {
                    "id": current_id,
                    "text": clean_text,
                    "anchor": anchor_letter
                }
            else:
                current_statement = None

        # Check if this is a continuation (no number in column A, but text in column B)
        elif current_statement is not None and pd.notna(row.iloc[1]):
            text = str(row.iloc[1]).strip()
            if text:
                # Check if it starts with a letter and colon (new statement without number)
                if len(text) >= 2 and text[0] in ['V', 'W', 'X', 'Y', 'Z'] and text[1] == ':':
                    # This is actually a new statement without a number
                    # Save previous and start new
                    statements.append(current_statement)
                    anchor_letter = text[0]
                    clean_text = text[2:].strip() if len(text) > 2 else text
                    current_statement = {
                        "id": current_id,  # Keep the same ID
                        "text": clean_text,
                        "anchor": anchor_letter
                    }
                else:
                    # Append to current statement
                    current_statement["text"] += " " + text

    # Add the last statement
    if current_statement is not None:
        statements.append(current_statement)

    # Extract anchor descriptions
    anchor_start = None
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
            if "Loopbaananker" in row.iloc[0]:
                anchor_start = idx + 1
                break

    if anchor_start is not None:
        current_anchor = None
        current_text = []

        for idx in range(anchor_start, min(anchor_start + 15, len(df))):
            row = df.iloc[idx]
            if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
                text = str(row.iloc[0]).strip()

                # Check for anchor headers
                if any(anchor in text for anchor in ["Omhoog", "Veilig", "Vrij", "Balans", "Uitdaging"]):
                    if current_anchor and current_text:
                        anchors[current_anchor] = " ".join(current_text)
                    current_anchor = text.split()[0] if text.split() else text
                    current_text = []
                elif current_anchor and text:
                    current_text.append(text)

        # Add the last anchor
        if current_anchor and current_text:
            anchors[current_anchor] = " ".join(current_text)

    return {
        "statements": statements,
        "anchors": anchors
    }


def extract_fase_2_1(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Extract Phase 2.1 data (Carrière Clusters)
    """
    clusters = []

    # Find the start - row with "Cluster" in column A
    start_row = None
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
            if "Cluster" in str(row.iloc[0]):
                start_row = idx + 1  # Start after the header
                break

    if start_row is None:
        return {"clusters": []}

    # Extract each cluster
    current_cluster = None
    current_activities = []
    current_competencies = []
    current_topics = []

    for idx in range(start_row, min(start_row + 100, len(df))):
        row = df.iloc[idx]

        # Check if this is a cluster number (has a number in column A)
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], (int, float)):
            # Save previous cluster if exists
            if current_cluster is not None:
                current_cluster["activities"] = current_activities
                current_cluster["competencies"] = current_competencies
                current_cluster["educational_topics"] = current_topics
                clusters.append(current_cluster)

            # Start a new cluster
            cluster_num = int(row.iloc[0])
            cluster_name = str(row.iloc[1]) if pd.notna(row.iloc[1]) else f"Cluster {cluster_num}"

            # Get segment and description from columns 8 and 9
            segment = str(row.iloc[8]) if pd.notna(row.iloc[8]) else ""
            description = str(row.iloc[9]) if pd.notna(row.iloc[9]) else ""

            current_cluster = {
                "id": cluster_num,
                "name": cluster_name,
                "segment": segment,
                "description": description,
                "activities": [],
                "competencies": [],
                "educational_topics": []
            }
            current_activities = []
            current_competencies = []
            current_topics = []

        # Extract activities (column C - index 2)
        elif current_cluster is not None and pd.notna(row.iloc[2]):
            activity = str(row.iloc[2])
            if activity and not activity.startswith("="):
                current_activities.append(activity.strip())

        # Extract competencies (column E - index 4)
        if current_cluster is not None and pd.notna(row.iloc[4]):
            competency = str(row.iloc[4])
            if competency and not competency.startswith("="):
                current_competencies.append(competency.strip())

        # Extract educational topics (column G - index 6)
        if current_cluster is not None and pd.notna(row.iloc[6]):
            topic = str(row.iloc[6])
            if topic and not topic.startswith("="):
                current_topics.append(topic.strip())

    # Add the last cluster
    if current_cluster is not None:
        current_cluster["activities"] = current_activities
        current_cluster["competencies"] = current_competencies
        current_cluster["educational_topics"] = current_topics
        clusters.append(current_cluster)

    return {"clusters": clusters}


def extract_fase_2_2(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Extract Phase 2.2 data (Cultuur Analyse)
    """
    cultures = []

    # Find the start - row with "Cultuur" in column A
    start_row = None
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
            if "Cultuur" in str(row.iloc[0]):
                start_row = idx + 1  # Start after the header
                break

    if start_row is None:
        return {"cultures": []}

    # Extract each culture
    for cluster_idx in range(4):
        start_idx = start_row + (cluster_idx * 5)
        culture_data = {
            "id": cluster_idx + 1,
            "questions": [],
            "culture": "",
            "description": ""
        }

        # Extract questions and culture info
        for offset in range(5):
            row_idx = start_idx + offset
            if row_idx >= len(df):
                break
            row = df.iloc[row_idx]

            # Extract question (column B - index 1)
            if pd.notna(row.iloc[1]):
                question = str(row.iloc[1])
                if question and not question.startswith("="):
                    culture_data["questions"].append(question.strip())

            # Extract culture name (column D - index 3)
            if pd.notna(row.iloc[3]) and isinstance(row.iloc[3], str):
                row_text = str(row.iloc[3]).strip()
                if row_text and "cultuur" in row_text.lower():
                    culture_data["culture"] = row_text

            # Extract description (column F - index 5)
            if pd.notna(row.iloc[5]) and isinstance(row.iloc[5], str):
                desc_text = str(row.iloc[5]).strip()
                if desc_text and not culture_data["description"]:
                    culture_data["description"] = desc_text

        if culture_data["questions"]:
            cultures.append(culture_data)

    return {"cultures": cultures}


def extract_fase_2_3(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Extract Phase 2.3 data (J.C.M. - Job Characteristics Model)
    """
    components = []

    # Find the JCM components - row with "Component" in column A
    start_row = None
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
            if "Component" in str(row.iloc[0]):
                start_row = idx + 1  # Start after the header
                break

    if start_row is None:
        return {"components": []}

    # Extract each component
    for offset in range(5):
        row_idx = start_row + (offset * 2)  # Each component has 2 rows (title + description)
        if row_idx >= len(df):
            break

        row = df.iloc[row_idx]
        component = {
            "id": offset + 1,
            "name": str(row.iloc[0]) if pd.notna(row.iloc[0]) else "",
            "description": str(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
            "question": str(row.iloc[2]) if pd.notna(row.iloc[2]) else "",
            "letter": str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
        }

        if component["name"]:
            components.append(component)

    return {"components": components}


def generate_json_files():
    """
    Main function to read the Excel file and generate all JSON files
    """
    print("📊 Reading Excel file...")

    # Read all sheets
    xlsx = pd.ExcelFile(EXCEL_FILE)

    # Process each phase
    results = {}

    # Fase 1.0
    print("📝 Extracting Fase 1.0...")
    df = pd.read_excel(xlsx, sheet_name="Fase 1.0 | Loopbaanfasen")
    results["fase_1_0"] = extract_fase_1_0(df)

    # Fase 1.1
    print("📝 Extracting Fase 1.1...")
    df = pd.read_excel(xlsx, sheet_name="Fase 1.1 | Big Five Diemensies")
    results["fase_1_1"] = extract_fase_1_1(df)

    # Fase 2.0
    print("📝 Extracting Fase 2.0...")
    df = pd.read_excel(xlsx, sheet_name="Fase 2.0 | Loopbaanankers")
    results["fase_2_0"] = extract_fase_2_0(df)

    # Fase 2.1
    print("📝 Extracting Fase 2.1...")
    df = pd.read_excel(xlsx, sheet_name="Fase 2.1 | carriere Clusters")
    results["fase_2_1"] = extract_fase_2_1(df)

    # Fase 2.2
    print("📝 Extracting Fase 2.2...")
    df = pd.read_excel(xlsx, sheet_name="Fase 2.2 | Cultuur analyse")
    results["fase_2_2"] = extract_fase_2_2(df)

    # Fase 2.3
    print("📝 Extracting Fase 2.3...")
    df = pd.read_excel(xlsx, sheet_name="Fase 2.3 | J.C.M.")
    results["fase_2_3"] = extract_fase_2_3(df)

    # Create output directory
    create_output_directory()

    # Write each phase to its own JSON file
    for phase_name, data in results.items():
        output_file = Path(OUTPUT_DIR) / f"{phase_name}.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"✅ Created: {output_file}")

    print("\n🎉 All JSON files generated successfully!")


def main():
    """Entry point for the script."""
    try:
        # Check if the Excel file exists
        if not Path(EXCEL_FILE).exists():
            print(f"❌ Error: Excel file '{EXCEL_FILE}' not found!")
            print("   Please make sure the file is in the same directory as this script.")
            return

        generate_json_files()

    except Exception as e:
        print(f"❌ An error occurred: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()