# app/prognosis_report.py
from fastapi import APIRouter, Request, Form, BackgroundTasks
from fastapi.responses import HTMLResponse, RedirectResponse
from pathlib import Path
import json
from typing import Optional

# Import utilities
from app.utils import get_client, save_client
from app.clients import save_questionnaire_results

router = APIRouter(prefix="/prognosis", tags=["prognosis"])


# Load data
def load_json(filename):
    """Load JSON data from app/data/"""
    json_path = Path(f"app/data/{filename}")
    if json_path.exists():
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


# Load all data
FASE_1_0_DATA = load_json("fase_1_0.json")
PROGNOSIS_QUESTIONS = FASE_1_0_DATA.get("prognosis_questions", []) if FASE_1_0_DATA else []
DONALD_SUPER = FASE_1_0_DATA.get("donald_super", []) if FASE_1_0_DATA else []
LEVINSON = FASE_1_0_DATA.get("levinson", []) if FASE_1_0_DATA else []


def render_template(content, title="Prognose Rapport"):
    """Wrap content in base HTML template with external CSS"""
    return f"""<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <link rel="stylesheet" href="/static/style.css">
</head>
<body>
    {content}
</body>
</html>"""


def render_question(question_data, existing_data):
    """Render a single question based on its type"""
    q_id = question_data.get("id")
    q_text = question_data.get("question")
    q_type = question_data.get("type")
    options = question_data.get("options", [])

    html = f"""
    <div class="question">
        <span class="question-number">{q_id}.</span>
        <span class="question-text">{q_text}</span>
        <div style="display: flex; gap: 15px; margin-left: auto; flex-wrap: wrap;">
    """

    if q_type == "radio":
        for option in options:
            checked = 'checked' if existing_data.get(f"q{q_id}") == option else ''
            html += f'<label><input type="radio" name="q{q_id}" value="{option}" {checked}> {option}</label>'

    elif q_type == "checkbox":
        existing_values = existing_data.get(f"q{q_id}", [])
        if not isinstance(existing_values, list):
            existing_values = []
        for option in options:
            checked = 'checked' if option in existing_values else ''
            html += f'<label><input type="checkbox" name="q{q_id}" value="{option}" {checked}> {option}</label>'

    elif q_type == "select":
        html += f'''
        <select name="q{q_id}" style="padding: 8px 12px; border-radius: 6px; border: 1px solid #ddd;">
            <option value="">Selecteer...</option>
        '''
        for option in options:
            selected = 'selected' if existing_data.get(f"q{q_id}") == option else ''
            html += f'<option value="{option}" {selected}>{option}</option>'
        html += '</select>'

    html += """
        </div>
    </div>
    """
    return html


def get_age_from_birthdate(geboortedatum: str) -> int:
    """Calculate age from birthdate string"""
    if not geboortedatum:
        return None
    try:
        from datetime import datetime
        birth = datetime.strptime(geboortedatum, "%Y-%m-%d")
        today = datetime.now()
        age = today.year - birth.year - ((today.month, today.day) < (birth.month, birth.day))
        return age
    except:
        return None


def get_age_range(age: int) -> str:
    """Get age range string from age"""
    if age is None:
        return None
    if age <= 25:
        return "15-25"
    elif age <= 45:
        return "25-45"
    elif age <= 65:
        return "46-65"
    else:
        return "65+"


@router.get("/", response_class=HTMLResponse)
async def prognosis_index(request: Request, client_id: str = None):
    """Prognosis questionnaire - Main page"""
    if not client_id:
        client_id = request.query_params.get('client_id')

    client = get_client(client_id) if client_id else None
    prognosis_data = client.get("prognosis", {}) if client else {}

    # Auto-fill age from client birthdate if available
    if client and client.get('geboortedatum') and not prognosis_data.get('q2'):
        age = get_age_from_birthdate(client.get('geboortedatum'))
        if age:
            age_range = get_age_range(age)
            prognosis_data['q2'] = age_range

    # Group questions by section
    sections = {}
    for q in PROGNOSIS_QUESTIONS:
        section = q.get("section", "Algemeen")
        if section not in sections:
            sections[section] = []
        sections[section].append(q)

    # Build questions HTML
    questions_html = ""
    for section_name, questions in sections.items():
        questions_html += f"""
        <div class="dimension-section">
            <div class="dimension-header">
                <span class="name">{section_name}</span>
            </div>
        """
        for q in questions:
            questions_html += render_question(q, prognosis_data)
        questions_html += "</div>"

    content = f"""
    <div class="container">
        <div class="header">
            <h1>📊 Integratie Prognose Model</h1>
            <div style="display: flex; gap: 10px;">
                <a href="/clients/{client_id if client_id else ''}" class="btn-back">← Terug naar Klant</a>
                <a href="/" class="btn-back">← Home</a>
            </div>
        </div>

        <div class="card">
            <p style="color: #666; margin-bottom: 20px;">
                Deze vragenlijst brengt uw persoonlijke situatie, vaardigheden en arbeidsmarktpositie in kaart. 
                Op basis van uw antwoorden wordt een prognose opgesteld over uw kansen op de arbeidsmarkt.
            </p>
        </div>

        <form method="post" action="/prognosis/submit">
            <input type="hidden" name="client_id" value="{client_id if client_id else ''}">
            {questions_html}
            <div class="footer">
                <div class="info">
                    <span>Alle vragen moeten beantwoord worden</span>
                </div>
                <button type="submit" class="btn-primary">💾 Opslaan & Rapport Genereren</button>
            </div>
        </form>
    </div>
    """
    return HTMLResponse(content=render_template(content, "Prognose Vragenlijst"))


@router.post("/submit")
async def submit_prognosis(request: Request, background_tasks: BackgroundTasks):
    """Handle prognosis questionnaire submission"""
    form_data = await request.form()
    client_id = form_data.get('client_id')

    # Collect all answers
    answers = {}
    for key, value in form_data.items():
        if key.startswith('q'):
            # Handle checkbox arrays
            if key in answers:
                if isinstance(answers[key], list):
                    answers[key].append(value)
                else:
                    answers[key] = [answers[key], value]
            else:
                answers[key] = value

    # Convert checkbox answers to lists
    for key, value in list(answers.items()):
        if isinstance(value, str) and key in form_data.getlist(key):
            answers[key] = form_data.getlist(key)

    # Prepare results for saving
    results = {
        "completed": True,
        "answers": answers
    }

    # Save to client with background tasks
    if client_id:
        save_questionnaire_results(client_id, "prognosis", results, background_tasks)

        return RedirectResponse(
            url=f"/prognosis/result?client_id={client_id}",
            status_code=303
        )

    return RedirectResponse(url="/", status_code=303)


def get_career_phase(age_range: str) -> dict:
    """Get career phase based on age range"""
    if not age_range:
        return {"phase": "Onbekend", "description": "Geen leeftijd opgegeven"}

    # Parse age range
    if '-' in age_range:
        parts = age_range.split('-')
        try:
            age_min = int(parts[0])
            age_max = int(parts[1]) if len(parts) > 1 and parts[1] else 100
        except:
            return {"phase": "Onbekend", "description": "Ongeldige leeftijd"}
    else:
        return {"phase": "Onbekend", "description": "Ongeldige leeftijd"}

    # Check Donald Super phases
    for phase in DONALD_SUPER:
        leeftijd = phase.get('leeftijd', '')
        if '-' in leeftijd:
            p_min, p_max = leeftijd.split('-')
            try:
                p_min = int(p_min)
                p_max = int(p_max) if p_max else 100
                if p_min <= age_min and age_max <= p_max:
                    return {"phase": phase.get('fase', ''), "description": phase.get('omschrijving', '')}
            except:
                continue

    return {"phase": "Onbekend", "description": "Geen passende fase gevonden"}


@router.get("/result", response_class=HTMLResponse)
async def prognosis_result(request: Request, client_id: str = None):
    """Show prognosis results"""
    if not client_id:
        client_id = request.query_params.get('client_id')

    client = get_client(client_id) if client_id else None
    if not client:
        return HTMLResponse(content="<h1>Client not found</h1>")

    prognosis = client.get("prognosis", {})
    answers = prognosis.get("answers", {})

    # Get career phase
    age = answers.get('q2', '')
    career_phase = get_career_phase(age)

    # Build result content
    content = f"""
    <div class="container">
        <div class="header">
            <h1>📊 Prognose Rapport</h1>
            <div style="display: flex; gap: 10px;">
                <a href="/clients/{client_id}" class="btn-back">← Terug naar Klant</a>
                <a href="/" class="btn-back">← Home</a>
            </div>
        </div>

        <div class="results-card">
            <h2 style="margin-bottom: 20px;">Klant: {client.get('name', 'Onbekend')}</h2>

            <div class="top-results">
                <h3>📈 Loopbaanfase</h3>
                <div class="top-item">Fase: {career_phase.get('phase', 'Onbekend')}</div>
                <p style="margin-top: 10px; color: #555;">{career_phase.get('description', '')}</p>
            </div>

            <h3 style="margin: 20px 0 15px;">Antwoorden</h3>
    """

    # Show all answers grouped by section
    sections = {}
    for q in PROGNOSIS_QUESTIONS:
        section = q.get("section", "Algemeen")
        if section not in sections:
            sections[section] = []
        sections[section].append(q)

    for section_name, questions in sections.items():
        content += f"""
        <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
            <h4 style="color: #4a6cf7; margin-bottom: 10px;">{section_name}</h4>
        """
        for q in questions:
            q_id = q.get("id")
            q_text = q.get("question")
            answer = answers.get(f"q{q_id}", "Niet beantwoord")
            if isinstance(answer, list):
                answer = ", ".join(answer)
            content += f"""
            <div style="display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid #eee;">
                <span style="font-weight: 500;">{q_id}. {q_text}</span>
                <span style="color: #333;">{answer}</span>
            </div>
            """
        content += "</div>"

    content += """
        </div>
    </div>
    """
    return HTMLResponse(content=render_template(content, "Prognose Resultaten"))


@router.get("/test")
async def test_prognosis():
    """Test route to check loaded questions"""
    return {
        "total_questions": len(PROGNOSIS_QUESTIONS),
        "questions": PROGNOSIS_QUESTIONS,
        "donald_super": DONALD_SUPER,
        "levinson": LEVINSON
    }