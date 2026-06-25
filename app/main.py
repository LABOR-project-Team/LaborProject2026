# app/main.py
import sys
import os
from pathlib import Path
from typing import Optional
from jinja2 import Environment, FileSystemLoader, select_autoescape

from fastapi import FastAPI, Request, Form, BackgroundTasks
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
import json


# ============================================================
# PATH CONFIGURATION FOR .EXE SUPPORT
# ============================================================

def get_base_path():
    """Get the base path for the application (supports .exe and script mode)"""
    if getattr(sys, 'frozen', False):
        # Running as .exe - use the temp folder where PyInstaller extracts files
        return sys._MEIPASS
    else:
        # Running as script - use the current directory
        return os.path.dirname(os.path.abspath(__file__))


# Set up paths
BASE_PATH = get_base_path()
TEMPLATE_PATH = os.path.join(BASE_PATH, 'app', 'templates')
STATIC_PATH = os.path.join(BASE_PATH, 'app', 'static')

# Also set up data path (for JSON files)
DATA_PATH = os.path.join(BASE_PATH, 'app', 'data')

# ============================================================
# IMPORTS
# ============================================================

# Import from app.clients
from app.clients import router as clients_router
from app.clients import save_questionnaire_results
from app.utils import get_client
from app.prognosis_report import router as prognosis_router

# ============================================================
# FASTAPI APP SETUP
# ============================================================

# Create FastAPI app
app = FastAPI(title="Loopbaan Onderzoek")

# Mount static files
app.mount("/static", StaticFiles(directory=STATIC_PATH), name="static")

# Include client routes
app.include_router(clients_router)
app.include_router(prognosis_router)


# ============================================================
# DATA LOADING FUNCTIONS
# ============================================================

def load_json(filename):
    """Load JSON data from app/data/"""
    # Try multiple paths for .exe support
    possible_paths = [
        Path(f"app/data/{filename}"),  # Normal script mode
        Path(DATA_PATH) / filename,  # .exe mode with base path
        Path(f"{BASE_PATH}/app/data/{filename}"),  # Alternative .exe path
    ]

    for json_path in possible_paths:
        if json_path.exists():
            with open(json_path, 'r', encoding='utf-8') as f:
                return json.load(f)

    print(f"Warning: Could not find {filename} in any of the expected paths")
    return None


# Load all data
BIG_FIVE_DATA = load_json("fase_1_1.json")
BIG_FIVE_QUESTIONS = BIG_FIVE_DATA.get("questions", []) if BIG_FIVE_DATA else []

CAREER_ANCHORS = load_json("fase_2_0.json")
CAREER_ANCHOR_STATEMENTS = CAREER_ANCHORS.get("statements", []) if CAREER_ANCHORS else []
CAREER_ANCHOR_DESCRIPTIONS = CAREER_ANCHORS.get("anchors", {}) if CAREER_ANCHORS else {}

CAREER_CLUSTERS = load_json("fase_2_1.json")
CAREER_CLUSTERS_DATA = CAREER_CLUSTERS.get("clusters", []) if CAREER_CLUSTERS else []


# ============================================================
# TEMPLATE RENDER FUNCTION
# ============================================================

def render_template(template_name: str, context: dict = None) -> str:
    """Render a Jinja2 template with the given context"""
    if context is None:
        context = {}

    # Set up Jinja2 environment with the correct path
    env = Environment(
        loader=FileSystemLoader(TEMPLATE_PATH),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template(template_name)
    return template.render(**context)


# ============================================================
# HOME PAGE
# ============================================================
@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    """Home page"""
    content = render_template("index.html", {"request": request})
    return HTMLResponse(content=content)


# ============================================================
# PHASE 1.0 - Coming Soon
# ============================================================
@app.get("/phase/1.0", response_class=HTMLResponse)
async def phase_1_0(request: Request):
    """Phase 1.0 - Career Phase based on age"""
    content = render_template("coming_soon.html", {"request": request, "phase": "1.0"})
    return HTMLResponse(content=content)


# ============================================================
# PHASE 1.1 - Big Five
# ============================================================
@app.get("/phase/1.1", response_class=HTMLResponse)
async def big_five_questionnaire(request: Request, client_id: str = None):
    """Big Five Personality Questionnaire - Phase 1.1"""
    if not BIG_FIVE_QUESTIONS:
        return HTMLResponse(
            content="<h1>Error: Questions not loaded</h1><p>Make sure fase_1_1.json exists in the data folder.</p>")

    if not client_id:
        client_id = request.query_params.get('client_id')

    # Group questions by dimension
    dimensions = {
        "extraversie": {"label": "Extraversie", "questions": []},
        "altruïsme": {"label": "Altruïsme", "questions": []},
        "conciëntieusheid": {"label": "Conciëntieusheid", "questions": []},
        "neuroticisme": {"label": "Neuroticisme", "questions": []},
        "openheid": {"label": "Openheid", "questions": []}
    }

    for q in BIG_FIVE_QUESTIONS:
        dim = q.get("dimension", "")
        if dim in dimensions:
            dimensions[dim]["questions"].append(q)

    content = render_template("phase_1_1.html", {
        "request": request,
        "client_id": client_id,
        "dimensions": dimensions,
        "total_questions": len(BIG_FIVE_QUESTIONS)
    })
    return HTMLResponse(content=content)


@app.post("/phase/1.1/submit")
async def submit_big_five(request: Request, background_tasks: BackgroundTasks):
    """Handle Big Five questionnaire submission"""
    form_data = await request.form()
    client_id = form_data.get('client_id')

    # Process answers
    answers = {}
    for key, value in form_data.items():
        if key.startswith('q'):
            q_num = int(key[1:])
            answers[q_num] = int(value)

    scores = {"extraversie": 0, "altruïsme": 0, "conciëntieusheid": 0, "neuroticisme": 0, "openheid": 0}

    for q in BIG_FIVE_QUESTIONS:
        q_id = q.get("id")
        dim = q.get("dimension")
        if q_id in answers and dim in scores:
            score = answers[q_id]
            if q.get("reverse_scored", False):
                score = 6 - score
            scores[dim] += score

    max_score = 50
    percentages = {dim: (score / max_score) * 100 for dim, score in scores.items()}

    def get_label(pct):
        if pct >= 70:
            return "Zeer sterk"
        elif pct >= 54:
            return "Sterk"
        elif pct >= 36:
            return "Neutraal"
        elif pct >= 18:
            return "Matig"
        else:
            return "Zwak"

    labels = {dim: get_label(pct) for dim, pct in percentages.items()}

    # Prepare results for saving
    results = {
        "completed": True,
        "scores": scores,
        "labels": labels,
        "percentages": percentages
    }

    dimension_labels = {
        "extraversie": "Extraversie",
        "altruïsme": "Altruïsme",
        "conciëntieusheid": "Conciëntieusheid",
        "neuroticisme": "Neuroticisme",
        "openheid": "Openheid"
    }

    colors = {
        "extraversie": "#4a6cf7",
        "altruïsme": "#27ae60",
        "conciëntieusheid": "#e67e22",
        "neuroticisme": "#e74c3c",
        "openheid": "#9b59b6"
    }

    # Build results data for template
    results_data = {}
    for dim, label in dimension_labels.items():
        results_data[dim] = {
            "label": label,
            "score": scores.get(dim, 0),
            "percentage": percentages.get(dim, 0),
            "level": labels.get(dim, "Onbekend"),
            "color": colors.get(dim, "#4a6cf7")
        }

    for dim in scores:
        results[dim] = results_data[dim]

    # Save to client with background tasks
    if client_id:
        save_questionnaire_results(client_id, "big_five", results, background_tasks)

    next_url = f"/phase/2.0{'?client_id=' + client_id if client_id else ''}"
    retry_url = f"/phase/1.1{'?client_id=' + client_id if client_id else ''}"

    content = render_template("result_1_1.html", {
        "request": request,
        "results": results_data,
        "next_url": next_url,
        "retry_url": retry_url
    })
    return HTMLResponse(content=content)


# ============================================================
# PHASE 2.0 - Career Anchors
# ============================================================
@app.get("/phase/2.0", response_class=HTMLResponse)
async def career_anchors_questionnaire(request: Request, client_id: str = None):
    """Career Anchors Questionnaire - Phase 2.0"""
    if not CAREER_ANCHOR_STATEMENTS:
        return HTMLResponse(
            content="<h1>Error: Career anchors not loaded</h1><p>Make sure fase_2_0.json exists in the data folder.</p>")

    if not client_id:
        client_id = request.query_params.get('client_id')

    grouped = {}
    for s in CAREER_ANCHOR_STATEMENTS:
        sid = s.get("id", 0)
        if sid not in grouped:
            grouped[sid] = []
        grouped[sid].append(s)

    anchor_names = {"Omhoog": "Omhoog komen", "Veilig": "Veilig voelen", "Vrij": "Vrij zijn",
                    "Balans": "Balans vinden", "Uitdaging": "Uitdaging zoeken"}

    descriptions = {}
    for k, v in CAREER_ANCHOR_DESCRIPTIONS.items():
        if k in anchor_names:
            descriptions[anchor_names[k]] = v

    content = render_template("phase_2_0.html", {
        "request": request,
        "client_id": client_id,
        "grouped": grouped,
        "descriptions": descriptions,
        "total_questions": len(CAREER_ANCHOR_STATEMENTS)
    })
    return HTMLResponse(content=content)


@app.post("/phase/2.0/submit")
async def submit_career_anchors(request: Request, background_tasks: BackgroundTasks):
    """Handle Career Anchors submission"""
    form_data = await request.form()
    client_id = form_data.get('client_id')

    scores = {"V": 0, "W": 0, "X": 0, "Y": 0, "Z": 0}

    for key, value in form_data.items():
        if key.startswith('s'):
            s_num = int(key[1:])
            for s in CAREER_ANCHOR_STATEMENTS:
                if s.get("id") == s_num:
                    anchor = s.get("anchor", "")
                    if anchor in scores:
                        scores[anchor] += int(value)
                    break

    anchor_names = {"V": "Omhoog komen", "W": "Veilig voelen", "X": "Vrij zijn",
                    "Y": "Balans vinden", "Z": "Uitdaging zoeken"}
    sorted_anchors = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    top1 = sorted_anchors[0] if len(sorted_anchors) > 0 else ("", 0)
    top2 = sorted_anchors[1] if len(sorted_anchors) > 1 else ("", 0)

    # Prepare results for saving
    results = {
        "completed": True,
        "scores": scores,
        "top1": top1[0],
        "top1_name": anchor_names.get(top1[0], ""),
        "top1_score": top1[1],
        "top2": top2[0],
        "top2_name": anchor_names.get(top2[0], ""),
        "top2_score": top2[1]
    }

    if client_id:
        save_questionnaire_results(client_id, "career_anchors", results, background_tasks)

    # Show results page
    colors = {"V": "#4a6cf7", "W": "#27ae60", "X": "#e67e22", "Y": "#9b59b6", "Z": "#e74c3c"}

    next_url = f"/phase/2.1{'?client_id=' + client_id if client_id else ''}"
    retry_url = f"/phase/2.0{'?client_id=' + client_id if client_id else ''}"

    content = render_template("result_2_0.html", {
        "request": request,
        "anchor_names": anchor_names,
        "sorted_anchors": sorted_anchors,
        "colors": colors,
        "top1": top1,
        "top2": top2,
        "next_url": next_url,
        "retry_url": retry_url
    })
    return HTMLResponse(content=content)


# ============================================================
# PHASE 2.1 - Career Clusters
# ============================================================
@app.get("/phase/2.1", response_class=HTMLResponse)
async def career_clusters_questionnaire(request: Request, client_id: str = None):
    """Career Clusters Questionnaire - Phase 2.1"""
    if not CAREER_CLUSTERS_DATA:
        return HTMLResponse(
            content="<h1>Error: Career clusters not loaded</h1><p>Make sure fase_2_1.json exists in the data folder.</p>")

    if not client_id:
        client_id = request.query_params.get('client_id')

    content = render_template("phase_2_1.html", {
        "request": request,
        "client_id": client_id,
        "clusters": CAREER_CLUSTERS_DATA
    })
    return HTMLResponse(content=content)


@app.post("/phase/2.1/submit")
async def submit_career_clusters(request: Request, background_tasks: BackgroundTasks):
    """Handle Career Clusters submission"""
    form_data = await request.form()
    client_id = form_data.get('client_id')

    cluster_scores = {}

    # Count all checked items per cluster
    for key, value in form_data.items():
        if key.startswith('c') and value == '1':
            parts = key.split('_')
            if len(parts) >= 2:
                cluster_id = parts[0][1:]  # Remove 'c' from the start
                if cluster_id not in cluster_scores:
                    cluster_scores[cluster_id] = 0
                cluster_scores[cluster_id] += 1

    # Sort clusters by score (highest first)
    sorted_clusters = sorted(cluster_scores.items(), key=lambda x: x[1], reverse=True)
    top1 = sorted_clusters[0] if len(sorted_clusters) > 0 else ("", 0)
    top2 = sorted_clusters[1] if len(sorted_clusters) > 1 else ("", 0)

    # Get cluster names
    cluster_names = {str(c.get("id")): c.get("name", f"Cluster {c.get('id')}") for c in CAREER_CLUSTERS_DATA}

    # Calculate max possible for each cluster based on displayed items
    max_possible = {}
    for cluster in CAREER_CLUSTERS_DATA:
        cid = str(cluster.get("id"))
        activities = len(cluster.get("activities", [])[:7])
        competencies = len(cluster.get("competencies", [])[:5])
        topics = len(cluster.get("educational_topics", [])[:5])
        max_possible[cid] = activities + competencies + topics

    # Prepare results for saving
    results = {
        "completed": True,
        "scores": cluster_scores,
        "top1": top1[0],
        "top1_name": cluster_names.get(top1[0], ""),
        "top1_score": top1[1],
        "top2": top2[0],
        "top2_name": cluster_names.get(top2[0], ""),
        "top2_score": top2[1]
    }

    # Save to client with background tasks
    if client_id:
        save_questionnaire_results(client_id, "career_clusters", results, background_tasks)

    next_url = f"/phase/2.2{'?client_id=' + client_id if client_id else ''}"
    retry_url = f"/phase/2.1{'?client_id=' + client_id if client_id else ''}"

    content = render_template("result_2_1.html", {
        "request": request,
        "cluster_names": cluster_names,
        "sorted_clusters": sorted_clusters,
        "top1": top1,
        "top2": top2,
        "max_possible": max_possible,
        "next_url": next_url,
        "retry_url": retry_url
    })
    return HTMLResponse(content=content)


# ============================================================
# PHASE 2.2 - Culture Analysis
# ============================================================
@app.get("/phase/2.2", response_class=HTMLResponse)
async def culture_questionnaire(request: Request, client_id: str = None):
    """Culture Analysis Questionnaire - Phase 2.2"""
    if not client_id:
        client_id = request.query_params.get('client_id')

    cultures = [
        {"id": 1, "name": "Mensgerichte cultuur", "subtitle": "(Clan-cultuur)",
         "description": "Deze cultuur is typerend voor zorginstellingen en service gerichte organisaties.",
         "questions": ["Er bestaan gemeenschappelijke waarden en doelstellingen.",
                       "Onderlinge samenhang (wij-gevoel). Participatieve instelling, teamwork.",
                       "Klanten worden als partners beschouwd.",
                       "Regels en procedures zijn ondergeschikt aan het gevoel een team te zijn."]},
        {"id": 2, "name": "Innovatieve cultuur", "subtitle": "(Adhocratie-cultuur)",
         "description": "Deze cultuur is typerend voor softwarebedrijven, luchtvaart, ruimtevaart e.d.",
         "questions": ["Snel reageren op snel veranderende omstandigheden.",
                       "Centraal staan innovatie en vernieuwing zoals de ontwikkeling van nieuwe diensten en producten.",
                       "Medewerkers worden gestimuleerd om creatief en flexibel te zijn.",
                       "De organisatie is flexibel en kan snel een nieuwe vorm aannemen."]},
        {"id": 3, "name": "Beheersgerichte cultuur", "subtitle": "(Hiërarchie-cultuur)",
         "description": "Deze cultuur is typerend voor overheids- of onderwijsinstellingen.",
         "questions": ["Het centraal staan van procedures en regels.", "De leidinggevenden coördineren en organiseren.",
                       "Langetermijndoelen zijn stabiliteit, efficiëntie en voorspelbaarheid.",
                       "De organisatie kenmerkt zich door trage besluitvormingsprocessen."]},
        {"id": 4, "name": "Resultaatgerichte cultuur", "subtitle": "(Markt-cultuur)",
         "description": "Deze cultuur is typerend voor organisaties die zich richten op externe transacties met belanghebbenden.",
         "questions": ["Centrale waarden zijn productiviteit, resultaten, winst en taakgerichtheid.",
                       "De externe positionering wordt benadrukt om de concurrentiepositie te versterken.",
                       "Er is een duidelijk doel aanwezig en er wordt gebruikgemaakt van een agressieve strategie.",
                       "Leidinggevenden zijn veeleisend en leggen de nadruk op de marktleider zijn."]}
    ]

    content = render_template("phase_2_2.html", {
        "request": request,
        "client_id": client_id,
        "cultures": cultures
    })
    return HTMLResponse(content=content)


@app.post("/phase/2.2/submit")
async def submit_culture(request: Request, background_tasks: BackgroundTasks):
    """Handle Culture Analysis submission"""
    form_data = await request.form()
    client_id = form_data.get('client_id')

    culture_scores = {}

    for key, value in form_data.items():
        if key.startswith('cq'):
            parts = key.split('_')
            if len(parts) >= 2:
                culture_id = parts[0][2:]
                if culture_id not in culture_scores:
                    culture_scores[culture_id] = 0
                culture_scores[culture_id] += int(value)

    sorted_cultures = sorted(culture_scores.items(), key=lambda x: x[1], reverse=True)
    top1 = sorted_cultures[0] if len(sorted_cultures) > 0 else ("", 0)
    top2 = sorted_cultures[1] if len(sorted_cultures) > 1 else ("", 0)

    culture_names = {"1": "Mensgerichte cultuur", "2": "Innovatieve cultuur",
                     "3": "Beheersgerichte cultuur", "4": "Resultaatgerichte cultuur"}
    culture_descriptions = {"1": "Deze cultuur is typerend voor zorginstellingen en service gerichte organisaties.",
                            "2": "Deze cultuur is typerend voor softwarebedrijven, luchtvaart, ruimtevaart e.d.",
                            "3": "Deze cultuur is typerend voor overheids- of onderwijsinstellingen.",
                            "4": "Deze cultuur is typerend voor organisaties die zich richten op externe transacties."}

    results = {
        "completed": True,
        "scores": culture_scores,
        "top1": top1[0],
        "top1_name": culture_names.get(top1[0], ""),
        "top1_score": top1[1],
        "top2": top2[0],
        "top2_name": culture_names.get(top2[0], ""),
        "top2_score": top2[1]
    }

    if client_id:
        save_questionnaire_results(client_id, "culture", results, background_tasks)

    next_url = f"/phase/2.3{'?client_id=' + client_id if client_id else ''}"
    retry_url = f"/phase/2.2{'?client_id=' + client_id if client_id else ''}"

    content = render_template("result_2_2.html", {
        "request": request,
        "culture_names": culture_names,
        "culture_descriptions": culture_descriptions,
        "sorted_cultures": sorted_cultures,
        "top1": top1,
        "top2": top2,
        "next_url": next_url,
        "retry_url": retry_url
    })
    return HTMLResponse(content=content)


# ============================================================
# PHASE 2.3 - JCM
# ============================================================
@app.get("/phase/2.3", response_class=HTMLResponse)
async def jcm_questionnaire(request: Request, client_id: str = None):
    """Job Characteristics Model Questionnaire - Phase 2.3"""
    if not client_id:
        client_id = request.query_params.get('client_id')

    jcm_components = [
        {"id": 1, "name": "Taakvaardigheid", "letter": "a",
         "description": "De mate waarin een baan verschillende activiteiten vereist, waarbij de werknemer een verscheidenheid aan vaardigheden en talenten moet ontwikkelen. Werknemers kunnen meer betekenis ervaren in banen die verschillende vaardigheden en capaciteiten vereisen dan wanneer de banen elementair en routinematig zijn.",
         "question": "Hoe belangrijk vind je het dat je werk bestaat uit verschillende soorten activiteiten waarbij je diverse vaardigheden en talenten kunt inzetten en/of ontwikkelen? En wat zou voor jou een ideale verhouding zijn tussen afwisselende taken en meer routinematige werkzaamheden?"},
        {"id": 2, "name": "Taakidentiteit", "letter": "b",
         "description": "De mate waarin de functie vereist dat de functiehouders een werkstuk identificeren en voltooien met een zichtbaar resultaat. Werknemers ervaren meer zingeving in een baan wanneer ze betrokken zijn bij het hele proces in plaats van alleen verantwoordelijk te zijn voor een deel van het werk.",
         "question": "Hoe belangrijk vind je het om in je werk betrokken te zijn bij een volledig proces of werkstuk, van begin tot eind, met een duidelijk zichtbaar resultaat? En in hoeverre zou je idealiter verantwoordelijk willen zijn voor het hele werkproces versus alleen een deel ervan?"},
        {"id": 3, "name": "Taakbetekenis", "letter": "c",
         "description": "De mate waarin de baan het leven van anderen beïnvloedt. De invloed kan zowel in de directe organisatie als in de externe omgeving zijn. Werknemers ervaren meer zingeving in een baan die het psychisch of fysiek welzijn van anderen aanzienlijk verbetert, dan een baan die een beperkt effect heeft op iemand anders.",
         "question": "Hoe belangrijk vind je het dat je werk een merkbare invloed heeft op het leven of welzijn van anderen, binnen en/of buiten de organisatie? En in welke mate zou je idealiter willen dat jouw werk impact heeft op anderen?"},
        {"id": 4, "name": "Autonomie", "letter": "d",
         "description": "De mate waarin de baan de werknemer aanzienlijke vrijheid, onafhankelijkheid en discretie biedt om het werk te plannen en de procedures in de baan te bepalen. Voor banen met een hoge mate van autonomie hangen de resultaten van het werk af van de eigen inspanningen, initiatieven en beslissingen van de werknemers, in plaats van in opdracht van een manager of een handleiding met werkprocedures. In dergelijke gevallen ervaren de werknemers een grotere persoonlijke verantwoordelijkheid voor hun eigen successen en mislukkingen op het werk.",
         "question": "Hoe belangrijk vind je het om in je werk zelf te kunnen bepalen hoe en wanneer je taken uitvoert, zonder dat alles strak voorgeschreven is? En hoeveel vrijheid zou je idealiter willen hebben in het plannen en uitvoeren van je werk?"},
        {"id": 5, "name": "Feedback", "letter": "e",
         "description": "De mate waarin de werknemer kennis heeft van resultaten. Dit is duidelijke, specifieke, gedetailleerde, bruikbare informatie over de effectiviteit van zijn of haar werkprestaties. Wanneer werknemers duidelijke, bruikbare informatie over hun werkprestaties ontvangen, hebben ze een betere algemene kennis van het effect van hun werkactiviteiten en welke specifieke acties ze moeten ondernemen (indien aanwezig) om hun productiviteit te verbeteren.",
         "question": "Hoe belangrijk vind je het om duidelijke en bruikbare feedback te krijgen over hoe goed je je werk doet? En in welke mate zou je idealiter op de hoogte willen zijn van het effect van je werk en van punten waarop je jezelf kunt verbeteren?"}
    ]

    content = render_template("phase_2_3.html", {
        "request": request,
        "client_id": client_id,
        "components": jcm_components
    })
    return HTMLResponse(content=content)


@app.post("/phase/2.3/submit")
async def submit_jcm(request: Request, background_tasks: BackgroundTasks):
    """Handle JCM questionnaire submission"""
    form_data = await request.form()
    client_id = form_data.get('client_id')

    answers = {}

    for key, value in form_data.items():
        if key.startswith('jcm_'):
            comp_id = key[4:]
            answers[comp_id] = value.strip()

    results = {
        "completed": True,
        "answers": answers
    }

    for comp_id, answer in answers.items():
        results[comp_id] = answer

    if client_id:
        save_questionnaire_results(client_id, "jcm", results, background_tasks)

    content = render_template("result_2_3.html", {
        "request": request,
        "answers": answers,
        "client_id": client_id
    })
    return HTMLResponse(content=content)


# ============================================================
# PROGNOSIS - Integrated via router from prognosis_report.py
# ============================================================
