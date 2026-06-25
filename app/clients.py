# app/clients.py
from fastapi import APIRouter, Request, Form, HTTPException, UploadFile, File, BackgroundTasks
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
import json
import uuid
from datetime import datetime
from typing import List, Dict, Any, Optional
import webbrowser
import os
from pathlib import Path
import shutil
import re

# Import from utils
from app.utils import (
    get_client, get_all_clients, save_client, delete_client,
    get_client_folder_name, get_client_excel_path, get_client_word_path
)

# Import from format_report
from app.format_report import create_client_excel, generate_word_report, update_client_excel

# Initialize router
router = APIRouter(prefix="/clients", tags=["clients"])

# Initialize router
router = APIRouter(prefix="/clients", tags=["clients"])

# Data directory
DATA_DIR = Path("data")
CLIENTS_DIR = DATA_DIR / "clients"
CLIENTS_DIR.mkdir(parents=True, exist_ok=True)


def render_template(content, title="Klantenbeheer"):
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


def save_questionnaire_results(
    client_id: str,
    phase: str,
    results: Dict[str, Any],
    background_tasks: BackgroundTasks = None
) -> bool:
    """Save questionnaire results to client data"""
    client = get_client(client_id)
    if not client:
        return False

    # Save the results immediately (fast operation)
    if phase == "big_five":
        client["big_five"] = results
    elif phase == "career_anchors":
        client["career_anchors"] = results
    elif phase == "career_clusters":
        client["career_clusters"] = results
    elif phase == "culture":
        client["culture"] = results
    elif phase == "jcm":
        client["jcm"] = results
    elif phase == "prognosis":
        client["prognosis"] = results

    # Check if all assessment questionnaires are completed (excluding prognosis)
    client["questionnaire_completed"] = all([
        client.get("big_five", {}).get("completed", False),
        client.get("career_anchors", {}).get("completed", False),
        client.get("career_clusters", {}).get("completed", False),
        client.get("culture", {}).get("completed", False),
        client.get("jcm", {}).get("completed", False)
    ])

    client["updated_at"] = datetime.now().isoformat()

    # STEP 1: Save client data immediately (fast - JSON only)
    if not save_client(client):
        return False

    # STEP 2: If all completed, generate reports in background
    if client["questionnaire_completed"]:
        if background_tasks:
            # Run in background - user gets immediate response
            background_tasks.add_task(update_client_excel, client_id)
            background_tasks.add_task(generate_word_report, client_id)
        else:
            # Fallback to synchronous (if no background tasks provided)
            update_client_excel(client_id)
            generate_word_report(client_id)

    return True


# ============================================================
# ROUTES
# ============================================================

@router.get("/", response_class=HTMLResponse)
async def client_list(request: Request):
    """Display all clients"""
    clients = get_all_clients()

    content = """
    <div class="container">
        <div class="header">
            <h1>👥 Klantenbeheer</h1>
            <div class="header-actions">
                <div class="search-box">
                    <span class="icon">🔍</span>
                    <input type="text" id="searchInput" placeholder="Zoeken op naam..." onkeyup="filterClients()">
                </div>
                <div style="display: flex; gap: 10px;">
                    <a href="/clients/open_folder" class="btn-secondary" style="padding: 10px 20px; text-decoration: none; display: inline-flex; align-items: center; gap: 8px;">
                        📁 Open Map
                    </a>
                    <a href="/clients/new" class="btn-primary">
                        <span>+</span> Nieuwe Klant
                    </a>
                </div>
            </div>
        </div>

        <div class="stats">
            <div class="stat-card">
                <div class="number">""" + str(len(clients)) + """</div>
                <div class="label">Totaal Klanten</div>
            </div>
            <div class="stat-card">
                <div class="number">""" + str(sum(1 for c in clients if c.get('questionnaire_completed', False))) + """</div>
                <div class="label">Vragenlijst Voltooid</div>
            </div>
            <div class="stat-card">
                <div class="number">""" + str(sum(1 for c in clients if c.get('excel_created', False))) + """</div>
                <div class="label">Excel Aangemaakt</div>
            </div>
        </div>

        <div class="client-grid" id="clientGrid">
    """

    if clients:
        for client in clients:
            initials = ''.join([word[0] for word in client.get('name', '?').split()[:2]]) or '?'
            status_badge = "✅ Voltooid" if client.get('questionnaire_completed', False) else "⏳ In Behandeling"
            status_color = "#4a6cf7" if client.get('questionnaire_completed', False) else "#f39c12"

            content += f"""
                <a href="/clients/{client.get('id')}" class="client-card" data-name="{client.get('name', '').lower()}">
                    <div class="card-header">
                        <div class="avatar">{initials.upper()}</div>
                        <div>
                            <div class="name">{client.get('name', 'Onbekend')}</div>
                            <div class="client-id">ID: {client.get('id', '')[:8]}</div>
                        </div>
                    </div>
                    <div class="details">
                        <div class="item"><span class="label">📧</span> {client.get('email', 'Geen email')}</div>
                        <div class="item"><span class="label">📱</span> {client.get('telefoonnummer', 'Geen telefoon')}</div>
                        <div class="item"><span class="label">🌍</span> {client.get('land_van_herkomst', 'Niet opgegeven')}</div>
                        <div class="item"><span class="label">📊</span> {client.get('questionnaire_completed', False) and '✅' or '⏳'}</div>
                    </div>
                    <div class="badge" style="background: {status_color}20; color: {status_color};">{status_badge}</div>
                    <div class="card-actions">
                        <span class="btn-secondary" onclick="event.stopPropagation();">Bekijk</span>
                    </div>
                </a>
            """
    else:
        content += """
                <div class="empty-state">
                    <div class="icon">👤</div>
                    <h2>Geen klanten gevonden</h2>
                    <p>Voeg je eerste klant toe om te beginnen met het loopbaanonderzoek.</p>
                    <a href="/clients/new" class="btn-primary">+ Nieuwe Klant Toevoegen</a>
                </div>
        """

    content += """
        </div>
    </div>

    <script>
        function filterClients() {
            const input = document.getElementById('searchInput');
            const filter = input.value.toLowerCase();
            const cards = document.querySelectorAll('.client-card');
            let visibleCount = 0;

            cards.forEach(card => {
                const name = card.getAttribute('data-name');
                if (name && name.includes(filter)) {
                    card.classList.remove('hidden');
                    visibleCount++;
                } else {
                    card.classList.add('hidden');
                }
            });

            const grid = document.getElementById('clientGrid');
            let emptyMsg = grid.querySelector('.no-results');

            if (visibleCount === 0 && cards.length > 0) {
                if (!emptyMsg) {
                    emptyMsg = document.createElement('div');
                    emptyMsg.className = 'empty-state no-results';
                    emptyMsg.style.gridColumn = '1 / -1';
                    emptyMsg.innerHTML = `
                        <div class="icon">🔍</div>
                        <h2>Geen resultaten</h2>
                        <p>Geen klanten gevonden met de zoekterm "${input.value}"</p>
                    `;
                    grid.appendChild(emptyMsg);
                }
            } else if (emptyMsg) {
                emptyMsg.remove();
            }
        }
    </script>
    """
    return HTMLResponse(content=render_template(content))


@router.get("/new", response_class=HTMLResponse)
async def new_client_form(request: Request):
    """Display form to create a new client"""
    content = """
    <div class="container">
        <div class="header">
            <h1>➕ Nieuwe Klant Toevoegen</h1>
            <a href="/clients/" class="btn-back">← Terug naar overzicht</a>
        </div>

        <div class="form-card">
            <form method="post" action="/clients/new" enctype="multipart/form-data">
                <div class="form-row">
                    <div class="form-group">
                        <label for="name">Naam *</label>
                        <input type="text" id="name" name="name" required placeholder="Volledige naam">
                    </div>
                    <div class="form-group">
                        <label for="geboortedatum">Geboortedatum</label>
                        <input type="date" id="geboortedatum" name="geboortedatum">
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label for="email">Email</label>
                        <input type="email" id="email" name="email" placeholder="voorbeeld@email.nl">
                    </div>
                    <div class="form-group">
                        <label for="telefoonnummer">Telefoonnummer</label>
                        <input type="tel" id="telefoonnummer" name="telefoonnummer" placeholder="06-12345678">
                    </div>
                </div>

                <div class="form-group">
                    <label for="adres">Adres</label>
                    <input type="text" id="adres" name="adres" placeholder="Straat + huisnummer, Postcode, Plaats">
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label for="opleidingen">Opleidingen</label>
                        <input type="text" id="opleidingen" name="opleidingen" placeholder="Bijv. HBO Bedrijfskunde">
                    </div>
                    <div class="form-group">
                        <label for="land_van_herkomst">Land van herkomst</label>
                        <input type="text" id="land_van_herkomst" name="land_van_herkomst" placeholder="Bijv. Nederland">
                    </div>
                </div>

                <div class="form-group">
                    <label for="leef_situatie">Leef situatie</label>
                    <select id="leef_situatie" name="leef_situatie">
                        <option value="">Selecteer...</option>
                        <option value="Alleenstaand">Alleenstaand</option>
                        <option value="Samenwonend">Samenwonend</option>
                        <option value="Getrouwd">Getrouwd</option>
                        <option value="Gescheiden">Gescheiden</option>
                        <option value="Weduwe/weduwnaar">Weduwe/weduwnaar</option>
                    </select>
                </div>

                <div class="form-group">
                    <label for="toelichting">Toelichting (optioneel)</label>
                    <textarea id="toelichting" name="toelichting" rows="4" placeholder="Extra informatie over de klant..."></textarea>
                </div>

                <div class="form-group">
                    <label for="excel_file">Excel Bestand (optioneel)</label>
                    <input type="file" id="excel_file" name="excel_file" accept=".xlsx,.xls">
                    <small style="color: #666; display: block; margin-top: 5px;">Upload een bestaand Excel bestand of laat leeg om een template te gebruiken.</small>
                </div>

                <div class="form-actions">
                    <button type="submit" class="btn-primary">➕ Toevoegen</button>
                    <a href="/clients/" class="btn-secondary">Annuleren</a>
                </div>
            </form>
        </div>
    </div>
    """
    return HTMLResponse(content=render_template(content, "Nieuwe Klant"))


@router.get("/open_folder")
async def open_clients_folder():
    """Open the clients data folder in Windows Explorer"""
    try:
        import os
        from pathlib import Path

        # Get absolute path to data/clients
        clients_path = Path("data/clients").absolute()

        # Create the folder if it doesn't exist
        clients_path.mkdir(parents=True, exist_ok=True)

        # Open folder in Windows Explorer
        os.startfile(str(clients_path))

        return RedirectResponse(url="/clients/", status_code=303)
    except Exception as e:
        print(f"Error opening folder: {e}")
        return RedirectResponse(url="/clients/", status_code=303)


@router.post("/new")
async def create_client(
        name: str = Form(...),
        geboortedatum: str = Form(""),
        email: str = Form(""),
        telefoonnummer: str = Form(""),
        adres: str = Form(""),
        opleidingen: str = Form(""),
        land_van_herkomst: str = Form(""),
        leef_situatie: str = Form(""),
        toelichting: str = Form(""),
        excel_file: Optional[UploadFile] = File(None)
):
    """Create a new client"""
    client_id = str(uuid.uuid4())[:8]
    now = datetime.now().isoformat()

    folder_name = get_client_folder_name(name)

    client_data = {
        "id": client_id,
        "name": name,
        "folder_name": folder_name,
        "geboortedatum": geboortedatum,
        "email": email,
        "telefoonnummer": telefoonnummer,
        "adres": adres,
        "opleidingen": opleidingen,
        "land_van_herkomst": land_van_herkomst,
        "leef_situatie": leef_situatie,
        "toelichting": toelichting,
        "created_at": now,
        "updated_at": now,
        "questionnaire_completed": False,
        "excel_created": False,
        "word_created": False,
        "big_five": {},
        "career_anchors": {},
        "career_clusters": {},
        "culture": {},
        "jcm": {},
        "prognosis": {}
    }

    if save_client(client_data):
        # Handle Excel file upload or create from template (using format_report)
        if excel_file and excel_file.filename:
            excel_path = get_client_excel_path(client_id)
            try:
                contents = await excel_file.read()
                with open(excel_path, 'wb') as f:
                    f.write(contents)
                client_data["excel_created"] = True
                save_client(client_data)
            except Exception as e:
                print(f"Error uploading Excel: {e}")
        else:
            # Create Excel from template using format_report
            create_client_excel(client_id, client_data)
            client_data["excel_created"] = True
            save_client(client_data)

        return RedirectResponse(url="/clients/", status_code=303)
    else:
        raise HTTPException(status_code=500, detail="Failed to create client")


@router.get("/{client_id}", response_class=HTMLResponse)
async def client_detail(request: Request, client_id: str):
    """Display a specific client's details"""
    client = get_client(client_id)
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    initials = ''.join([word[0] for word in client.get('name', '?').split()[:2]]) or '?'

    # Check if files exist
    excel_exists = get_client_excel_path(client_id).exists()
    word_exists = get_client_word_path(client_id).exists()

    content = f"""
    <div class="container">
        <div class="header">
            <h1>👤 Klant Detail</h1>
            <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                <a href="/clients/{client_id}/edit" class="btn-secondary">✏️ Bewerken</a>
                <a href="/clients/{client_id}/generate_report" class="btn-success">📄 Genereer Rapport</a>
                <a href="/clients/" class="btn-back">← Terug</a>
            </div>
        </div>

        <div class="profile-card">
            <div class="profile-header">
                <div class="profile-avatar">{initials.upper()}</div>
                <div>
                    <div class="profile-name">{client.get('name', 'Onbekend')}</div>
                    <div class="profile-meta">
                        Klant sinds: {client.get('created_at', '')[:10] if client.get('created_at') else 'N.v.t.'}
                        {' • Laatst bijgewerkt: ' + client.get('updated_at', '')[:10] if client.get('updated_at') and client.get('updated_at') != client.get('created_at') else ''}
                        {' • Map: ' + client.get('folder_name', client_id)}
                    </div>
                </div>
            </div>

            <div class="info-grid">
                <div class="info-item">
                    <div class="label">📧 Email</div>
                    <div class="value {'empty' if not client.get('email') else ''}">{client.get('email', 'Niet ingevuld')}</div>
                </div>
                <div class="info-item">
                    <div class="label">📱 Telefoonnummer</div>
                    <div class="value {'empty' if not client.get('telefoonnummer') else ''}">{client.get('telefoonnummer', 'Niet ingevuld')}</div>
                </div>
                <div class="info-item">
                    <div class="label">🎂 Geboortedatum</div>
                    <div class="value {'empty' if not client.get('geboortedatum') else ''}">{client.get('geboortedatum', 'Niet ingevuld')}</div>
                </div>
                <div class="info-item">
                    <div class="label">🌍 Land van herkomst</div>
                    <div class="value {'empty' if not client.get('land_van_herkomst') else ''}">{client.get('land_van_herkomst', 'Niet ingevuld')}</div>
                </div>
                <div class="info-item">
                    <div class="label">🏠 Leef situatie</div>
                    <div class="value {'empty' if not client.get('leef_situatie') else ''}">{client.get('leef_situatie', 'Niet ingevuld')}</div>
                </div>
                <div class="info-item">
                    <div class="label">📚 Opleidingen</div>
                    <div class="value {'empty' if not client.get('opleidingen') else ''}">{client.get('opleidingen', 'Niet ingevuld')}</div>
                </div>
                <div class="info-item" style="grid-column: 1 / -1;">
                    <div class="label">📍 Adres</div>
                    <div class="value {'empty' if not client.get('adres') else ''}">{client.get('adres', 'Niet ingevuld')}</div>
                </div>
                {f'<div class="info-item" style="grid-column: 1 / -1;"><div class="label">📝 Toelichting</div><div class="value">{client.get("toelichting", "")}</div></div>' if client.get('toelichting') else ''}
            </div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; margin-bottom: 30px;">
            <div class="stat-card">
                <div class="number">{'✅' if client.get('questionnaire_completed', False) else '⏳'}</div>
                <div class="label">Vragenlijst Status</div>
            </div>
            <div class="stat-card">
                <div class="number">{'✅' if excel_exists else '⏳'}</div>
                <div class="label">Excel Bestand</div>
            </div>
            <div class="stat-card">
                <div class="number">{'✅' if word_exists else '⏳'}</div>
                <div class="label">Word Rapport</div>
            </div>
        </div>

        <div class="tabs">
            <div class="tab-nav">
                <button class="tab-btn active" onclick="switchTab('assessment')">📋 Assessment Vragenlijst</button>
                <button class="tab-btn" onclick="switchTab('prognosis')">📊 Prognosis Vragenlijst</button>
                <button class="tab-btn" onclick="switchTab('report')">📄 Format Rapport</button>
                <button class="tab-btn" onclick="switchTab('files')">📁 Bestanden</button>
            </div>

            <div class="tab-content">
                <div id="tab-assessment" class="tab-pane active">
                    <h3>📋 Assessment Vragenlijst</h3>
                    <div class="coming-soon">
                        <div class="icon">🚀</div>
                        <h3>Start Assessment</h3>
                        <p>Begin het loopbaanonderzoek voor {client.get('name', 'deze klant')}.</p>
                        <br>
                        <div style="display: flex; gap: 10px; flex-wrap: wrap; justify-content: center;">
                            <a href="/phase/1.1?client_id={client_id}" class="btn-success">▶ Start Vragenlijst</a>
                            <a href="/phase/1.1?client_id={client_id}" class="btn-primary">📋 Fase 1.0</a>
                            <a href="/phase/2.0?client_id={client_id}" class="btn-primary">📋 Fase 2.0</a>
                            <a href="/phase/2.1?client_id={client_id}" class="btn-primary">📋 Fase 2.1</a>
                            <a href="/phase/2.2?client_id={client_id}" class="btn-primary">📋 Fase 2.2</a>
                            <a href="/phase/2.3?client_id={client_id}" class="btn-primary">📋 Fase 2.3</a>
                        </div>
                    </div>
                </div>

                <div id="tab-prognosis" class="tab-pane">
                    <h3>📊 Prognosis Vragenlijst</h3>
                    <div class="coming-soon">
                        <div class="icon">📈</div>
                        <h3>Integratie Prognose Model</h3>
                        <p>Deze vragenlijst brengt uw persoonlijke situatie, vaardigheden en arbeidsmarktpositie in kaart.</p>
                        <br>
                        <a href="/prognosis/?client_id={client_id}" class="btn-success">▶ Start Prognose Vragenlijst</a>
                        <br><br>
                        {'<a href="/prognosis/result?client_id=' + client_id + '" class="btn-primary">📊 Bekijk Resultaten</a>' if client.get('prognosis', {}).get('completed', False) else ''}
                    </div>
                </div>

                <div id="tab-report" class="tab-pane">
                    <h3>📄 Format Rapport</h3>
                    <div class="coming-soon">
                        <div class="icon">📄</div>
                        <h3>Rapport</h3>
                        <p>Het format rapport wordt gegenereerd na het invullen van de vragenlijsten.</p>
                        <br>
                        <a href="/clients/{client_id}/generate_report" class="btn-success">📄 Genereer Rapport</a>
                    </div>
                </div>

                <div id="tab-files" class="tab-pane">
                    <h3>📁 Bestanden</h3>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 20px;">
                        <div class="info-item" style="border-left-color: #27ae60;">
                            <div class="label">📊 Excel Bestand</div>
                            <div class="value">
                                {'✅ Bestaat' if excel_exists else '❌ Niet aangemaakt'}
                            </div>
                            <br>
                            <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                                <a href="/clients/{client_id}/download_excel" class="btn-primary" style="font-size: 14px; padding: 8px 16px;">⬇ Download</a>
                                <a href="/clients/{client_id}/upload_excel" class="btn-secondary" style="font-size: 14px; padding: 8px 16px;">📤 Upload Nieuw</a>
                            </div>
                        </div>
                        <div class="info-item" style="border-left-color: #4a6cf7;">
                            <div class="label">📄 Word Rapport</div>
                            <div class="value">
                                {'✅ Bestaat' if word_exists else '❌ Niet aangemaakt'}
                            </div>
                            <br>
                            <a href="/clients/{client_id}/download_word" class="btn-primary" style="font-size: 14px; padding: 8px 16px;">⬇ Download</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        function switchTab(tabName) {{
            document.querySelectorAll('.tab-pane').forEach(pane => {{
                pane.classList.remove('active');
            }});
            document.querySelectorAll('.tab-btn').forEach(btn => {{
                btn.classList.remove('active');
            }});
            document.getElementById('tab-' + tabName).classList.add('active');
            document.querySelectorAll('.tab-btn').forEach(btn => {{
                if (btn.textContent.includes(tabName.charAt(0).toUpperCase() + tabName.slice(1))) {{
                    btn.classList.add('active');
                }}
            }});
        }}
    </script>
    """
    return HTMLResponse(content=render_template(content, f"{client.get('name', 'Klant')} - Details"))


@router.get("/{client_id}/download_excel")
async def download_excel(client_id: str):
    """Download the client's Excel file"""
    client = get_client(client_id)
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    excel_path = get_client_excel_path(client_id)
    if not excel_path.exists():
        raise HTTPException(status_code=404, detail="Excel file not found")

    return FileResponse(
        path=excel_path,
        filename=f"{client.get('name', 'client')}_loopbaan_onderzoek.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@router.get("/{client_id}/upload_excel", response_class=HTMLResponse)
async def upload_excel_form(request: Request, client_id: str):
    """Display form to upload Excel file"""
    client = get_client(client_id)
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    content = f"""
    <div class="container">
        <div class="header">
            <h1>📤 Excel Bestand Uploaden</h1>
            <a href="/clients/{client_id}" class="btn-back">← Terug naar klant</a>
        </div>

        <div class="form-card">
            <form method="post" action="/clients/{client_id}/upload_excel" enctype="multipart/form-data">
                <div class="form-group">
                    <label for="excel_file">Selecteer Excel Bestand</label>
                    <input type="file" id="excel_file" name="excel_file" accept=".xlsx,.xls" required>
                    <small style="color: #666; display: block; margin-top: 5px;">Upload een .xlsx of .xls bestand.</small>
                </div>

                <div class="form-actions">
                    <button type="submit" class="btn-primary">📤 Uploaden</button>
                    <a href="/clients/{client_id}" class="btn-secondary">Annuleren</a>
                </div>
            </form>
        </div>
    </div>
    """
    return HTMLResponse(content=render_template(content, "Excel Uploaden"))


@router.post("/{client_id}/upload_excel")
async def upload_excel(client_id: str, excel_file: UploadFile = File(...)):
    """Upload a new Excel file for a client"""
    client = get_client(client_id)
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    if not excel_file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload .xlsx or .xls")

    excel_path = get_client_excel_path(client_id)

    try:
        contents = await excel_file.read()
        with open(excel_path, 'wb') as f:
            f.write(contents)

        client["excel_created"] = True
        save_client(client)

        return RedirectResponse(url=f"/clients/{client_id}", status_code=303)
    except Exception as e:
        print(f"Error uploading Excel: {e}")
        raise HTTPException(status_code=500, detail="Failed to upload Excel file")


@router.get("/{client_id}/download_word")
async def download_word(client_id: str):
    """Download the client's Word report"""
    client = get_client(client_id)
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    word_path = get_client_word_path(client_id)
    if not word_path.exists():
        raise HTTPException(status_code=404, detail="Word report not found")

    return FileResponse(
        path=word_path,
        filename=f"{client.get('name', 'client')}_loopbaan_rapport.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@router.get("/{client_id}/generate_report")
async def generate_report_route(client_id: str):
    """Generate the Word report for a client"""
    client = get_client(client_id)
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    # Generate report using format_report
    if generate_word_report(client_id):
        client["word_created"] = True
        save_client(client)
        return RedirectResponse(url=f"/clients/{client_id}", status_code=303)
    else:
        raise HTTPException(status_code=500, detail="Failed to generate report")


@router.get("/{client_id}/edit", response_class=HTMLResponse)
async def edit_client_form(request: Request, client_id: str):
    """Display form to edit a client"""
    client = get_client(client_id)
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    content = f"""
    <div class="container">
        <div class="header">
            <h1>✏️ Klant Bewerken</h1>
            <a href="/clients/{client_id}" class="btn-back">← Terug naar klant</a>
        </div>

        <div class="form-card">
            <form method="post" action="/clients/{client_id}/edit">
                <div class="form-row">
                    <div class="form-group">
                        <label for="name">Naam *</label>
                        <input type="text" id="name" name="name" required value="{client.get('name', '')}">
                    </div>
                    <div class="form-group">
                        <label for="geboortedatum">Geboortedatum</label>
                        <input type="date" id="geboortedatum" name="geboortedatum" value="{client.get('geboortedatum', '')}">
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label for="email">Email</label>
                        <input type="email" id="email" name="email" value="{client.get('email', '')}">
                    </div>
                    <div class="form-group">
                        <label for="telefoonnummer">Telefoonnummer</label>
                        <input type="tel" id="telefoonnummer" name="telefoonnummer" value="{client.get('telefoonnummer', '')}">
                    </div>
                </div>

                <div class="form-group">
                    <label for="adres">Adres</label>
                    <input type="text" id="adres" name="adres" value="{client.get('adres', '')}">
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label for="opleidingen">Opleidingen</label>
                        <input type="text" id="opleidingen" name="opleidingen" value="{client.get('opleidingen', '')}">
                    </div>
                    <div class="form-group">
                        <label for="land_van_herkomst">Land van herkomst</label>
                        <input type="text" id="land_van_herkomst" name="land_van_herkomst" value="{client.get('land_van_herkomst', '')}">
                    </div>
                </div>

                <div class="form-group">
                    <label for="leef_situatie">Leef situatie</label>
                    <select id="leef_situatie" name="leef_situatie">
                        <option value="">Selecteer...</option>
                        <option value="Alleenstaand" {'selected' if client.get('leef_situatie') == 'Alleenstaand' else ''}>Alleenstaand</option>
                        <option value="Samenwonend" {'selected' if client.get('leef_situatie') == 'Samenwonend' else ''}>Samenwonend</option>
                        <option value="Getrouwd" {'selected' if client.get('leef_situatie') == 'Getrouwd' else ''}>Getrouwd</option>
                        <option value="Gescheiden" {'selected' if client.get('leef_situatie') == 'Gescheiden' else ''}>Gescheiden</option>
                        <option value="Weduwe/weduwnaar" {'selected' if client.get('leef_situatie') == 'Weduwe/weduwnaar' else ''}>Weduwe/weduwnaar</option>
                    </select>
                </div>

                <div class="form-group">
                    <label for="toelichting">Toelichting (optioneel)</label>
                    <textarea id="toelichting" name="toelichting" rows="4">{client.get('toelichting', '')}</textarea>
                </div>

                <div class="form-actions">
                    <button type="submit" class="btn-primary">💾 Opslaan</button>
                    <a href="/clients/{client_id}" class="btn-secondary">Annuleren</a>
                </div>
            </form>
        </div>
    </div>
    """
    return HTMLResponse(content=render_template(content, "Klant Bewerken"))


@router.post("/{client_id}/edit")
async def update_client(
        client_id: str,
        name: str = Form(...),
        geboortedatum: str = Form(""),
        email: str = Form(""),
        telefoonnummer: str = Form(""),
        adres: str = Form(""),
        opleidingen: str = Form(""),
        land_van_herkomst: str = Form(""),
        leef_situatie: str = Form(""),
        toelichting: str = Form("")
):
    """Update an existing client"""
    existing_client = get_client(client_id)
    if not existing_client:
        raise HTTPException(status_code=404, detail="Client not found")

    client_data = {
        "id": client_id,
        "name": name,
        "folder_name": existing_client.get("folder_name", get_client_folder_name(name)),
        "geboortedatum": geboortedatum,
        "email": email,
        "telefoonnummer": telefoonnummer,
        "adres": adres,
        "opleidingen": opleidingen,
        "land_van_herkomst": land_van_herkomst,
        "leef_situatie": leef_situatie,
        "toelichting": toelichting,
        "created_at": existing_client.get("created_at", ""),
        "updated_at": datetime.now().isoformat(),
        "questionnaire_completed": existing_client.get("questionnaire_completed", False),
        "excel_created": existing_client.get("excel_created", False),
        "word_created": existing_client.get("word_created", False),
        "big_five": existing_client.get("big_five", {}),
        "career_anchors": existing_client.get("career_anchors", {}),
        "career_clusters": existing_client.get("career_clusters", {}),
        "culture": existing_client.get("culture", {}),
        "jcm": existing_client.get("jcm", {})
    }

    if save_client(client_data):
        # Update Excel file with new data using format_report
        update_client_excel(client_id)
        return RedirectResponse(url=f"/clients/{client_id}", status_code=303)
    else:
        raise HTTPException(status_code=500, detail="Failed to update client")


@router.post("/{client_id}/delete")
async def delete_client_route(client_id: str):
    """Delete a client"""
    if delete_client(client_id):
        return RedirectResponse(url="/clients/", status_code=303)
    else:
        raise HTTPException(status_code=404, detail="Client not found")