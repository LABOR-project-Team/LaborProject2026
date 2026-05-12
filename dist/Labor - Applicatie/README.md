# LABOR Arbeidsmarktintegratie - Career Assessment Platform

## Overzicht

LABOR is een lokale Windows-applicatie voor carrièrebeoordeling en cliëntbeheer. De applicatie biedt vijf geïntegreerde assessmentfases, een aanvullend prognosemodel en lokale Excel-exporten. De eindgebruiker hoeft geen Python-omgeving te installeren: de applicatie wordt als gecompileerde build geleverd.

## Belangrijkste functionaliteit

- Beheer van meerdere cliëntprofielen
- Zoeken op cliëntnaam
- Nieuwe cliënten aanmaken met persoonlijke gegevens
- Cliëntgegevens bewerken en cliënten verwijderen
- Assessmentvragenlijsten starten, pauzeren en hervatten
- Prognosevragenlijst per cliënt
- Automatische opslag van voortgang in JSON
- Professionele export naar Excel
- Offline werking zonder internetverbinding

## Hoe het programma werkt

### Cliëntbeheer

Bij het openen van de applicatie zie je een overzicht van alle cliënten. Je kunt:
- een nieuwe cliënt toevoegen
- een cliënt selecteren om het dashboard te openen
- cliëntgegevens bewerken
- een cliënt verwijderen

### Dashboard per cliënt

Het dashboard toont:
- de cliëntnaam en persoonlijke informatie
- knoppen om de assessmentvragenlijst en prognosevragenlijst te starten
- de mogelijkheid om onvoltooide sessies te hervatten

### Assessmentflow

De applicatie bevat deze fases:
1. Fase 1.1: Big Five persoonlijkheidsassessment
2. Fase 2.0: Career Anchors (loopbaanwaarden)
3. Fase 2.1: Career Interests (interesseclusters)
4. Fase 2.2: Cultuurfit (organisatiecultuurvoorkeuren)
5. Fase 2.3: Job Characteristics (taakkenmerken)

Antwoorden worden automatisch opgeslagen tijdens het invullen. Zodra een fase is afgerond, wordt een Excel-bestand gegenereerd in de clientmap.

### Prognosemodel

Naast het reguliere assessment is er een prognosevragenlijst. Deze brengt de persoonlijke situatie, vaardigheden en arbeidsmarktpositie in kaart.

## Dataopslag

- Elke cliënt heeft een eigen map in `clients/`
- Cliëntgegevens worden opgeslagen in `info.json`
- Onvoltooide sessies worden opgeslagen zodat je kunt hervatten
- Excel-rapporten worden per cliënt opgeslagen in hun map

## Gebruik zonder Python

De applicatie wordt als gecompileerde Windows-build geleverd. Eindgebruikers hoeven geen Python of dependencies te installeren. Open de gecompileerde applicatie en ga direct aan de slag.

## Voor ontwikkelaars

Wil je de broncode aanpassen of lokaal draaien? Dan heb je wel een Python-omgeving nodig.

1. Installeer Python 3.x
2. Installeer dependencies met:
   ```bash
   pip install -r requirements.txt
   ```
3. Start de applicatie met:
   ```bash
   python app.py
   ```

## Technische architectuur

- Python met Tkinter voor de gebruikersinterface
- JSON voor lokale opslag van cliëntdata en sessies
- Excel-export via OpenPyXL
- Modulaire indeling in `phases/`, `ui/` en `utils/`

## Systeemvereisten

- Windows 7 of hoger
- Minimaal 500 MB vrije schijfruimte
- Microsoft Excel of een compatibele spreadsheet-app voor het openen van de exportbestanden
