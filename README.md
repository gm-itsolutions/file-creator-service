# File Creator Service

Ein einfacher Service der **PPTX, DOCX und XLSX Dateien** erstellt und Download-Links zurückgibt. Perfekt für die Integration mit LLMs in OpenWebUI.

## Features

- 📊 **PowerPoint erstellen** - Präsentationen mit Titel und Folien
- 📝 **Word-Dokumente erstellen** - Dokumente mit Überschriften und Abschnitten
- 📈 **Excel-Tabellen erstellen** - Spreadsheets mit mehreren Arbeitsblättern
- 🔗 **Download-Links** - Funktionierende URLs statt `sandbox:/` Pfade
- 🧹 **Auto-Cleanup** - Alte Dateien werden nach 24h gelöscht

## API Endpoints

| Endpoint | Methode | Beschreibung |
|----------|---------|--------------|
| `/create/pptx` | POST | PowerPoint erstellen |
| `/create/docx` | POST | Word-Dokument erstellen |
| `/create/xlsx` | POST | Excel-Tabelle erstellen |
| `/files/{filename}` | GET | Datei herunterladen |
| `/files` | GET | Alle Dateien auflisten |
| `/docs` | GET | API-Dokumentation |

## Beispiel: PowerPoint erstellen

```bash
curl -X POST https://files.deine-domain.de/create/pptx \
  -H "Content-Type: application/json" \
  -d '{
    "title": "Meine Präsentation",
    "slides": [
      {"title": "Einleitung", "content": "Willkommen zu meiner Präsentation"},
      {"title": "Hauptteil", "content": "• Punkt 1\n• Punkt 2\n• Punkt 3"},
      {"title": "Fazit", "content": "Vielen Dank!"}
    ]
  }'
```

**Antwort:**
```json
{
  "success": true,
  "filename": "presentation_20240205_143052_a1b2c3d4.pptx",
  "download_url": "https://files.deine-domain.de/files/presentation_20240205_143052_a1b2c3d4.pptx",
  "message": "PowerPoint 'Meine Präsentation' wurde erstellt."
}
```

## Deployment auf Coolify

1. **Repository auf GitHub erstellen und Code pushen**

2. **In Coolify:**
   - Neues Projekt → Public Repository
   - Build Pack: Docker Compose
   - Domain zuweisen: z.B. `files.deine-domain.de`
   - Port: `8002`

3. **Environment Variable setzen:**
   ```
   BASE_URL=https://files.deine-domain.de
   ```

4. **Deploy!**

## OpenWebUI Integration

1. Gehe zu **Admin Panel** → **Settings** → **Tools** / **OpenAPI Servers**
2. Klicke **"+"** und füge hinzu:
   - **URL**: `https://files.deine-domain.de`
   - **Auth**: None
3. **Save**

Jetzt kann das LLM Dateien erstellen mit:
- `create_powerpoint`
- `create_word_document`
- `create_excel_spreadsheet`

## Lokal testen

```bash
# Dependencies installieren
pip install -r requirements.txt

# Server starten
python src/server.py

# Öffne http://localhost:8002/docs
```

## Environment Variables

| Variable | Default | Beschreibung |
|----------|---------|--------------|
| `PORT` | 8002 | Server-Port |
| `BASE_URL` | http://localhost:8002 | Öffentliche URL für Download-Links |
| `FILES_DIR` | /app/files | Verzeichnis für generierte Dateien |
| `FILE_RETENTION_HOURS` | 24 | Nach wie vielen Stunden Dateien gelöscht werden |
