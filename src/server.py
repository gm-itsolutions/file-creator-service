"""
File Creator Service
====================
Erstellt PPTX, DOCX, XLSX Dateien und gibt Download-Links zurück.
Für die Integration mit OpenWebUI/LLMs.
"""

import os
import uuid
import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from starlette.applications import Starlette
from starlette.responses import JSONResponse, HTMLResponse, FileResponse
from starlette.routing import Route
from starlette.middleware.cors import CORSMiddleware
import uvicorn

# Datei-Erstellungs-Libraries
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RgbColor
from docx import Document
from docx.shared import Pt as DocxPt, Inches as DocxInches
from openpyxl import Workbook

# Konfiguration
FILES_DIR = Path(os.getenv("FILES_DIR", "./files"))
BASE_URL = os.getenv("BASE_URL", "http://localhost:8002")
FILES_DIR.mkdir(exist_ok=True)

# Alte Dateien nach X Stunden löschen
FILE_RETENTION_HOURS = int(os.getenv("FILE_RETENTION_HOURS", "24"))


def cleanup_old_files():
    """Löscht Dateien die älter als FILE_RETENTION_HOURS sind."""
    cutoff = datetime.now() - timedelta(hours=FILE_RETENTION_HOURS)
    for f in FILES_DIR.iterdir():
        if f.is_file() and datetime.fromtimestamp(f.stat().st_mtime) < cutoff:
            f.unlink()


def generate_filename(prefix: str, extension: str) -> str:
    """Generiert einen einzigartigen Dateinamen."""
    unique_id = uuid.uuid4().hex[:8]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefix}_{timestamp}_{unique_id}.{extension}"


# ============ PPTX ERSTELLUNG ============

def create_pptx(title: str, slides: list[dict]) -> str:
    """
    Erstellt eine PowerPoint-Präsentation.
    
    slides = [
        {"title": "Slide Titel", "content": "Inhalt oder Bullet Points"},
        {"title": "Slide 2", "content": "• Punkt 1\n• Punkt 2\n• Punkt 3"},
    ]
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Titel-Slide
    title_slide_layout = prs.slide_layouts[6]  # Blank
    slide = prs.slides.add_slide(title_slide_layout)
    
    # Titel hinzufügen
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.333), Inches(1.5))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = title
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.alignment = 1  # Center
    
    # Content Slides
    for slide_data in slides:
        content_layout = prs.slide_layouts[6]  # Blank
        slide = prs.slides.add_slide(content_layout)
        
        # Slide Titel
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.333), Inches(1))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = slide_data.get("title", "")
        p.font.size = Pt(32)
        p.font.bold = True
        
        # Content
        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.5))
        tf = content_box.text_frame
        tf.word_wrap = True
        
        content = slide_data.get("content", "")
        lines = content.split("\n")
        
        for i, line in enumerate(lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = line.strip()
            p.font.size = Pt(18)
            p.space_after = Pt(12)
    
    # Speichern
    filename = generate_filename("presentation", "pptx")
    filepath = FILES_DIR / filename
    prs.save(str(filepath))
    
    return filename


# ============ DOCX ERSTELLUNG ============

def create_docx(title: str, sections: list[dict]) -> str:
    """
    Erstellt ein Word-Dokument.
    
    sections = [
        {"heading": "Überschrift", "content": "Paragraph text..."},
        {"heading": "Kapitel 2", "content": "Mehr text..."},
    ]
    """
    doc = Document()
    
    # Titel
    doc.add_heading(title, 0)
    
    # Sections
    for section in sections:
        if section.get("heading"):
            doc.add_heading(section["heading"], 1)
        if section.get("content"):
            doc.add_paragraph(section["content"])
    
    # Speichern
    filename = generate_filename("document", "docx")
    filepath = FILES_DIR / filename
    doc.save(str(filepath))
    
    return filename


# ============ XLSX ERSTELLUNG ============

def create_xlsx(title: str, sheets: list[dict]) -> str:
    """
    Erstellt eine Excel-Datei.
    
    sheets = [
        {
            "name": "Sheet Name",
            "headers": ["Spalte A", "Spalte B", "Spalte C"],
            "rows": [
                ["Wert 1", "Wert 2", "Wert 3"],
                ["Wert 4", "Wert 5", "Wert 6"],
            ]
        }
    ]
    """
    wb = Workbook()
    
    # Erstes Sheet entfernen wenn wir eigene haben
    if sheets:
        wb.remove(wb.active)
    
    for sheet_data in sheets:
        ws = wb.create_sheet(title=sheet_data.get("name", "Sheet"))
        
        # Headers
        headers = sheet_data.get("headers", [])
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = cell.font.copy(bold=True)
        
        # Rows
        rows = sheet_data.get("rows", [])
        for row_idx, row_data in enumerate(rows, 2):
            for col_idx, value in enumerate(row_data, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)
    
    # Speichern
    filename = generate_filename("spreadsheet", "xlsx")
    filepath = FILES_DIR / filename
    wb.save(str(filepath))
    
    return filename


# ============ HTTP ENDPOINTS ============

async def health(request):
    cleanup_old_files()  # Cleanup bei Health Check
    return JSONResponse({"status": "healthy", "files_dir": str(FILES_DIR)})


async def root(request):
    return JSONResponse({
        "name": "File Creator Service",
        "version": "1.0.0",
        "description": "Erstellt PPTX, DOCX, XLSX Dateien",
        "endpoints": {
            "/create/pptx": "PowerPoint erstellen",
            "/create/docx": "Word-Dokument erstellen",
            "/create/xlsx": "Excel-Datei erstellen",
            "/files/{filename}": "Datei herunterladen"
        }
    })


async def create_pptx_endpoint(request):
    """
    POST /create/pptx
    Body: {
        "title": "Präsentations-Titel",
        "slides": [
            {"title": "Slide 1", "content": "Inhalt..."},
            {"title": "Slide 2", "content": "• Punkt 1\n• Punkt 2"}
        ]
    }
    """
    try:
        data = await request.json()
        title = data.get("title", "Präsentation")
        slides = data.get("slides", [])
        
        if not slides:
            slides = [{"title": "Leere Folie", "content": "Inhalt hier einfügen"}]
        
        filename = create_pptx(title, slides)
        download_url = f"{BASE_URL}/files/{filename}"
        
        return JSONResponse({
            "success": True,
            "filename": filename,
            "download_url": download_url,
            "message": f"PowerPoint '{title}' wurde erstellt. Download: {download_url}"
        })
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, status_code=500)


async def create_docx_endpoint(request):
    """
    POST /create/docx
    Body: {
        "title": "Dokument-Titel",
        "sections": [
            {"heading": "Kapitel 1", "content": "Text..."},
            {"heading": "Kapitel 2", "content": "Mehr Text..."}
        ]
    }
    """
    try:
        data = await request.json()
        title = data.get("title", "Dokument")
        sections = data.get("sections", [])
        
        if not sections:
            sections = [{"heading": "", "content": "Inhalt hier einfügen"}]
        
        filename = create_docx(title, sections)
        download_url = f"{BASE_URL}/files/{filename}"
        
        return JSONResponse({
            "success": True,
            "filename": filename,
            "download_url": download_url,
            "message": f"Word-Dokument '{title}' wurde erstellt. Download: {download_url}"
        })
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, status_code=500)


async def create_xlsx_endpoint(request):
    """
    POST /create/xlsx
    Body: {
        "title": "Spreadsheet Name",
        "sheets": [
            {
                "name": "Daten",
                "headers": ["Name", "Wert"],
                "rows": [["Item 1", 100], ["Item 2", 200]]
            }
        ]
    }
    """
    try:
        data = await request.json()
        title = data.get("title", "Spreadsheet")
        sheets = data.get("sheets", [])
        
        if not sheets:
            sheets = [{"name": "Sheet1", "headers": ["A", "B"], "rows": []}]
        
        filename = create_xlsx(title, sheets)
        download_url = f"{BASE_URL}/files/{filename}"
        
        return JSONResponse({
            "success": True,
            "filename": filename,
            "download_url": download_url,
            "message": f"Excel-Datei '{title}' wurde erstellt. Download: {download_url}"
        })
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, status_code=500)


async def download_file(request):
    """Datei herunterladen."""
    filename = request.path_params["filename"]
    filepath = FILES_DIR / filename
    
    if not filepath.exists():
        return JSONResponse({"error": "Datei nicht gefunden"}, status_code=404)
    
    return FileResponse(
        filepath,
        filename=filename,
        media_type="application/octet-stream"
    )


async def list_files(request):
    """Liste aller verfügbaren Dateien."""
    files = []
    for f in FILES_DIR.iterdir():
        if f.is_file():
            files.append({
                "filename": f.name,
                "download_url": f"{BASE_URL}/files/{f.name}",
                "size_bytes": f.stat().st_size,
                "created": datetime.fromtimestamp(f.stat().st_mtime).isoformat()
            })
    return JSONResponse({"files": files})


# ============ OPENAPI SCHEMA ============

async def openapi_schema(request):
    base_url = str(request.base_url).rstrip("/").replace("http://", "https://")
    
    return JSONResponse({
        "openapi": "3.1.0",
        "info": {
            "title": "File Creator Service",
            "description": "Erstellt Office-Dateien (PPTX, DOCX, XLSX) und gibt Download-Links zurück. Nutze diese Tools um Dokumente für den Benutzer zu erstellen.",
            "version": "1.0.0"
        },
        "servers": [{"url": base_url}],
        "paths": {
            "/create/pptx": {
                "post": {
                    "operationId": "create_powerpoint",
                    "summary": "Erstelle eine PowerPoint-Präsentation",
                    "description": "Erstellt eine PPTX-Datei mit Titel und mehreren Folien. Gibt einen Download-Link zurück.",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {
                                    "type": "object",
                                    "required": ["title", "slides"],
                                    "properties": {
                                        "title": {
                                            "type": "string",
                                            "description": "Titel der Präsentation"
                                        },
                                        "slides": {
                                            "type": "array",
                                            "description": "Liste der Folien",
                                            "items": {
                                                "type": "object",
                                                "properties": {
                                                    "title": {"type": "string", "description": "Folientitel"},
                                                    "content": {"type": "string", "description": "Folieninhalt (kann Bullet Points mit • enthalten)"}
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    },
                    "responses": {
                        "200": {
                            "description": "PowerPoint erfolgreich erstellt",
                            "content": {
                                "application/json": {
                                    "schema": {
                                        "type": "object",
                                        "properties": {
                                            "success": {"type": "boolean"},
                                            "filename": {"type": "string"},
                                            "download_url": {"type": "string"},
                                            "message": {"type": "string"}
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            },
            "/create/docx": {
                "post": {
                    "operationId": "create_word_document",
                    "summary": "Erstelle ein Word-Dokument",
                    "description": "Erstellt eine DOCX-Datei mit Titel und Abschnitten. Gibt einen Download-Link zurück.",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {
                                    "type": "object",
                                    "required": ["title", "sections"],
                                    "properties": {
                                        "title": {
                                            "type": "string",
                                            "description": "Titel des Dokuments"
                                        },
                                        "sections": {
                                            "type": "array",
                                            "description": "Liste der Abschnitte",
                                            "items": {
                                                "type": "object",
                                                "properties": {
                                                    "heading": {"type": "string", "description": "Überschrift des Abschnitts"},
                                                    "content": {"type": "string", "description": "Textinhalt des Abschnitts"}
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    },
                    "responses": {
                        "200": {
                            "description": "Word-Dokument erfolgreich erstellt",
                            "content": {
                                "application/json": {
                                    "schema": {
                                        "type": "object",
                                        "properties": {
                                            "success": {"type": "boolean"},
                                            "filename": {"type": "string"},
                                            "download_url": {"type": "string"},
                                            "message": {"type": "string"}
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            },
            "/create/xlsx": {
                "post": {
                    "operationId": "create_excel_spreadsheet",
                    "summary": "Erstelle eine Excel-Tabelle",
                    "description": "Erstellt eine XLSX-Datei mit Tabellen. Gibt einen Download-Link zurück.",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {
                                    "type": "object",
                                    "required": ["title", "sheets"],
                                    "properties": {
                                        "title": {
                                            "type": "string",
                                            "description": "Name der Excel-Datei"
                                        },
                                        "sheets": {
                                            "type": "array",
                                            "description": "Liste der Arbeitsblätter",
                                            "items": {
                                                "type": "object",
                                                "properties": {
                                                    "name": {"type": "string", "description": "Name des Arbeitsblatts"},
                                                    "headers": {"type": "array", "items": {"type": "string"}, "description": "Spaltenüberschriften"},
                                                    "rows": {"type": "array", "items": {"type": "array"}, "description": "Datenzeilen"}
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    },
                    "responses": {
                        "200": {
                            "description": "Excel-Datei erfolgreich erstellt",
                            "content": {
                                "application/json": {
                                    "schema": {
                                        "type": "object",
                                        "properties": {
                                            "success": {"type": "boolean"},
                                            "filename": {"type": "string"},
                                            "download_url": {"type": "string"},
                                            "message": {"type": "string"}
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    })


async def docs(request):
    return HTMLResponse("""
    <!DOCTYPE html><html><head><title>File Creator Service</title>
    <link rel="stylesheet" href="https://unpkg.com/swagger-ui-dist@5/swagger-ui.css"></head>
    <body><div id="swagger-ui"></div>
    <script src="https://unpkg.com/swagger-ui-dist@5/swagger-ui-bundle.js"></script>
    <script>SwaggerUIBundle({url: "/openapi.json", dom_id: '#swagger-ui'});</script>
    </body></html>
    """)


# ============ APP ============

routes = [
    Route("/", root),
    Route("/health", health),
    Route("/openapi.json", openapi_schema),
    Route("/docs", docs),
    Route("/create/pptx", create_pptx_endpoint, methods=["POST"]),
    Route("/create/docx", create_docx_endpoint, methods=["POST"]),
    Route("/create/xlsx", create_xlsx_endpoint, methods=["POST"]),
    Route("/files", list_files),
    Route("/files/{filename}", download_file),
]

app = Starlette(routes=routes)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8002"))
    host = os.getenv("HOST", "0.0.0.0")
    print(f"🚀 File Creator Service auf http://{host}:{port}")
    print(f"📁 Dateien werden gespeichert in: {FILES_DIR}")
    uvicorn.run(app, host=host, port=port)
