import os
import re
import io
import json
import uuid
import sqlite3
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional

from dotenv import load_dotenv
from fastapi import FastAPI, Request, UploadFile, File, Form, HTTPException, Header, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pypdf import PdfReader

try:
    from pymongo import MongoClient
    PYMONGO_AVAILABLE = True
except ImportError:
    PYMONGO_AVAILABLE = False

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
STATIC_DIR.mkdir(exist_ok=True)

MONGODB_URI = os.getenv("MONGODB_URI", "").strip()
MONGODB_DB_NAME = os.getenv("MONGODB_DB_NAME", "albidan_corpus").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
DEFAULT_GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash").strip()
ADMIN_PIN = os.getenv("ADMIN_PIN", "").strip()

ALLOWED_MODELS = {
    "gemini-3.8-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.1-pro-preview",
}


# ==========================================
# 1. هياكل البيانات البحثية (Pydantic Schemas)
# ==========================================
class AILineAnalysis(BaseModel):
    original_line: str = Field(description="الشطر الأصلي كما أدخله الباحث")
    phonetic_diacritized: str = Field(description="الشطر مضبوطاً بالشكل الصوتي الحساني مع بيان السواكن")
    mutaharrikat_count: int = Field(description="العدد الدقيق للمتحركات في الشطر وفق العروض الحساني")
    estimated_meter: str = Field(description="اسم البحر الحساني")
    rhyme_ending: str = Field(description="الروي أو القافية الصوتية للشطر")


class GlossaryTerm(BaseModel):
    term: str = Field(description="المفردة الحسانية")
    arabic_meaning: str = Field(description="معناها بالعربية الفصحى وجذرها إن وُجد")
    english_meaning: str = Field(description="معناها المختصر بالإنجليزية")


class AIPoemAnalysis(BaseModel):
    poem_type: str = Field(description="الجنس الشعري: تبراع، كاف، طلعة، بت، أو نص حر")
    dominant_meter: str = Field(description="البحر العروضي السائد في القصيدة")
    dominant_rhyme: str = Field(description="بنية القافية والروي العام في القطعة")
    cultural_context: str = Field(description="تحليل موجز للغرض الشعري والسياق الثقافي الحساني")
    linguistic_notes: str = Field(description="ملاحظات لغوية وعروضية منهجية حول النص (الابتداء بالساكن، الظواهر الصوتية)")
    lines: List[AILineAnalysis]
    glossary: List[GlossaryTerm] = Field(description="المفردات الحسانية التراثية الواردة في النص وشرحها")
    english_translation: str = Field(description="ترجمة أدبية دقيقة للإنجليزية تحافظ على المعنى والروح")
    french_translation: str = Field(description="ترجمة أدبية دقيقة للفرنسية تحافظ على المعنى والروح")


class AnalyzeRequest(BaseModel):
    poem_text: str
    model_name: str = DEFAULT_GEMINI_MODEL
    use_ai: bool = True


class SavePoemRequest(BaseModel):
    title: str
    poet: str = "تراث شفهي / غير معروف"
    region: str = "الصحراء المغربية"
    poem_text: str
    analysis: Dict
    source_engine: str = "AI"


class UpdatePoemRequest(BaseModel):
    title: str
    poet: str
    region: str
    poem_text: str


class ExportAnalysisDocxRequest(BaseModel):
    title: str = "تقرير تحليل عروضي حساني"
    poet: str = "تراث شفهي / غير معروف"
    region: str = "الصحراء المغربية"
    poem_text: str
    analysis: Dict


class ContactMessageRequest(BaseModel):
    name: str
    contact: str = ""
    category: str = "تواصل بحثي"
    message: str


# ==========================================
# 2. أدوات إنشاء ملفات Word عربية احترافية (RTL)
# ==========================================
def set_paragraph_rtl(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)


def set_run_arabic(run, size_pt: float = 11.5, bold: bool = False, color_hex: Optional[str] = None, font_name: str = "Arial"):
    run.bold = bold
    run.font.size = Pt(size_pt)
    run.font.name = font_name
    if color_hex:
        r, g, b = int(color_hex[0:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)
        run.font.color.rgb = RGBColor(r, g, b)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rtl = OxmlElement("w:rtl")
    rtl.set(qn("w:val"), "1")
    rPr.append(rtl)


def set_table_rtl(table):
    tblPr = table._tbl.tblPr
    bidi = OxmlElement("w:bidiVisual")
    tblPr.append(bidi)


def set_cell_shading(cell, fill_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def add_rtl_para(doc: Document, text: str, size: float = 11.5, bold: bool = False, color: Optional[str] = None, space_after: float = 6):
    p = doc.add_paragraph()
    set_paragraph_rtl(p)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(str(text or ""))
    set_run_arabic(run, size_pt=size, bold=bold, color_hex=color)
    return p


def add_ltr_para(doc: Document, text: str, size: float = 10.5, bold: bool = False, color: Optional[str] = None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(str(text or ""))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    if color:
        r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        run.font.color.rgb = RGBColor(r, g, b)
    return p


def create_base_arabic_doc(doc_title: str, subtitle: str = "") -> Document:
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(0.8)
        sec.right_margin = Inches(0.8)
    add_rtl_para(doc, "مختبر البيظان للذكاء الاصطناعي — ALBIDAN AI LAB", size=10, bold=True, color="9A6B13", space_after=2)
    add_rtl_para(doc, doc_title, size=16, bold=True, color="1B365D", space_after=4)
    if subtitle:
        add_rtl_para(doc, subtitle, size=10, bold=False, color="555555", space_after=10)
    return doc


def append_poem_section_to_doc(doc: Document, title: str, poet: str, region: str, poem_text: str, analysis: Dict):
    add_rtl_para(doc, f"القصيدة: {title}", size=14, bold=True, color="1B365D", space_after=3)
    meta = f"الشاعر / الراوي: {poet}  |  المنطقة: {region}  |  الجنس الشعري: {analysis.get('poem_type', '—')}  |  البحر السائد: {analysis.get('dominant_meter', '—')}"
    add_rtl_para(doc, meta, size=10, color="555555", space_after=8)

    add_rtl_para(doc, "النص الشعري:", size=11.5, bold=True, color="9A6B13", space_after=3)
    for line in [l.strip() for l in poem_text.split("\n") if l.strip()]:
        add_rtl_para(doc, line, size=13, bold=False, color="111111", space_after=3)

    if analysis.get("cultural_context"):
        add_rtl_para(doc, "السياق الثقافي والغرض الشعري:", size=11, bold=True, color="1B365D", space_after=2)
        add_rtl_para(doc, analysis["cultural_context"], size=10.5, space_after=6)

    notes = analysis.get("linguistic_notes") or ""
    if notes:
        add_rtl_para(doc, "ملاحظات عروضية ولغوية:", size=11, bold=True, color="1B365D", space_after=2)
        add_rtl_para(doc, notes, size=10.5, space_after=6)

    lines_data = analysis.get("lines") or []
    if lines_data:
        add_rtl_para(doc, "جدول التقطيع الصوتي والعروضي:", size=11, bold=True, color="1B365D", space_after=4)
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        set_table_rtl(tbl)
        headers = ["الشطر الأصلي", "التشكيل الصوتي", "المتحركات", "البحر / الوزن", "الروي"]
        for i, h in enumerate(headers):
            cell = tbl.rows[0].cells[i]
            set_cell_shading(cell, "1B365D")
            p = cell.paragraphs[0]
            set_paragraph_rtl(p)
            set_run_arabic(p.add_run(h), size_pt=10, bold=True, color_hex="FFFFFF")
        for row_item in lines_data:
            cells = tbl.add_row().cells
            vals = [
                row_item.get("original_line", ""),
                row_item.get("phonetic_diacritized", ""),
                str(row_item.get("mutaharrikat_count", "")),
                row_item.get("estimated_meter", ""),
                row_item.get("rhyme_ending", ""),
            ]
            for i, v in enumerate(vals):
                p = cells[i].paragraphs[0]
                set_paragraph_rtl(p)
                set_run_arabic(p.add_run(v), size_pt=10)
        doc.add_paragraph()

    glossary = analysis.get("glossary") or []
    if glossary:
        add_rtl_para(doc, "المفردات الحسانية المشروحة:", size=11, bold=True, color="1B365D", space_after=4)
        gtbl = doc.add_table(rows=1, cols=3)
        gtbl.style = "Table Grid"
        set_table_rtl(gtbl)
        for i, h in enumerate(["المفردة", "المعنى بالعربية الفصحى", "English Meaning"]):
            cell = gtbl.rows[0].cells[i]
            set_cell_shading(cell, "EFE9DC")
            p = cell.paragraphs[0]
            set_paragraph_rtl(p)
            set_run_arabic(p.add_run(h), size_pt=10, bold=True, color_hex="1B365D")
        for g in glossary:
            cells = gtbl.add_row().cells
            for i, v in enumerate([g.get("term", ""), g.get("arabic_meaning", ""), g.get("english_meaning", "")]):
                p = cells[i].paragraphs[0]
                set_paragraph_rtl(p)
                set_run_arabic(p.add_run(v), size_pt=10)
        doc.add_paragraph()

    if analysis.get("english_translation"):
        add_ltr_para(doc, "English Literary Translation:", size=10.5, bold=True, color="1B365D")
        add_ltr_para(doc, analysis["english_translation"], size=10.5)
    if analysis.get("french_translation"):
        add_ltr_para(doc, "Traduction Littéraire (Français):", size=10.5, bold=True, color="1B365D")
        add_ltr_para(doc, analysis["french_translation"], size=10.5)


def doc_to_stream(doc: Document, filename: str) -> StreamingResponse:
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ==========================================
# 3. مدير قاعدة البيانات (MongoDB Atlas + SQLite احتياطي)
# ==========================================
class ArchiveManager:
    def __init__(self, mongo_uri: str, db_name: str, sqlite_path: str = "albidan_archive.db"):
        self.sqlite_path = str(BASE_DIR / sqlite_path)
        self.db_name = db_name
        self.active_mode = "sqlite"
        self.status_message = f"قاعدة SQLite المحلية ({sqlite_path})"
        self.mongo_db = None

        if mongo_uri and PYMONGO_AVAILABLE:
            try:
                client = MongoClient(mongo_uri, serverSelectionTimeoutMS=5000)
                client.admin.command("ping")
                self.mongo_db = client[db_name]
                self.active_mode = "mongodb"
                self.status_message = f"متصل سحابياً بـ MongoDB Atlas ({db_name})"
            except Exception as exc:
                self.status_message = f"تعذر الاتصال بـ MongoDB ({exc})؛ تم التحويل إلى SQLite."
                self._init_sqlite()
        else:
            self._init_sqlite()

    def _init_sqlite(self):
        with sqlite3.connect(self.sqlite_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS poems (
                    id TEXT PRIMARY KEY,
                    title TEXT NOT NULL,
                    poet TEXT,
                    region TEXT,
                    poem_text TEXT NOT NULL,
                    poem_type TEXT,
                    dominant_meter TEXT,
                    azawan_mode TEXT,
                    source_engine TEXT,
                    analysis_json TEXT NOT NULL,
                    music_json TEXT NOT NULL,
                    created_at TEXT NOT NULL
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS glossary (
                    term TEXT PRIMARY KEY,
                    arabic_meaning TEXT,
                    english_meaning TEXT,
                    source_poem_title TEXT,
                    updated_at TEXT NOT NULL
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS messages (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    contact TEXT,
                    category TEXT,
                    message TEXT NOT NULL,
                    created_at TEXT NOT NULL
                )
            """)
            conn.commit()

    def save_poem(self, payload: SavePoemRequest) -> str:
        record_id = str(uuid.uuid4())[:8]
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M")

        record = {
            "id": record_id,
            "title": payload.title.strip() or "قصيدة بدون عنوان",
            "poet": payload.poet.strip() or "تراث شفهي / غير معروف",
            "region": payload.region.strip() or "الصحراء المغربية",
            "poem_text": payload.poem_text.strip(),
            "poem_type": payload.analysis.get("poem_type", "غير محدد"),
            "dominant_meter": payload.analysis.get("dominant_meter", "غير محدد"),
            "source_engine": payload.source_engine,
            "analysis": payload.analysis,
            "created_at": now_iso,
        }
        glossary_items = payload.analysis.get("glossary", [])

        if self.active_mode == "mongodb":
            doc = dict(record)
            doc["_id"] = record_id
            self.mongo_db["poems"].insert_one(doc)
            for g in glossary_items:
                term = str(g.get("term", "")).strip()
                if term:
                    self.mongo_db["glossary"].update_one(
                        {"term": term},
                        {
                            "$set": {
                                "term": term,
                                "arabic_meaning": g.get("arabic_meaning", ""),
                                "english_meaning": g.get("english_meaning", ""),
                                "source_poem_title": record["title"],
                                "updated_at": now_iso,
                            }
                        },
                        upsert=True,
                    )
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.execute(
                    """
                    INSERT INTO poems (
                        id, title, poet, region, poem_text, poem_type, dominant_meter,
                        azawan_mode, source_engine, analysis_json, music_json, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        record["id"],
                        record["title"],
                        record["poet"],
                        record["region"],
                        record["poem_text"],
                        record["poem_type"],
                        record["dominant_meter"],
                        "",
                        record["source_engine"],
                        json.dumps(record["analysis"], ensure_ascii=False),
                        "{}",
                        record["created_at"],
                    ),
                )
                for g in glossary_items:
                    term = str(g.get("term", "")).strip()
                    if term:
                        conn.execute(
                            """
                            INSERT OR REPLACE INTO glossary (
                                term, arabic_meaning, english_meaning, source_poem_title, updated_at
                            ) VALUES (?, ?, ?, ?, ?)
                            """,
                            (
                                term,
                                g.get("arabic_meaning", ""),
                                g.get("english_meaning", ""),
                                record["title"],
                                now_iso,
                            ),
                        )
                conn.commit()
        return record_id

    def get_poem(self, record_id: str) -> Optional[Dict]:
        if self.active_mode == "mongodb":
            doc = self.mongo_db["poems"].find_one({"$or": [{"_id": record_id}, {"id": record_id}]})
            if not doc:
                return None
            if "id" not in doc and "_id" in doc:
                doc["id"] = str(doc["_id"])
            doc.pop("_id", None)
            return doc
        with sqlite3.connect(self.sqlite_path) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT * FROM poems WHERE id = ?", (record_id,)).fetchone()
            if not row:
                return None
            return {
                "id": row["id"],
                "title": row["title"],
                "poet": row["poet"],
                "region": row["region"],
                "poem_text": row["poem_text"],
                "poem_type": row["poem_type"],
                "dominant_meter": row["dominant_meter"],
                "source_engine": row["source_engine"],
                "analysis": json.loads(row["analysis_json"]),
                "created_at": row["created_at"],
            }

    def update_poem(self, record_id: str, payload: UpdatePoemRequest) -> bool:
        updates = {
            "title": payload.title.strip() or "قصيدة بدون عنوان",
            "poet": payload.poet.strip() or "تراث شفهي / غير معروف",
            "region": payload.region.strip() or "الصحراء المغربية",
            "poem_text": payload.poem_text.strip(),
        }
        if self.active_mode == "mongodb":
            res = self.mongo_db["poems"].update_one(
                {"$or": [{"_id": record_id}, {"id": record_id}]},
                {"$set": updates},
            )
            return res.matched_count > 0
        with sqlite3.connect(self.sqlite_path) as conn:
            cur = conn.execute(
                "UPDATE poems SET title = ?, poet = ?, region = ?, poem_text = ? WHERE id = ?",
                (updates["title"], updates["poet"], updates["region"], updates["poem_text"], record_id),
            )
            conn.commit()
            return cur.rowcount > 0

    def list_poems(self, q: str = "", poem_type: str = "الكل") -> List[Dict]:
        records: List[Dict] = []
        if self.active_mode == "mongodb":
            for doc in self.mongo_db["poems"].find().sort("created_at", -1):
                if "id" not in doc and "_id" in doc:
                    doc["id"] = str(doc["_id"])
                doc.pop("_id", None)
                records.append(doc)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.row_factory = sqlite3.Row
                rows = conn.execute("SELECT * FROM poems ORDER BY created_at DESC").fetchall()
                for row in rows:
                    records.append({
                        "id": row["id"],
                        "title": row["title"],
                        "poet": row["poet"],
                        "region": row["region"],
                        "poem_text": row["poem_text"],
                        "poem_type": row["poem_type"],
                        "dominant_meter": row["dominant_meter"],
                        "source_engine": row["source_engine"],
                        "analysis": json.loads(row["analysis_json"]),
                        "created_at": row["created_at"],
                    })

        if poem_type and poem_type != "الكل":
            records = [r for r in records if poem_type in str(r.get("poem_type", ""))]

        if q.strip():
            ql = q.strip().lower()
            records = [
                r for r in records
                if ql in str(r.get("title", "")).lower()
                or ql in str(r.get("poem_text", "")).lower()
                or ql in str(r.get("poet", "")).lower()
                or ql in str(r.get("region", "")).lower()
                or ql in str(r.get("dominant_meter", "")).lower()
            ]
        return records

    def delete_poem(self, record_id: str) -> bool:
        if self.active_mode == "mongodb":
            res = self.mongo_db["poems"].delete_one({"$or": [{"_id": record_id}, {"id": record_id}]})
            return res.deleted_count > 0
        with sqlite3.connect(self.sqlite_path) as conn:
            cur = conn.execute("DELETE FROM poems WHERE id = ?", (record_id,))
            conn.commit()
            return cur.rowcount > 0

    def list_glossary(self, q: str = "") -> List[Dict]:
        items: List[Dict] = []
        if self.active_mode == "mongodb":
            for doc in self.mongo_db["glossary"].find().sort("updated_at", -1):
                doc.pop("_id", None)
                items.append(doc)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.row_factory = sqlite3.Row
                items = [dict(r) for r in conn.execute("SELECT * FROM glossary ORDER BY updated_at DESC").fetchall()]

        if q.strip():
            ql = q.strip().lower()
            items = [
                g for g in items
                if ql in str(g.get("term", "")).lower()
                or ql in str(g.get("arabic_meaning", "")).lower()
                or ql in str(g.get("english_meaning", "")).lower()
            ]
        return items

    def delete_glossary_term(self, term: str) -> bool:
        if self.active_mode == "mongodb":
            res = self.mongo_db["glossary"].delete_one({"term": term})
            return res.deleted_count > 0
        with sqlite3.connect(self.sqlite_path) as conn:
            cur = conn.execute("DELETE FROM glossary WHERE term = ?", (term,))
            conn.commit()
            return cur.rowcount > 0

    def save_message(self, payload: ContactMessageRequest) -> str:
        msg_id = str(uuid.uuid4())[:8]
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M")
        doc = {
            "id": msg_id,
            "name": payload.name.strip(),
            "contact": payload.contact.strip(),
            "category": payload.category.strip() or "تواصل بحثي",
            "message": payload.message.strip(),
            "created_at": now_iso,
        }
        if self.active_mode == "mongodb":
            m = dict(doc)
            m["_id"] = msg_id
            self.mongo_db["messages"].insert_one(m)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.execute(
                    "INSERT INTO messages (id, name, contact, category, message, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                    (doc["id"], doc["name"], doc["contact"], doc["category"], doc["message"], doc["created_at"]),
                )
                conn.commit()
        return msg_id

    def list_messages(self) -> List[Dict]:
        items: List[Dict] = []
        if self.active_mode == "mongodb":
            for doc in self.mongo_db["messages"].find().sort("created_at", -1):
                if "id" not in doc and "_id" in doc:
                    doc["id"] = str(doc["_id"])
                doc.pop("_id", None)
                items.append(doc)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.row_factory = sqlite3.Row
                items = [dict(r) for r in conn.execute("SELECT * FROM messages ORDER BY created_at DESC").fetchall()]
        return items

    def delete_message(self, msg_id: str) -> bool:
        if self.active_mode == "mongodb":
            res = self.mongo_db["messages"].delete_one({"$or": [{"_id": msg_id}, {"id": msg_id}]})
            return res.deleted_count > 0
        with sqlite3.connect(self.sqlite_path) as conn:
            cur = conn.execute("DELETE FROM messages WHERE id = ?", (msg_id,))
            conn.commit()
            return cur.rowcount > 0

    def get_stats(self) -> Dict:
        if self.active_mode == "mongodb":
            poems_count = self.mongo_db["poems"].count_documents({})
            terms_count = self.mongo_db["glossary"].count_documents({})
            messages_count = self.mongo_db["messages"].count_documents({})
            cursor = self.mongo_db["poems"].find({}, {"analysis.lines": 1})
            total_lines = sum(len(doc.get("analysis", {}).get("lines", [])) for doc in cursor)
        else:
            poems = self.list_poems()
            glossary = self.list_glossary()
            messages = self.list_messages()
            poems_count = len(poems)
            terms_count = len(glossary)
            messages_count = len(messages)
            total_lines = sum(len(p.get("analysis", {}).get("lines", [])) for p in poems)

        return {
            "poems_count": poems_count,
            "lines_count": total_lines,
            "terms_count": terms_count,
            "messages_count": messages_count,
            "db_mode": self.active_mode,
            "db_status": self.status_message,
            "ai_ready": bool(GEMINI_API_KEY),
            "default_model": DEFAULT_GEMINI_MODEL,
        }


# ==========================================
# 4. المحلل الاحتياطي وخدمات Gemini (تحليل + صوت + OCR صور وPDF)
# ==========================================
class LocalHassaniyaEngine:
    short_vowels = set("َُِ")
    long_vowels = set("اويى")
    meters_map = {
        6: "السغير / لبير (6 متحركات)",
        7: "لبتيت الناقص (7 متحركات)",
        8: "لبتيت التام (8 متحركات)",
        9: "البت لكبير (9 متحركات)",
    }

    def analyze(self, poem_text: str) -> Dict:
        lines = [l.strip() for l in poem_text.strip().split("\n") if l.strip()]
        details = []
        for line in lines:
            norm = re.sub(r"[ݣڤ]", "گ", line)
            has_diacritics = any(ch in self.short_vowels for ch in norm)
            if has_diacritics:
                count = sum(1 for ch in norm if ch in self.short_vowels)
            else:
                words = norm.split()
                count = sum(
                    max(1, len([c for c in re.sub(r"^[ال]+", "", w) if c not in self.long_vowels]) - 1)
                    for w in words
                )
            clean = re.sub(r"[ًٌٍَُِّْ\s]", "", norm)
            rhyme = clean[-2:] if len(clean) >= 2 else clean
            details.append({
                "original_line": line,
                "phonetic_diacritized": "يتطلب تفعيل Gemini API للتشكيل الصوتي الدقيق",
                "mutaharrikat_count": count,
                "estimated_meter": self.meters_map.get(count, f"وزن تقديري ({count} متحرك)"),
                "rhyme_ending": rhyme,
            })

        poem_type = "نص شعري"
        if len(details) == 2:
            poem_type = "تبراع (Tebraa)"
        elif len(details) == 4:
            poem_type = "كاف (Gaf)"
        elif len(details) >= 6:
            poem_type = "طلعة (Tal'aa)"

        return {
            "poem_type": poem_type,
            "dominant_meter": details[0]["estimated_meter"] if details else "غير محدد",
            "dominant_rhyme": details[0]["rhyme_ending"] if details else "—",
            "cultural_context": "تحليل أولي بالمحرك المحلي. فعّل Gemini API في الخادم للحصول على التحليل الأكاديمي والترجمة.",
            "linguistic_notes": "تحليل محلي تقديري لعدد المتحركات.",
            "lines": details,
            "glossary": [],
            "english_translation": "Configure GEMINI_API_KEY on the server for literary translation.",
            "french_translation": "Configurez GEMINI_API_KEY sur le serveur pour la traduction littéraire.",
        }


def analyze_with_gemini(poem_text: str, model_name: str) -> Dict:
    client = genai.Client(api_key=GEMINI_API_KEY)
    system_instruction = """
    أنت باحث أكاديمي وعروضي متخصص في الأدب والشعر الشعبي الحساني (لغنى، الكاف، الطلعة، التبراع).
    قدّم تحليلاً علمياً دقيقاً ومنهجياً للنص الحساني:
    1. راعِ ظاهرة "الابتداء بالساكن" في اللهجة الحسانية عند التشكيل الصوتي وعدّ المتحركات.
    2. اعتمد عدّ المتحركات أساساً لتحديد البحر (السغير/لبير = 6، لبتيت الناقص = 7، لبتيت التام = 8، البت لكبير = 9+).
    3. ميّز بدقة بين التبراع، الكاف، الطلعة، ووضّح بنية القافية (الروي).
    4. قدّم ملاحظات لغوية ومنهجية تفيد الباحث اللساني، واستخرج المفردات الحسانية الأصيلة مع شرحها بالعربية الفصحى والإنجليزية.
    5. قدّم ترجمة أدبية دقيقة للإنجليزية والفرنسية.
    """
    response = client.models.generate_content(
        model=model_name,
        contents=f"حلّل هذا النص الشعري الحساني تحليلاً أكاديمياً وعروضياً كاملاً:\n\n{poem_text}",
        config=types.GenerateContentConfig(
            system_instruction=system_instruction,
            response_mime_type="application/json",
            response_schema=AIPoemAnalysis,
        ),
    )
    return AIPoemAnalysis.model_validate_json(response.text).model_dump()


def transcribe_with_gemini(audio_bytes: bytes, mime_type: str, model_name: str) -> str:
    client = genai.Client(api_key=GEMINI_API_KEY)
    prompt = """
    أنت خبير في اللهجة الحسانية والشعر الحساني. فرّغ التسجيل الصوتي نصياً بدقة:
    1. اكتب الكلمات كما تُنطق بالحسانية دون تحويلها للفصحى، مستخدماً حرف (گ) للقاف المعقودة.
    2. ضع كل شطر شعري (تافلويت) في سطر مستقل.
    3. أخرج الأبيات الشعرية فقط دون عناوين أو تعليقات.
    """
    response = client.models.generate_content(
        model=model_name,
        contents=[types.Part.from_bytes(data=audio_bytes, mime_type=mime_type), prompt],
    )
    return (response.text or "").strip()


def extract_poems_from_media_gemini(parts_list: List[types.Part], model_name: str) -> str:
    client = genai.Client(api_key=GEMINI_API_KEY)
    prompt = """
    أنت خبير في رقمنة المخطوطات والكتب والأشعار الحسانية (OCR).
    اقرأ النص الشعري الحساني الموجود في هذه الصور أو المستند واستخرجه بدقة تامة:
    1. حافظ على الألفاظ الحسانية الأصلية واستخدم حرف (گ) للقاف المعقودة الحسانية عند وجودها.
    2. ضع كل شطر شعري (تافلويت) في سطر مستقل ومرتب.
    3. إذا احتوى الملف على أكثر من قصيدة، افصل بين كل قصيدة وأخرى بسطر فارغ.
    4. أخرج النصوص الشعرية فقط بدون شروحات أو مقدمات إضافية.
    """
    response = client.models.generate_content(
        model=model_name,
        contents=[*parts_list, prompt],
    )
    return (response.text or "").strip()


def require_admin(x_admin_pin: Optional[str]):
    if ADMIN_PIN and (x_admin_pin or "").strip() != ADMIN_PIN:
        raise HTTPException(status_code=403, detail="رمز المشرف غير صحيح.")


# ==========================================
# 5. تطبيق FastAPI ونقاط الاتصال (API Routes)
# ==========================================
app = FastAPI(title="ALBIDAN AI Lab — Academic API", version="3.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

archive = ArchiveManager(MONGODB_URI, MONGODB_DB_NAME)
local_engine = LocalHassaniyaEngine()

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
def serve_home(request: Request):
    index_file = STATIC_DIR / "index.html"
    if not index_file.exists():
        raise HTTPException(status_code=404, detail="ملف الواجهة static/index.html غير موجود.")
    base_url = str(request.base_url).rstrip("/")
    if base_url.startswith("http://") and "localhost" not in base_url and "127.0.0.1" not in base_url:
        base_url = base_url.replace("http://", "https://", 1)
    return HTMLResponse(content=index_file.read_text(encoding="utf-8").replace("__BASE_URL__", base_url))


@app.get("/health")
def health_check():
    return {"status": "ok", "db_mode": archive.active_mode}


@app.get("/api/stats")
def api_stats():
    return archive.get_stats()


@app.post("/api/admin/verify")
def api_admin_verify(x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin)
    return {"authorized": True}


@app.post("/api/analyze")
def api_analyze(req: AnalyzeRequest):
    text = req.poem_text.strip()
    if not text:
        raise HTTPException(status_code=400, detail="يرجى إدخال بيت شعري واحد على الأقل.")

    model = req.model_name if req.model_name in ALLOWED_MODELS else DEFAULT_GEMINI_MODEL
    warning = None

    if req.use_ai and GEMINI_API_KEY:
        try:
            analysis = analyze_with_gemini(text, model)
            source_engine = f"AI ({model})"
        except Exception as exc:
            analysis = local_engine.analyze(text)
            source_engine = "Local (Fallback)"
            warning = f"تعذر الاتصال بـ Gemini ({exc})، تم العرض بالمحرك المحلي."
    else:
        analysis = local_engine.analyze(text)
        source_engine = "Local"

    return {
        "source_engine": source_engine,
        "warning": warning,
        "analysis": analysis,
    }


@app.post("/api/transcribe")
async def api_transcribe(
    file: UploadFile = File(...),
    model_name: str = Form(DEFAULT_GEMINI_MODEL),
):
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=400, detail="مفتاح GEMINI_API_KEY غير مضبوط في الخادم.")
    audio_bytes = await file.read()
    if len(audio_bytes) > 20 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="حجم الملف يتجاوز 20 ميجابايت.")

    ext = Path(file.filename or "audio.mp3").suffix.lower()
    mime_map = {".mp3": "audio/mp3", ".wav": "audio/wav", ".m4a": "audio/mp4", ".ogg": "audio/ogg", ".flac": "audio/flac"}
    mime_type = mime_map.get(ext, file.content_type or "audio/mp3")
    model = model_name if model_name in ALLOWED_MODELS else DEFAULT_GEMINI_MODEL

    try:
        text = transcribe_with_gemini(audio_bytes, mime_type, model)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"خطأ أثناء التفريغ الصوتي: {exc}")
    return {"text": text}


@app.post("/api/ocr")
async def api_ocr_images(
    files: List[UploadFile] = File(...),
    model_name: str = Form(DEFAULT_GEMINI_MODEL),
):
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=400, detail="يتطلب استخراج النص من الصور تفعيل مفتاح GEMINI_API_KEY.")
    if not files or len(files) > 6:
        raise HTTPException(status_code=400, detail="يمكن رفع من صورة واحدة إلى 6 صور في المرة الواحدة.")

    parts: List[types.Part] = []
    for f in files:
        raw = await f.read()
        if len(raw) > 15 * 1024 * 1024:
            raise HTTPException(status_code=400, detail=f"حجم الصورة {f.filename} يتجاوز 15 ميجابايت.")
        ext = Path(f.filename or "img.jpg").suffix.lower()
        mime = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}.get(ext, f.content_type or "image/jpeg")
        parts.append(types.Part.from_bytes(data=raw, mime_type=mime))

    model = model_name if model_name in ALLOWED_MODELS else DEFAULT_GEMINI_MODEL
    try:
        text = extract_poems_from_media_gemini(parts, model)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"تعذر قراءة الصورة: {exc}")
    return {"text": text}


@app.post("/api/import-file")
async def api_import_file(
    file: UploadFile = File(...),
    model_name: str = Form(DEFAULT_GEMINI_MODEL),
):
    raw = await file.read()
    if len(raw) > 20 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="حجم الملف يتجاوز 20 ميجابايت.")

    ext = Path(file.filename or "").suffix.lower()
    if ext == ".docx":
        try:
            doc = Document(io.BytesIO(raw))
            lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            return {"text": "\n".join(lines), "method": "docx"}
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"تعذر قراءة ملف Word: {exc}")

    elif ext == ".pdf":
        extracted = ""
        try:
            reader = PdfReader(io.BytesIO(raw))
            extracted = "\n".join((page.extract_text() or "").strip() for page in reader.pages).strip()
        except Exception:
            extracted = ""

        if len(extracted) >= 15:
            return {"text": extracted, "method": "pdf-text"}

        if GEMINI_API_KEY:
            model = model_name if model_name in ALLOWED_MODELS else DEFAULT_GEMINI_MODEL
            try:
                ocr_text = extract_poems_from_media_gemini(
                    [types.Part.from_bytes(data=raw, mime_type="application/pdf")], model
                )
                return {"text": ocr_text, "method": "pdf-ocr"}
            except Exception as exc:
                raise HTTPException(status_code=500, detail=f"تعذر قراءة ملف PDF المصوّر: {exc}")
        raise HTTPException(status_code=400, detail="ملف PDF مصوّر ولا يحتوي على نص رقمي مباشر.")

    elif ext == ".txt":
        return {"text": raw.decode("utf-8", errors="ignore").strip(), "method": "txt"}

    raise HTTPException(status_code=400, detail="الصيغ المدعومة للاستيراد هي: Word (.docx) و PDF (.pdf) و TXT.")


@app.get("/api/poems")
def api_list_poems(q: str = Query(""), poem_type: str = Query("الكل")):
    return {"items": archive.list_poems(q=q, poem_type=poem_type)}


@app.post("/api/poems")
def api_save_poem(req: SavePoemRequest):
    if not req.poem_text.strip():
        raise HTTPException(status_code=400, detail="نص القصيدة فارغ.")
    new_id = archive.save_poem(req)
    return {"id": new_id, "message": "تم التوثيق في المدونة وتحديث المعجم بنجاح."}


@app.put("/api/poems/{record_id}")
def api_update_poem(record_id: str, req: UpdatePoemRequest, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin)
    if not req.poem_text.strip():
        raise HTTPException(status_code=400, detail="نص القصيدة لا يمكن أن يكون فارغاً.")
    if not archive.update_poem(record_id, req):
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    return {"updated": True, "id": record_id}


@app.delete("/api/poems/{record_id}")
def api_delete_poem(record_id: str, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin)
    if not archive.delete_poem(record_id):
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    return {"deleted": True, "id": record_id}


@app.get("/api/glossary")
def api_list_glossary(q: str = Query("")):
    return {"items": archive.list_glossary(q=q)}


@app.delete("/api/glossary/{term}")
def api_delete_glossary(term: str, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin)
    if not archive.delete_glossary_term(term):
        raise HTTPException(status_code=404, detail="المفردة غير موجودة.")
    return {"deleted": True, "term": term}


@app.post("/api/messages")
def api_save_message(req: ContactMessageRequest):
    if not req.name.strip() or not req.message.strip():
        raise HTTPException(status_code=400, detail="يرجى كتابة الاسم والرسالة.")
    msg_id = archive.save_message(req)
    return {"id": msg_id, "message": "تم استلام مساهمتك بنجاح! شكراً لإثرائك مختبر البيظان."}


@app.get("/api/messages")
def api_list_messages(x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin)
    return {"items": archive.list_messages()}


@app.delete("/api/messages/{msg_id}")
def api_delete_message(msg_id: str, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin)
    if not archive.delete_message(msg_id):
        raise HTTPException(status_code=404, detail="الرسالة غير موجودة.")
    return {"deleted": True, "id": msg_id}


# ==========================================
# 6. مسارات التصدير الأكاديمي (Word .docx و JSONL)
# ==========================================
@app.get("/api/export/glossary.docx")
def api_export_glossary_docx(q: str = Query("")):
    items = archive.list_glossary(q=q)
    now_str = datetime.now().strftime("%Y-%m-%d")
    doc = create_base_arabic_doc(
        "المعجم الحساني المتنامي",
        f"تاريخ التصدير: {now_str}  |  عدد المفردات الموثقة: {len(items)}"
    )
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    set_table_rtl(tbl)
    headers = ["المفردة الحسانية", "المعنى بالعربية الفصحى", "English Meaning", "المصدر (القصيدة)"]
    for i, h in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        set_cell_shading(cell, "1B365D")
        p = cell.paragraphs[0]
        set_paragraph_rtl(p)
        set_run_arabic(p.add_run(h), size_pt=10.5, bold=True, color_hex="FFFFFF")
    for g in items:
        cells = tbl.add_row().cells
        vals = [g.get("term", ""), g.get("arabic_meaning", ""), g.get("english_meaning", ""), g.get("source_poem_title", "")]
        for i, v in enumerate(vals):
            p = cells[i].paragraphs[0]
            set_paragraph_rtl(p)
            set_run_arabic(p.add_run(v), size_pt=10.5, bold=(i == 0))
    return doc_to_stream(doc, "albidan_hassaniya_glossary.docx")


@app.post("/api/export/analysis.docx")
def api_export_analysis_docx(req: ExportAnalysisDocxRequest):
    now_str = datetime.now().strftime("%Y-%m-%d")
    doc = create_base_arabic_doc("تقرير التحليل العروضي واللساني", f"تاريخ الإصدار: {now_str}")
    append_poem_section_to_doc(doc, req.title, req.poet, req.region, req.poem_text, req.analysis)
    return doc_to_stream(doc, "albidan_poem_analysis.docx")


@app.get("/api/export/poem/{record_id}.docx")
def api_export_single_poem_docx(record_id: str):
    rec = archive.get_poem(record_id)
    if not rec:
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    doc = create_base_arabic_doc("وثيقة شعرية حسانية محقّقة", f"معرّف التوثيق: {rec['id']}  |  تاريخ الأرشفة: {rec.get('created_at', '')}")
    append_poem_section_to_doc(doc, rec["title"], rec["poet"], rec["region"], rec["poem_text"], rec.get("analysis", {}))
    return doc_to_stream(doc, f"albidan_poem_{rec['id']}.docx")


@app.get("/api/export/poems.docx")
def api_export_all_poems_docx(q: str = Query(""), poem_type: str = Query("الكل")):
    poems = archive.list_poems(q=q, poem_type=poem_type)
    now_str = datetime.now().strftime("%Y-%m-%d")
    doc = create_base_arabic_doc(
        "المدونة الحسانية المفتوحة — ديوان النصوص الموثقة",
        f"تاريخ التصدير: {now_str}  |  عدد القصائد: {len(poems)}"
    )
    for idx, rec in enumerate(poems):
        append_poem_section_to_doc(doc, rec["title"], rec["poet"], rec["region"], rec["poem_text"], rec.get("analysis", {}))
        if idx < len(poems) - 1:
            doc.add_page_break()
    return doc_to_stream(doc, "albidan_hassaniya_anthology.docx")


@app.get("/api/export/jsonl")
def api_export_jsonl():
    poems = archive.list_poems()
    content = "\n".join(json.dumps(p, ensure_ascii=False) for p in poems)
    return StreamingResponse(
        io.BytesIO(content.encode("utf-8")),
        media_type="application/jsonl",
        headers={"Content-Disposition": "attachment; filename=albidan_hassaniya_corpus.jsonl"},
    )
