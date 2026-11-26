import os
import re
import io
import math
import hmac
import time
import json
import uuid
import html
import logging
import sqlite3
import tempfile
from collections import defaultdict, Counter
from pathlib import Path
from datetime import datetime
from urllib.parse import quote
from typing import Any, Dict, Iterator, List, Literal, Optional, Tuple

from dotenv import load_dotenv
from fastapi import FastAPI, Request, UploadFile, File, Form, HTTPException, Header, Query
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, PlainTextResponse, Response, StreamingResponse
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
    from pymongo import MongoClient, UpdateOne, ASCENDING, DESCENDING
    PYMONGO_AVAILABLE = True
except ImportError:
    PYMONGO_AVAILABLE = False

load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("albidan")

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
STATIC_DIR.mkdir(exist_ok=True)

# =============================================================================
# 0. إعدادات البيئة والأمان والثوابت العلمية
# =============================================================================
APP_VERSION = "5.0.0"

MONGODB_URI = os.getenv("MONGODB_URI", "").strip()
MONGODB_DB_NAME = os.getenv("MONGODB_DB_NAME", "albidan_corpus").strip()
RAG_COLLECTION_NAME = os.getenv("RAG_COLLECTION_NAME", "hassaniya_rag_chunks").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
DEFAULT_GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip() or "gemini-3.5-flash-lite"
ADMIN_PIN = os.getenv("ADMIN_PIN", "").strip()

PROTECT_PUBLIC_WRITES = os.getenv("PROTECT_PUBLIC_WRITES", "true").strip().lower() == "true"
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "https://albidan.pro").strip().rstrip("/")

ALLOWED_ORIGINS = [
    origin.strip().rstrip("/")
    for origin in os.getenv(
        "ALLOWED_ORIGINS",
        "https://albidan.pro,https://www.albidan.pro",
    ).split(",")
    if origin.strip()
]

try:
    TRUSTED_PROXY_HOPS = max(0, int(os.getenv("TRUSTED_PROXY_HOPS", "1").strip()))
except ValueError:
    TRUSTED_PROXY_HOPS = 1

ALLOWED_MODELS = {
    "gemini-3.5-flash-lite",
    "gemini-3.8-flash",
    "gemini-3.1-pro-preview",
}
FALLBACK_MODELS = [
    m.strip()
    for m in os.getenv("GEMINI_FALLBACK_MODELS", "gemini-3.5-flash-lite,gemini-3.8-flash").split(",")
    if m.strip()
]

AUDIO_MIME_MAP = {
    ".mp3": "audio/mp3", ".wav": "audio/wav", ".m4a": "audio/mp4", ".mp4": "audio/mp4",
    ".aac": "audio/aac", ".ogg": "audio/ogg", ".oga": "audio/ogg", ".opus": "audio/ogg",
    ".webm": "audio/webm", ".flac": "audio/flac",
}
IMAGE_MIME_MAP = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}
FILES_API_THRESHOLD = 2 * 1024 * 1024
MAX_ANALYSIS_JSON_BYTES = 250_000

DEFAULT_POET_LABEL = "تراث شفهي / غير معروف"
DEFAULT_REGION_LABEL = "الصحراء المغربية"

VerificationState = Literal["machine_generated", "reviewed", "verified"]
VERIFICATION_LABELS: Dict[str, str] = {
    "machine_generated": "مولّد آلياً (غير مراجَع)",
    "reviewed": "قيد المراجعة البشرية",
    "verified": "موثّق ومحكّم",
}


# =============================================================================
# 1. التطبيع اللغوي الحساني وتنظيف العناوين وأسماء الشعراء + أدوات الأمان
# =============================================================================
ARABIC_DIACRITICS_RE = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u0640]")
POET_EMBED_PATTERNS = [
    re.compile(r"\s*[—–\-|•]\s*(?:الشاعر\s*/\s*الراوي|الشاعر|الراوي|قائلها|نظم)\s*[:：]\s*(.+)$"),
    re.compile(r"\s*\((?:الشاعر\s*/\s*الراوي|الشاعر|الراوي)\s*[:：]?\s*([^)]+)\)\s*$"),
    re.compile(r"\s+للشاعر\s+(.+)$"),
]
REDUNDANT_POET_PREFIX_RE = re.compile(r"^(?:الشاعر\s*/\s*الراوي|الشاعر|الراوي)\s*[:：]?\s*")


def normalize_hassaniya(text: str, unify_gaf_to_kaf: bool = False) -> str:
    """تطبيع النص الحساني مع إبقاء دمج ق/گ/ك خياراً مرناً منفصلاً لا يخلط الفهرسة الافتراضية."""
    if not text:
        return ""
    t = ARABIC_DIACRITICS_RE.sub("", str(text))
    replacements = {
        "أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا",
        "ى": "ي", "ة": "ه",
        "ݣ": "گ", "ڭ": "گ", "ڨ": "گ",
    }
    for src, dst in replacements.items():
        t = t.replace(src, dst)
    if unify_gaf_to_kaf:
        t = t.replace("ڤ", "گ").replace("گ", "ك").replace("ق", "ك")
    return re.sub(r"\s+", " ", t).strip().lower()


def is_default_poet(poet: str) -> bool:
    p = (poet or "").strip()
    return not p or p in {DEFAULT_POET_LABEL, "غير معروف", "تراث شفهي", "مجهول", "—", "-"}


def clean_poet_name(poet: str) -> str:
    p = (poet or "").strip()
    if not p:
        return DEFAULT_POET_LABEL
    for _ in range(3):
        cleaned = REDUNDANT_POET_PREFIX_RE.sub("", p).strip()
        if cleaned == p:
            break
        p = cleaned
    p = re.sub(r"^[:：\-—–\s]+|[:：\-—–\s]+$", "", p).strip()
    return p or DEFAULT_POET_LABEL


def clean_title_and_poet(title: str, poet: str) -> Tuple[str, str]:
    raw_title = re.sub(r"\s+", " ", (title or "").strip())
    extracted_poet = ""
    for pat in POET_EMBED_PATTERNS:
        m = pat.search(raw_title)
        if m:
            extracted_poet = clean_poet_name(m.group(1))
            raw_title = pat.sub("", raw_title).strip(" -—–|:")
            break
    final_poet = clean_poet_name(poet)
    if is_default_poet(final_poet) and extracted_poet and not is_default_poet(extracted_poet):
        final_poet = extracted_poet
    return (raw_title or "قصيدة بدون عنوان"), final_poet


def build_search_fields(title: str, poet: str, region: str, poem_text: str, dominant_meter: str) -> Tuple[str, str]:
    joined = f"{title} {poet} {region} {dominant_meter} {poem_text}"
    return normalize_hassaniya(joined, unify_gaf_to_kaf=False), normalize_hassaniya(joined, unify_gaf_to_kaf=True)


def get_base_url() -> str:
    return PUBLIC_BASE_URL


def safe_json_ld(data: dict) -> str:
    return (
        json.dumps(data, ensure_ascii=False)
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("&", "\\u0026")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )


def new_record_id(length: int = 12) -> str:
    return uuid.uuid4().hex[:length]


# =============================================================================
# 2. حماية الطلبات (Rate Limiting) وقفل المشرف
# =============================================================================
class SimpleRateLimiter:
    def __init__(self, max_tracked_keys: int = 5000):
        self._hits: Dict[str, List[float]] = defaultdict(list)
        self._admin_failures: Dict[str, List[float]] = defaultdict(list)
        self._max_tracked_keys = max_tracked_keys

    def _client_ip(self, request: Request) -> str:
        fwd = request.headers.get("x-forwarded-for", "")
        if fwd and TRUSTED_PROXY_HOPS > 0:
            parts = [p.strip() for p in fwd.split(",") if p.strip()]
            if parts:
                idx = max(0, len(parts) - TRUSTED_PROXY_HOPS)
                return parts[idx]
        return request.client.host if request.client else "unknown"

    def _prune_if_needed(self):
        if len(self._hits) > self._max_tracked_keys:
            now = time.time()
            stale = [k for k, ts in self._hits.items() if not ts or ts[-1] < now - 3600]
            for k in stale:
                self._hits.pop(k, None)

    def check_limit(self, request: Request, bucket: str, max_requests: int, window_seconds: int):
        self._prune_if_needed()
        ip = self._client_ip(request)
        key = f"{bucket}:{ip}"
        now = time.time()
        cutoff = now - window_seconds
        recent = [t for t in self._hits[key] if t > cutoff]
        if len(recent) >= max_requests:
            raise HTTPException(
                status_code=429,
                detail="تم تجاوز الحد المسموح به مؤقتاً حفاظاً على استقرار الخادم؛ يرجى المحاولة بعد قليل.",
            )
        recent.append(now)
        self._hits[key] = recent

    def check_admin_lockout(self, request: Request):
        ip = self._client_ip(request)
        now = time.time()
        recent = [t for t in self._admin_failures[ip] if t > now - 300]
        self._admin_failures[ip] = recent
        if len(recent) >= 5:
            raise HTTPException(
                status_code=429,
                detail="تم قفل محاولات دخول المشرف لمدة 5 دقائق بسبب تكرار الرمز الخاطئ.",
            )

    def record_admin_failure(self, request: Request):
        ip = self._client_ip(request)
        self._admin_failures[ip].append(time.time())

    def clear_admin_failures(self, request: Request):
        ip = self._client_ip(request)
        self._admin_failures.pop(ip, None)


rate_limiter = SimpleRateLimiter()


def is_valid_admin_pin(x_admin_pin: Optional[str]) -> bool:
    if not ADMIN_PIN:
        return False
    provided = (x_admin_pin or "").strip()
    return bool(provided) and hmac.compare_digest(provided, ADMIN_PIN)


def require_admin(x_admin_pin: Optional[str], request: Optional[Request] = None):
    if not ADMIN_PIN:
        raise HTTPException(
            status_code=503,
            detail="لم يتم ضبط متغير البيئة ADMIN_PIN على الخادم؛ يرجى إضافته في إعدادات Render.",
        )
    if request is not None:
        rate_limiter.check_admin_lockout(request)
    if not is_valid_admin_pin(x_admin_pin):
        if request is not None:
            rate_limiter.record_admin_failure(request)
        raise HTTPException(status_code=403, detail="رمز المشرف غير صحيح.")
    if request is not None:
        rate_limiter.clear_admin_failures(request)


# =============================================================================
# 3. نماذج البيانات (Pydantic Models)
# =============================================================================
class AILineAnalysis(BaseModel):
    original_line: str = Field(description="التافلويت (الشطر) الأصلية كما أدخلها الباحث")
    phonetic_diacritized: str = Field(description="التافلويت مضبوطة بالشكل الصوتي الحساني مع ضمّ الساكن لما قبله")
    mutaharrikat_count: int = Field(description="العدد الدقيق للمتحركات في التافلويت مع إهمال السواكن وفق عروض «لغن»")
    estimated_meter: str = Field(description="اسم البَتْ (البحر الحساني من لَبْتُوتَة)")
    rhyme_ending: str = Field(description="الروي أو القافية الصوتية للتافلويت")


class GlossaryTerm(BaseModel):
    term: str = Field(description="المفردة الحسانية")
    arabic_meaning: str = Field(description="معناها بالعربية الفصحى وجذرها إن وُجد")
    english_meaning: str = Field(description="معناها المختصر بالإنجليزية")


class AIPoemAnalysis(BaseModel):
    poem_type: str = Field(description="وحدة/جنس «لغن»: طلعة، گاف/كاف، تبراع، طلعة وگاف، أو نص موزون")
    dominant_meter: str = Field(description="البَتْ (البحر الحساني من لَبْتُوتَة) السائد في القصيدة")
    dominant_rhyme: str = Field(description="بنية القافية والروي العام في القطعة")
    zahr_mode: str = Field(default="", description="الظّهر (المقام أو النغم الموسيقي إن أمكن تحديده)")
    cultural_context: str = Field(description="تحليل موجز للغرض الشعري والسياق الثقافي الحساني")
    linguistic_notes: str = Field(description="ملاحظات لغوية وعروضية منهجية حول التقطيع الإيقاعي وضم السواكن")
    lines: List[AILineAnalysis]
    glossary: List[GlossaryTerm] = Field(description="المفردات الحسانية التراثية الواردة في النص وشرحها")
    english_translation: str = Field(description="ترجمة أدبية دقيقة للإنجليزية")
    french_translation: str = Field(description="ترجمة أدبية دقيقة للفرنسية")


class BatchPoemItem(BaseModel):
    title: str = Field(description="عنوان القصيدة")
    poet: str = Field(default="", description="اسم الشاعر أو الراوي إن ذُكر، وإلا فارغ")
    poem_text: str = Field(description="نص الأبيات الحسانية مرتباً كل شطر في سطر")


class BatchSplitResult(BaseModel):
    poems: List[BatchPoemItem]


class AnalyzeRequest(BaseModel):
    poem_text: str = Field(..., min_length=2, max_length=15000)
    model_name: str = Field(default=DEFAULT_GEMINI_MODEL, max_length=80)
    use_ai: bool = True


class SavePoemRequest(BaseModel):
    title: str = Field(..., max_length=250)
    poet: str = Field(default=DEFAULT_POET_LABEL, max_length=150)
    region: str = Field(default=DEFAULT_REGION_LABEL, max_length=120)
    poem_text: str = Field(..., min_length=2, max_length=20000)
    analysis: Dict
    source_engine: str = Field(default="AI", max_length=80)


class BatchSaveItem(BaseModel):
    title: str = Field(default="", max_length=250)
    poet: str = Field(default=DEFAULT_POET_LABEL, max_length=150)
    region: str = Field(default=DEFAULT_REGION_LABEL, max_length=120)
    poem_text: str = Field(..., min_length=2, max_length=20000)
    analysis: Optional[Dict] = None
    source_engine: str = Field(default="Local (Batch Word)", max_length=80)


class BatchSaveRequest(BaseModel):
    items: List[BatchSaveItem] = Field(..., max_length=200)


class BulkDeletePoemsRequest(BaseModel):
    ids: List[str] = Field(default_factory=list, max_length=500)


class DeleteAllArchiveRequest(BaseModel):
    include_glossary: bool = False
    include_rag: bool = True


class UpdatePoemRequest(BaseModel):
    title: str = Field(..., min_length=1, max_length=250)
    poet: str = Field(..., max_length=150)
    region: str = Field(..., max_length=120)
    poem_text: str = Field(..., min_length=2, max_length=20000)


class VerifyPoemRequest(BaseModel):
    verification_state: VerificationState = Field(..., description="حالة التوثيق: machine_generated | reviewed | verified")
    reviewer_name: str = Field(..., min_length=2, max_length=120, description="اسم الباحث أو المحكّم")
    review_notes: str = Field(default="", max_length=1000)
    dominant_meter: Optional[str] = Field(default=None, max_length=120)
    poem_type: Optional[str] = Field(default=None, max_length=100)
    translation_status: Optional[Literal["machine", "human_reviewed"]] = None


class GlossarySenseCreateRequest(BaseModel):
    term: str = Field(..., min_length=1, max_length=120)
    arabic_meaning: str = Field(..., min_length=2, max_length=1000)
    english_meaning: str = Field(default="", max_length=500)
    source_poem_id: str = Field(default="", max_length=60)
    source_poem_title: str = Field(default="إضافة معجمية مباشرة", max_length=250)
    attestation: str = Field(default="", max_length=600)
    variants: List[str] = Field(default_factory=list, max_length=15)
    status: VerificationState = Field(default="verified")
    editor: str = Field(default="هيئة تحرير ALBIDAN", max_length=120)


class GlossarySenseUpdateRequest(BaseModel):
    term: Optional[str] = Field(default=None, min_length=1, max_length=120)
    arabic_meaning: Optional[str] = Field(default=None, min_length=2, max_length=1000)
    english_meaning: Optional[str] = Field(default=None, max_length=500)
    attestation: Optional[str] = Field(default=None, max_length=600)
    variants: Optional[List[str]] = Field(default=None, max_length=15)
    status: Optional[VerificationState] = None
    editor: Optional[str] = Field(default=None, max_length=120)


class ExportAnalysisDocxRequest(BaseModel):
    title: str = Field(default="تقرير تحليل عروضي حساني", max_length=250)
    poet: str = Field(default=DEFAULT_POET_LABEL, max_length=150)
    region: str = Field(default=DEFAULT_REGION_LABEL, max_length=120)
    poem_text: str = Field(..., min_length=1, max_length=20000)
    analysis: Dict


class ContactMessageRequest(BaseModel):
    name: str = Field(..., min_length=2, max_length=100)
    contact: str = Field(default="", max_length=150)
    category: str = Field(default="تواصل بحثي", max_length=80)
    message: str = Field(..., min_length=5, max_length=2000)


class EnrichGlossaryRequest(BaseModel):
    limit: int = Field(default=10, ge=1, le=25)
    model_name: str = Field(default=DEFAULT_GEMINI_MODEL, max_length=80)


# =============================================================================
# 4. أدوات استيراد وتصدير Word (.docx)
# =============================================================================
def extract_docx_plain_text(raw: bytes) -> str:
    doc = Document(io.BytesIO(raw))
    return "\n".join(p.text.strip() for p in doc.paragraphs).strip()


def split_docx_structurally(raw: bytes, default_poet: str, default_region: str) -> List[Dict]:
    doc = Document(io.BytesIO(raw))
    poems: List[Dict] = []
    cur_title, cur_lines = "", []

    def flush_block():
        nonlocal cur_title, cur_lines
        clean = [l.strip() for l in cur_lines if l.strip()]
        if not clean:
            cur_title, cur_lines = "", []
            return
        title, body_lines = cur_title.strip(), clean
        if not title and len(clean) >= 2 and (
            clean[0].startswith(("قصيدة", "عنوان", "كاف", "گاف", "طلعة", "تبراع", "#"))
            or (len(clean[0]) <= 48 and clean[0].endswith(":"))
            or any(p.search(clean[0]) for p in POET_EMBED_PATTERNS)
        ):
            title = re.sub(r"^[#\s]+|[:\s]+$", "", clean[0])
            body_lines = clean[1:]
        if body_lines:
            clean_title, clean_poet = clean_title_and_poet(
                title or f"قصيدة {len(poems) + 1}",
                default_poet.strip() or DEFAULT_POET_LABEL,
            )
            poems.append({
                "title": clean_title,
                "poet": clean_poet,
                "region": default_region.strip() or DEFAULT_REGION_LABEL,
                "poem_text": "\n".join(body_lines),
            })
        cur_title, cur_lines = "", []

    for p in doc.paragraphs:
        txt = p.text.strip()
        style_name = (p.style.name if p.style else "") or ""
        if not txt or re.match(r"^[-_*=•]{3,}$", txt):
            if cur_lines:
                flush_block()
            continue
        if style_name.lower().startswith("heading") or txt.startswith("#"):
            if cur_lines:
                flush_block()
            cur_title = txt.lstrip("#").strip()
            continue
        cur_lines.append(txt)
    if cur_lines:
        flush_block()
    return poems


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
        run.font.color.rgb = RGBColor(int(color_hex[0:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16))
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), font_name)
    rtl = OxmlElement("w:rtl")
    rtl.set(qn("w:val"), "1")
    rPr.append(rtl)


def set_table_rtl(table):
    table._tbl.tblPr.append(OxmlElement("w:bidiVisual"))


def set_cell_shading(cell, fill_hex: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    cell._tc.get_or_add_tcPr().append(shd)


def add_rtl_para(doc: Document, text: str, size: float = 11.5, bold: bool = False, color: Optional[str] = None, space_after: float = 6):
    p = doc.add_paragraph()
    set_paragraph_rtl(p)
    p.paragraph_format.space_after = Pt(space_after)
    set_run_arabic(p.add_run(str(text or "")), size_pt=size, bold=bold, color_hex=color)
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
        run.font.color.rgb = RGBColor(int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16))
    return p


def create_base_arabic_doc(doc_title: str, subtitle: str = "") -> Document:
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Inches(0.8)
    add_rtl_para(doc, "مختبر البيظان للذكاء الاصطناعي — ALBIDAN AI LAB", size=10, bold=True, color="9A6B13", space_after=2)
    add_rtl_para(doc, doc_title, size=16, bold=True, color="1B365D", space_after=4)
    if subtitle:
        add_rtl_para(doc, subtitle, size=10, color="555555", space_after=10)
    return doc


def append_poem_section_to_doc(doc: Document, title: str, poet: str, region: str, poem_text: str, analysis: Dict):
    add_rtl_para(doc, f"القصيدة: {title}", size=14, bold=True, color="1B365D", space_after=3)
    zahr = (analysis.get("zahr_mode") or "").strip()
    zahr_part = f"  |  الظّهر (المقام): {zahr}" if zahr else ""
    meta = f"الشاعر / الراوي: {poet}  |  المنطقة: {region}  |  الوحدة / الجنس: {analysis.get('poem_type', '—')}  |  البَتْ (البحر): {analysis.get('dominant_meter', '—')}{zahr_part}"
    add_rtl_para(doc, meta, size=10, color="555555", space_after=8)
    add_rtl_para(doc, "النص الشعري:", size=11.5, bold=True, color="9A6B13", space_after=3)
    for line in [l.strip() for l in poem_text.split("\n") if l.strip()]:
        add_rtl_para(doc, line, size=13, color="111111", space_after=3)
    if analysis.get("cultural_context"):
        add_rtl_para(doc, "السياق الثقافي والغرض الشعري:", size=11, bold=True, color="1B365D", space_after=2)
        add_rtl_para(doc, analysis["cultural_context"], size=10.5, space_after=6)
    if analysis.get("linguistic_notes"):
        add_rtl_para(doc, "ملاحظات عروضية ولغوية:", size=11, bold=True, color="1B365D", space_after=2)
        add_rtl_para(doc, analysis["linguistic_notes"], size=10.5, space_after=6)
    lines_data = analysis.get("lines") or []
    if lines_data:
        add_rtl_para(doc, "جدول التقطيع الصوتي والعروضي:", size=11, bold=True, color="1B365D", space_after=4)
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        set_table_rtl(tbl)
        for i, h in enumerate(["التافلويت الأصلية", "التشكيل الصوتي", "المتحركات", "البَتْ / البحر", "الروي"]):
            cell = tbl.rows[0].cells[i]
            set_cell_shading(cell, "1B365D")
            p = cell.paragraphs[0]
            set_paragraph_rtl(p)
            set_run_arabic(p.add_run(h), size_pt=10, bold=True, color_hex="FFFFFF")
        for row_item in lines_data:
            cells = tbl.add_row().cells
            for i, v in enumerate([
                row_item.get("original_line", ""),
                row_item.get("phonetic_diacritized", ""),
                str(row_item.get("mutaharrikat_count", "")),
                row_item.get("estimated_meter", ""),
                row_item.get("rhyme_ending", ""),
            ]):
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
        add_ltr_para(doc, "English Literary Translation (AI-assisted unless verified):", size=10.5, bold=True, color="1B365D")
        add_ltr_para(doc, analysis["english_translation"], size=10.5)
    if analysis.get("french_translation"):
        add_ltr_para(doc, "Traduction Littéraire (Français — assistée par IA sauf mention):", size=10.5, bold=True, color="1B365D")
        add_ltr_para(doc, analysis["french_translation"], size=10.5)


def doc_to_stream(doc: Document, filename: str) -> StreamingResponse:
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    ascii_fallback = re.sub(r"[^A-Za-z0-9._-]", "_", filename) or "albidan_export.docx"
    utf8_encoded = quote(filename)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{ascii_fallback}"; filename*=UTF-8\'\'{utf8_encoded}'
        },
    )


# =============================================================================
# 5. المحرك العروضي المحلي + العدّ البرمجي الحتمي للمتحركات (للتحقق من AI)
# =============================================================================
class LocalHassaniyaEngine:
    short_vowels = set("َُِ")
    tanween = set("ًٌٍ")
    sukun = "ْ"
    shadda = "ّ"
    long_vowels = set("اويى")
    arabic_letters_re = re.compile(r"[\u0621-\u064Aگݣڭڨڤ]")
    meters_map = {
        5: "اسغير / إملخ (تقدير: 5-6 متحركات)",
        6: "اسغير / لبير (تقدير: 6 متحركات)",
        7: "لبتيت الناقص / بوعمران (تقدير: 7 متحركات)",
        8: "لبتيت التام (تقدير: 8 متحركات)",
        9: "البت الكبير (تقدير: 9 متحركات)",
        10: "البت الكبير (تقدير: 10 متحركات)",
        11: "البت الكبير (تقدير: 11 متحرك)",
        12: "البت الكبير (تقدير: 12 متحرك)",
    }

    def count_diacritized_mutaharrikat(self, diacritized_line: str) -> Optional[int]:
        """عدّ برمجي حتمي للمتحركات انطلاقاً من النص المضبوط بالشكل الصوتي."""
        if not diacritized_line:
            return None
        letters = self.arabic_letters_re.findall(diacritized_line)
        consonants = [ch for ch in letters if ch not in self.long_vowels]
        if not consonants:
            return None
        count = 0
        for ch in diacritized_line:
            if ch in self.short_vowels or ch in self.tanween:
                count += 1
        # لا نعتبر النص مضبوطاً فعلياً إلا إذا حمل حركات كافية مقارنة بعدد الحروف
        if count >= 4 and (count / max(1, len(consonants))) >= 0.35:
            return count
        return None

    def cross_check_ai_analysis(self, analysis: Dict) -> Dict:
        """يقارن عدّ النموذج اللغوي بالعدّ البرمجي الحتمي ويوسم أي تعارض للمراجعة البشرية."""
        lines = analysis.get("lines") or []
        mismatches = 0
        for item in lines:
            diac = str(item.get("phonetic_diacritized") or "")
            det_count = self.count_diacritized_mutaharrikat(diac)
            ai_count = item.get("mutaharrikat_count")
            item["deterministic_count"] = det_count
            if det_count is not None and isinstance(ai_count, int) and abs(det_count - ai_count) >= 1:
                item["meter_mismatch"] = True
                mismatches += 1
            else:
                item["meter_mismatch"] = False
        analysis["deterministic_mismatches"] = mismatches
        analysis["translation_status"] = analysis.get("translation_status") or "machine"
        return analysis

    def _estimate_line_mutaharrikat(self, line: str) -> int:
        norm = re.sub(r"[ݣڭڨڤ]", "گ", line)
        det = self.count_diacritized_mutaharrikat(norm)
        if det is not None:
            return det

        est = 0.0
        for raw_w in norm.split():
            w = ARABIC_DIACRITICS_RE.sub("", raw_w)
            w = "".join(self.arabic_letters_re.findall(w))
            if not w:
                continue
            if len(w) > 3 and w.startswith("ال"):
                w = w[2:]
            c_count = sum(1 for ch in w if ch not in self.long_vowels)
            if c_count <= 1:
                est += 1.0
            elif c_count == 2:
                est += 1.35
            else:
                est += max(1.5, c_count * 0.62)
        return max(4, min(14, int(round(est))))

    def _extract_rhyme_rawi(self, line: str) -> str:
        """يستخرج الروي مع مراعاة حرف المد السابق له إن وُجد."""
        norm = re.sub(r"[ݣڭڨڤ]", "گ", line)
        clean = re.sub(r"[\u0610-\u061A\u064B-\u065F\u0670\u0640\s\W]", "", norm)
        if not clean:
            return "—"
        if len(clean) >= 3 and clean[-2] in self.long_vowels:
            return clean[-3:]
        return clean[-2:] if len(clean) >= 2 else clean

    def _meter_label(self, count: int) -> str:
        if count in self.meters_map:
            return self.meters_map[count]
        if count > 12:
            return "البت الكبير (تقدير ممتد)"
        return "قيد التدقيق العروضي"

    def analyze(self, poem_text: str) -> Dict:
        lines = [l.strip() for l in poem_text.strip().split("\n") if l.strip()]
        details = []
        counts = []

        for line in lines:
            norm = re.sub(r"[ݣڭڨڤ]", "گ", line)
            count = self._estimate_line_mutaharrikat(norm)
            counts.append(count)
            rhyme = self._extract_rhyme_rawi(norm)
            details.append({
                "original_line": line,
                "phonetic_diacritized": "تحليل أولي — يتطلب تشكيلاً صوتياً ومراجعة بشرية",
                "mutaharrikat_count": count,
                "deterministic_count": self.count_diacritized_mutaharrikat(norm),
                "meter_mismatch": False,
                "estimated_meter": self._meter_label(count),
                "rhyme_ending": rhyme,
            })

        dominant_meter = self._meter_label(sorted(counts)[len(counts) // 2]) if counts else "قيد التدقيق العروضي"
        rhyme_counts = Counter(d["rhyme_ending"] for d in details if d["rhyme_ending"] and d["rhyme_ending"] != "—")
        dominant_rhyme = rhyme_counts.most_common(1)[0][0] if rhyme_counts else "—"

        n = len(details)
        first3_rhymes = [d["rhyme_ending"] for d in details[:3]]
        has_tal3a_opening = n >= 3 and len(set(first3_rhymes)) == 1 and first3_rhymes[0] != "—"

        if n == 2:
            poem_type = "تبراع (Tebraa)"
        elif n == 3:
            poem_type = "طلعة (Tal'aa)"
        elif n == 4:
            poem_type = "گاف (Gaf)"
        elif n in (6, 8):
            poem_type = "طلعة وگاف (Tal'aa & Gaf)" if (has_tal3a_opening and details[3]["rhyme_ending"] != first3_rhymes[0]) else "گاف (Gaf)"
        elif n >= 5:
            poem_type = "طلعة (Tal'aa)" if has_tal3a_opening else "نص موزون (لغن)"
        else:
            poem_type = "نص شعري (لغن)"

        return {
            "poem_type": poem_type,
            "dominant_meter": dominant_meter,
            "dominant_rhyme": dominant_rhyme,
            "zahr_mode": "يُحدَّد بعد المراجعة العروضية (الظّهر)",
            "cultural_context": "تحليل آلي أولي؛ يتطلب مراجعة باحث متخصص لتأكيد البَتْ والظّهر والسياق.",
            "linguistic_notes": "تقدير محلي لعدّ المتحركات مع مراعاة حرف المد في الروي؛ يحتاج إلى تأكيد بشري.",
            "lines": details,
            "glossary": [],
            "english_translation": "",
            "french_translation": "",
            "translation_status": "machine",
            "deterministic_mismatches": 0,
        }


# =============================================================================
# 6. مدير قواعد البيانات والأرشيف (Pagination + Indexed Search + Multi-Sense Glossary)
# =============================================================================
class ArchiveManager:
    def __init__(self, mongo_uri: str, db_name: str, sqlite_path: str = "albidan_archive.db"):
        self.sqlite_path = str(BASE_DIR / sqlite_path)
        self.db_name = db_name
        self.active_mode = "sqlite"
        self.status_message = f"قاعدة SQLite المحلية ({sqlite_path})"
        self.mongo_db = None
        self._stats_cache: Optional[Dict] = None
        self._stats_cached_at: float = 0.0
        self._sqlite_fts_enabled = False

        if mongo_uri and PYMONGO_AVAILABLE:
            try:
                client = MongoClient(mongo_uri, serverSelectionTimeoutMS=5000)
                client.admin.command("ping")
                self.mongo_db = client[db_name]
                self.active_mode = "mongodb"
                self.status_message = f"متصل سحابياً بـ MongoDB Atlas ({db_name})"
                self._ensure_mongo_indexes()
            except Exception as exc:
                logger.warning("MongoDB connection failed, falling back to SQLite: %s", exc)
                self.status_message = f"تعذر الاتصال بـ MongoDB ({exc})؛ تم التحويل إلى SQLite."
                self._init_sqlite()
        else:
            self._init_sqlite()

    def _invalidate_stats_cache(self):
        self._stats_cache = None
        self._stats_cached_at = 0.0

    def _ensure_mongo_indexes(self):
        if self.mongo_db is None:
            return
        try:
            poems = self.mongo_db["poems"]
            poems.create_index([("created_at", DESCENDING)])
            poems.create_index([("poet", ASCENDING), ("dominant_meter", ASCENDING)])
            poems.create_index([("poem_type", ASCENDING)])
            poems.create_index([("verification_state", ASCENDING)])
            poems.create_index([("search_text", ASCENDING)])
            poems.create_index([("search_text_flex", ASCENDING)])

            gloss = self.mongo_db["glossary"]
            gloss.create_index([("sense_id", ASCENDING)])
            gloss.create_index([("normalized_term", ASCENDING), ("source_poem_id", ASCENDING)])
            gloss.create_index([("status", ASCENDING), ("updated_at", DESCENDING)])

            self.mongo_db["messages"].create_index([("created_at", DESCENDING)])
        except Exception as exc:
            logger.warning("Failed ensuring MongoDB indexes: %s", exc)

    def _init_sqlite(self):
        with sqlite3.connect(self.sqlite_path) as conn:
            conn.execute(
                "CREATE TABLE IF NOT EXISTS poems ("
                "id TEXT PRIMARY KEY, title TEXT NOT NULL, poet TEXT, region TEXT, "
                "poem_text TEXT NOT NULL, poem_type TEXT, dominant_meter TEXT, "
                "azawan_mode TEXT, source_engine TEXT, analysis_json TEXT NOT NULL, "
                "music_json TEXT NOT NULL, created_at TEXT NOT NULL)"
            )
            # ترقية آمنة للأعمدة الجديدة في SQLite دون فقدان البيانات القديمة
            existing_cols = {row[1] for row in conn.execute("PRAGMA table_info(poems)").fetchall()}
            migrations = {
                "verification_state": "TEXT DEFAULT 'machine_generated'",
                "verification_status": "TEXT DEFAULT 'مولّد آلياً (غير مراجَع)'",
                "reviewed_by": "TEXT DEFAULT ''",
                "reviewed_at": "TEXT DEFAULT ''",
                "review_notes": "TEXT DEFAULT ''",
                "search_text": "TEXT DEFAULT ''",
                "search_text_flex": "TEXT DEFAULT ''",
            }
            for col_name, col_def in migrations.items():
                if col_name not in existing_cols:
                    conn.execute(f"ALTER TABLE poems ADD COLUMN {col_name} {col_def}")

            conn.execute("CREATE INDEX IF NOT EXISTS idx_poems_created ON poems(created_at DESC)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_poems_poet ON poems(poet)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_poems_meter ON poems(dominant_meter)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_poems_state ON poems(verification_state)")

            # تفعيل البحث المفهرس SQLite FTS5 إن كان متوفراً
            try:
                conn.execute(
                    "CREATE VIRTUAL TABLE IF NOT EXISTS poems_fts USING fts5("
                    "poem_id UNINDEXED, search_text, search_text_flex)"
                )
                self._sqlite_fts_enabled = True
            except sqlite3.OperationalError:
                self._sqlite_fts_enabled = False

            # جدول المعجم متعدد المعاني (لا يمحو المتجانسات ولا الموثّق بشرياً)
            conn.execute(
                "CREATE TABLE IF NOT EXISTS glossary_senses ("
                "sense_id TEXT PRIMARY KEY, term TEXT NOT NULL, normalized_term TEXT NOT NULL, "
                "arabic_meaning TEXT, english_meaning TEXT, source_poem_id TEXT DEFAULT '', "
                "source_poem_title TEXT DEFAULT '', attestation TEXT DEFAULT '', "
                "variants_json TEXT DEFAULT '[]', status TEXT DEFAULT 'machine_generated', "
                "editor TEXT DEFAULT 'AI', updated_at TEXT NOT NULL)"
            )
            conn.execute("CREATE INDEX IF NOT EXISTS idx_gloss_norm ON glossary_senses(normalized_term)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_gloss_poem ON glossary_senses(source_poem_id)")

            # ترحيل تلقائي لمرة واحدة من جدول glossary القديم إن وجد
            old_table = conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='glossary'").fetchone()
            if old_table:
                new_count = conn.execute("SELECT COUNT(*) FROM glossary_senses").fetchone()[0]
                if new_count == 0:
                    old_rows = conn.execute("SELECT term, arabic_meaning, english_meaning, source_poem_title, updated_at FROM glossary").fetchall()
                    for r in old_rows:
                        conn.execute(
                            "INSERT OR IGNORE INTO glossary_senses "
                            "(sense_id, term, normalized_term, arabic_meaning, english_meaning, source_poem_id, source_poem_title, status, editor, updated_at) "
                            "VALUES (?, ?, ?, ?, ?, '', ?, 'machine_generated', 'Legacy Import', ?)",
                            (new_record_id(10), r[0], normalize_hassaniya(r[0]), r[1], r[2], r[3], r[4]),
                        )

            conn.execute(
                "CREATE TABLE IF NOT EXISTS messages ("
                "id TEXT PRIMARY KEY, name TEXT NOT NULL, contact TEXT, category TEXT, "
                "message TEXT NOT NULL, created_at TEXT NOT NULL)"
            )
            conn.commit()

    def _sync_sqlite_fts(self, conn: sqlite3.Connection, poem_id: str, s_strict: str, s_flex: str):
        if not self._sqlite_fts_enabled:
            return
        conn.execute("DELETE FROM poems_fts WHERE poem_id = ?", (poem_id,))
        conn.execute("INSERT INTO poems_fts (poem_id, search_text, search_text_flex) VALUES (?, ?, ?)", (poem_id, s_strict, s_flex))

    def _normalize_poem_doc(self, doc: Dict, include_analysis: bool = True) -> Dict:
        state = doc.get("verification_state") or "machine_generated"
        if state not in VERIFICATION_LABELS:
            state = "machine_generated"
        res = {
            "id": doc.get("id") or str(doc.get("_id", "")),
            "title": doc.get("title", "قصيدة بدون عنوان"),
            "poet": doc.get("poet", DEFAULT_POET_LABEL),
            "region": doc.get("region", DEFAULT_REGION_LABEL),
            "poem_text": doc.get("poem_text", ""),
            "poem_type": doc.get("poem_type", "نص شعري"),
            "dominant_meter": doc.get("dominant_meter", "قيد التدقيق العروضي"),
            "source_engine": doc.get("source_engine", "Local"),
            "verification_state": state,
            "verification_status": doc.get("verification_status") or VERIFICATION_LABELS[state],
            "reviewed_by": doc.get("reviewed_by", ""),
            "reviewed_at": doc.get("reviewed_at", ""),
            "review_notes": doc.get("review_notes", ""),
            "created_at": doc.get("created_at", ""),
        }
        if include_analysis:
            res["analysis"] = doc.get("analysis") or {}
        return res

    def _prepare_poem_record(self, payload: SavePoemRequest, fallback_title: str = "قصيدة بدون عنوان") -> Dict:
        record_id = new_record_id(10)
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M")
        clean_title, clean_poet = clean_title_and_poet(payload.title or fallback_title, payload.poet)
        region = (payload.region or "").strip() or DEFAULT_REGION_LABEL
        poem_text = payload.poem_text.strip()
        analysis = payload.analysis or {}
        dominant_meter = analysis.get("dominant_meter", "قيد التدقيق العروضي")
        poem_type = analysis.get("poem_type", "غير محدد")
        s_strict, s_flex = build_search_fields(clean_title, clean_poet, region, poem_text, dominant_meter)

        return {
            "id": record_id,
            "title": clean_title,
            "poet": clean_poet,
            "region": region,
            "poem_text": poem_text,
            "poem_type": poem_type,
            "dominant_meter": dominant_meter,
            "source_engine": payload.source_engine,
            "verification_state": "machine_generated",
            "verification_status": VERIFICATION_LABELS["machine_generated"],
            "reviewed_by": "",
            "reviewed_at": "",
            "review_notes": "",
            "search_text": s_strict,
            "search_text_flex": s_flex,
            "analysis": analysis,
            "created_at": now_iso,
        }

    def _upsert_glossary_senses(self, entries: List[Dict], allow_overwrite_verified: bool = False):
        """
        حفظ معاني المعجم كمداخل مستقلة مرتبطة بالقصيدة المصدر.
        يحظر على المخرجات الآلية الكتابة فوق أي معنى موثّق بشرياً (verified).
        """
        if not entries:
            return
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M")

        if self.active_mode == "mongodb":
            col = self.mongo_db["glossary"]
            for g in entries:
                term = str(g.get("term", "")).strip()
                if not term:
                    continue
                norm_term = normalize_hassaniya(term)
                poem_id = str(g.get("source_poem_id", "")).strip()
                status = g.get("status", "machine_generated")
                editor = g.get("editor", "AI")

                # البحث عن معنى مطابق لنفس الكلمة ونفس القصيدة المصدر (حتى لا تتداخل المعاني المتجانسة)
                existing = col.find_one({"normalized_term": norm_term, "source_poem_id": poem_id})
                if existing:
                    if existing.get("status") == "verified" and not allow_overwrite_verified:
                        continue
                    col.update_one(
                        {"_id": existing["_id"]},
                        {
                            "$set": {
                                "term": term,
                                "arabic_meaning": g.get("arabic_meaning", existing.get("arabic_meaning", "")),
                                "english_meaning": g.get("english_meaning", existing.get("english_meaning", "")),
                                "source_poem_title": g.get("source_poem_title", existing.get("source_poem_title", "")),
                                "attestation": g.get("attestation", existing.get("attestation", "")),
                                "variants": g.get("variants", existing.get("variants", [])),
                                "status": status,
                                "editor": editor,
                                "updated_at": g.get("updated_at", now_iso),
                            }
                        },
                    )
                else:
                    sense_id = g.get("sense_id") or new_record_id(10)
                    doc = {
                        "_id": sense_id,
                        "sense_id": sense_id,
                        "term": term,
                        "normalized_term": norm_term,
                        "arabic_meaning": g.get("arabic_meaning", ""),
                        "english_meaning": g.get("english_meaning", ""),
                        "source_poem_id": poem_id,
                        "source_poem_title": g.get("source_poem_title", "قصيدة موثّقة"),
                        "attestation": g.get("attestation", ""),
                        "variants": g.get("variants", []),
                        "status": status,
                        "editor": editor,
                        "updated_at": g.get("updated_at", now_iso),
                    }
                    col.insert_one(doc)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                for g in entries:
                    term = str(g.get("term", "")).strip()
                    if not term:
                        continue
                    norm_term = normalize_hassaniya(term)
                    poem_id = str(g.get("source_poem_id", "")).strip()
                    status = g.get("status", "machine_generated")
                    editor = g.get("editor", "AI")
                    row = conn.execute(
                        "SELECT sense_id, status FROM glossary_senses WHERE normalized_term = ? AND source_poem_id = ?",
                        (norm_term, poem_id),
                    ).fetchone()
                    if row:
                        if row[1] == "verified" and not allow_overwrite_verified:
                            continue
                        conn.execute(
                            "UPDATE glossary_senses SET term=?, arabic_meaning=?, english_meaning=?, source_poem_title=?, "
                            "attestation=?, variants_json=?, status=?, editor=?, updated_at=? WHERE sense_id=?",
                            (
                                term,
                                g.get("arabic_meaning", ""),
                                g.get("english_meaning", ""),
                                g.get("source_poem_title", ""),
                                g.get("attestation", ""),
                                json.dumps(g.get("variants", []), ensure_ascii=False),
                                status,
                                editor,
                                g.get("updated_at", now_iso),
                                row[0],
                            ),
                        )
                    else:
                        sense_id = g.get("sense_id") or new_record_id(10)
                        conn.execute(
                            "INSERT INTO glossary_senses "
                            "(sense_id, term, normalized_term, arabic_meaning, english_meaning, source_poem_id, "
                            "source_poem_title, attestation, variants_json, status, editor, updated_at) "
                            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                            (
                                sense_id,
                                term,
                                norm_term,
                                g.get("arabic_meaning", ""),
                                g.get("english_meaning", ""),
                                poem_id,
                                g.get("source_poem_title", "قصيدة موثّقة"),
                                g.get("attestation", ""),
                                json.dumps(g.get("variants", []), ensure_ascii=False),
                                status,
                                editor,
                                g.get("updated_at", now_iso),
                            ),
                        )
                conn.commit()

    def _extract_glossary_entries_from_record(self, record: Dict) -> List[Dict]:
        poem_lines = [l.strip() for l in record.get("poem_text", "").split("\n") if l.strip()]
        out = []
        for g in (record.get("analysis", {}).get("glossary") or []):
            term = str(g.get("term", "")).strip()
            if not term:
                continue
            norm_t = normalize_hassaniya(term)
            att = next((ln for ln in poem_lines if norm_t in normalize_hassaniya(ln)), "")
            out.append({
                "term": term,
                "arabic_meaning": g.get("arabic_meaning", ""),
                "english_meaning": g.get("english_meaning", ""),
                "source_poem_id": record["id"],
                "source_poem_title": record["title"],
                "attestation": att,
                "variants": [],
                "status": "machine_generated",
                "editor": record.get("source_engine", "AI"),
                "updated_at": record["created_at"],
            })
        return out

    def save_poem(self, payload: SavePoemRequest) -> str:
        record = self._prepare_poem_record(payload)
        glossary_items = self._extract_glossary_entries_from_record(record)

        if self.active_mode == "mongodb":
            self.mongo_db["poems"].insert_one(dict(record, _id=record["id"]))
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.execute(
                    "INSERT INTO poems (id, title, poet, region, poem_text, poem_type, dominant_meter, azawan_mode, "
                    "source_engine, analysis_json, music_json, created_at, verification_state, verification_status, "
                    "reviewed_by, reviewed_at, review_notes, search_text, search_text_flex) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (
                        record["id"], record["title"], record["poet"], record["region"],
                        record["poem_text"], record["poem_type"], record["dominant_meter"], "",
                        record["source_engine"], json.dumps(record["analysis"], ensure_ascii=False), "{}",
                        record["created_at"], record["verification_state"], record["verification_status"],
                        "", "", "", record["search_text"], record["search_text_flex"],
                    ),
                )
                self._sync_sqlite_fts(conn, record["id"], record["search_text"], record["search_text_flex"])
                conn.commit()

        self._upsert_glossary_senses(glossary_items, allow_overwrite_verified=False)
        self._invalidate_stats_cache()
        return record["id"]

    def save_poems_batch(self, items: List[SavePoemRequest]) -> List[str]:
        saved_ids, docs_to_insert, glossary_updates = [], [], []
        for idx, payload in enumerate(items, start=1):
            if not payload.poem_text.strip():
                continue
            record = self._prepare_poem_record(payload, fallback_title=f"قصيدة {idx}")
            saved_ids.append(record["id"])
            docs_to_insert.append(record)
            glossary_updates.extend(self._extract_glossary_entries_from_record(record))

        if not docs_to_insert:
            return []

        if self.active_mode == "mongodb":
            self.mongo_db["poems"].insert_many([dict(d, _id=d["id"]) for d in docs_to_insert])
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.executemany(
                    "INSERT INTO poems (id, title, poet, region, poem_text, poem_type, dominant_meter, azawan_mode, "
                    "source_engine, analysis_json, music_json, created_at, verification_state, verification_status, "
                    "reviewed_by, reviewed_at, review_notes, search_text, search_text_flex) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    [
                        (
                            d["id"], d["title"], d["poet"], d["region"], d["poem_text"],
                            d["poem_type"], d["dominant_meter"], "", d["source_engine"],
                            json.dumps(d["analysis"], ensure_ascii=False), "{}", d["created_at"],
                            d["verification_state"], d["verification_status"], "", "", "",
                            d["search_text"], d["search_text_flex"],
                        )
                        for d in docs_to_insert
                    ],
                )
                for d in docs_to_insert:
                    self._sync_sqlite_fts(conn, d["id"], d["search_text"], d["search_text_flex"])
                conn.commit()

        self._upsert_glossary_senses(glossary_updates, allow_overwrite_verified=False)
        self._invalidate_stats_cache()
        return saved_ids

    def get_poem(self, record_id: str) -> Optional[Dict]:
        if self.active_mode == "mongodb":
            doc = self.mongo_db["poems"].find_one({"$or": [{"_id": record_id}, {"id": record_id}]})
            return self._normalize_poem_doc(doc, include_analysis=True) if doc else None
        with sqlite3.connect(self.sqlite_path) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT * FROM poems WHERE id = ?", (record_id,)).fetchone()
            if not row:
                return None
            d = dict(row)
            d["analysis"] = json.loads(d.pop("analysis_json", "{}") or "{}")
            return self._normalize_poem_doc(d, include_analysis=True)

    def update_poem(self, record_id: str, payload: UpdatePoemRequest) -> bool:
        existing = self.get_poem(record_id)
        if not existing:
            return False
        clean_title, clean_poet = clean_title_and_poet(payload.title, payload.poet)
        region = payload.region.strip() or DEFAULT_REGION_LABEL
        poem_text = payload.poem_text.strip()
        dominant_meter = existing.get("dominant_meter", "قيد التدقيق العروضي")
        s_strict, s_flex = build_search_fields(clean_title, clean_poet, region, poem_text, dominant_meter)
        updates = {
            "title": clean_title,
            "poet": clean_poet,
            "region": region,
            "poem_text": poem_text,
            "search_text": s_strict,
            "search_text_flex": s_flex,
        }
        if self.active_mode == "mongodb":
            ok = self.mongo_db["poems"].update_one({"$or": [{"_id": record_id}, {"id": record_id}]}, {"$set": updates}).matched_count > 0
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                cur = conn.execute(
                    "UPDATE poems SET title=?, poet=?, region=?, poem_text=?, search_text=?, search_text_flex=? WHERE id=?",
                    (clean_title, clean_poet, region, poem_text, s_strict, s_flex, record_id),
                )
                self._sync_sqlite_fts(conn, record_id, s_strict, s_flex)
                conn.commit()
                ok = cur.rowcount > 0
        if ok:
            self._invalidate_stats_cache()
        return ok

    def verify_poem(self, record_id: str, payload: VerifyPoemRequest) -> Optional[Dict]:
        """تحديث حالة التوثيق الأكاديمي (مولّد آلياً -> قيد المراجعة -> موثّق ومحكّم)."""
        rec = self.get_poem(record_id)
        if not rec:
            return None
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M")
        label = VERIFICATION_LABELS[payload.verification_state]
        dominant_meter = (payload.dominant_meter or "").strip() or rec["dominant_meter"]
        poem_type = (payload.poem_type or "").strip() or rec["poem_type"]
        analysis = dict(rec.get("analysis") or {})
        analysis["dominant_meter"] = dominant_meter
        analysis["poem_type"] = poem_type
        if payload.translation_status:
            analysis["translation_status"] = payload.translation_status

        s_strict, s_flex = build_search_fields(rec["title"], rec["poet"], rec["region"], rec["poem_text"], dominant_meter)
        updates = {
            "verification_state": payload.verification_state,
            "verification_status": label,
            "reviewed_by": payload.reviewer_name.strip(),
            "reviewed_at": now_iso,
            "review_notes": payload.review_notes.strip(),
            "dominant_meter": dominant_meter,
            "poem_type": poem_type,
            "analysis": analysis,
            "search_text": s_strict,
            "search_text_flex": s_flex,
        }
        if self.active_mode == "mongodb":
            self.mongo_db["poems"].update_one({"$or": [{"_id": record_id}, {"id": record_id}]}, {"$set": updates})
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.execute(
                    "UPDATE poems SET verification_state=?, verification_status=?, reviewed_by=?, reviewed_at=?, "
                    "review_notes=?, dominant_meter=?, poem_type=?, analysis_json=?, search_text=?, search_text_flex=? WHERE id=?",
                    (
                        updates["verification_state"], updates["verification_status"], updates["reviewed_by"],
                        updates["reviewed_at"], updates["review_notes"], dominant_meter, poem_type,
                        json.dumps(analysis, ensure_ascii=False), s_strict, s_flex, record_id,
                    ),
                )
                self._sync_sqlite_fts(conn, record_id, s_strict, s_flex)
                conn.commit()
        self._invalidate_stats_cache()
        return self.get_poem(record_id)

    def update_poem_analysis(self, record_id: str, analysis: Dict, source_engine: str) -> bool:
        rec = self.get_poem(record_id)
        if not rec:
            return False
        dominant_meter = analysis.get("dominant_meter", rec.get("dominant_meter", "قيد التدقيق العروضي"))
        poem_type = analysis.get("poem_type", rec.get("poem_type", "نص شعري"))
        s_strict, s_flex = build_search_fields(rec["title"], rec["poet"], rec["region"], rec["poem_text"], dominant_meter)
        rec_copy = dict(rec, analysis=analysis, source_engine=source_engine, created_at=datetime.now().strftime("%Y-%m-%d %H:%M"))
        glossary_items = self._extract_glossary_entries_from_record(rec_copy)

        if self.active_mode == "mongodb":
            self.mongo_db["poems"].update_one(
                {"$or": [{"_id": record_id}, {"id": record_id}]},
                {
                    "$set": {
                        "analysis": analysis,
                        "dominant_meter": dominant_meter,
                        "poem_type": poem_type,
                        "source_engine": source_engine,
                        "search_text": s_strict,
                        "search_text_flex": s_flex,
                    }
                },
            )
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.execute(
                    "UPDATE poems SET poem_type=?, dominant_meter=?, source_engine=?, analysis_json=?, "
                    "search_text=?, search_text_flex=? WHERE id=?",
                    (poem_type, dominant_meter, source_engine, json.dumps(analysis, ensure_ascii=False), s_strict, s_flex, record_id),
                )
                self._sync_sqlite_fts(conn, record_id, s_strict, s_flex)
                conn.commit()
        self._upsert_glossary_senses(glossary_items, allow_overwrite_verified=False)
        self._invalidate_stats_cache()
        return True

    def list_poems_paginated(
        self,
        q: str = "",
        poem_type: str = "الكل",
        poet: str = "",
        meter: str = "",
        region: str = "",
        verification_state: str = "",
        flexible: bool = False,
        page: int = 1,
        limit: int = 24,
        include_analysis: bool = True,
    ) -> Dict[str, Any]:
        """استعلام مفهرس بترقيم الصفحات مباشرة من قاعدة البيانات دون تحميل المدونة كاملة في الذاكرة."""
        offset = (page - 1) * limit

        if self.active_mode == "mongodb":
            query: Dict[str, Any] = {}
            if poem_type and poem_type != "الكل":
                query["poem_type"] = {"$regex": re.escape(poem_type.strip()), "$options": "i"}
            if poet.strip():
                query["poet"] = {"$regex": re.escape(poet.strip()), "$options": "i"}
            if meter.strip():
                query["dominant_meter"] = {"$regex": re.escape(meter.strip()), "$options": "i"}
            if region.strip():
                query["region"] = {"$regex": re.escape(region.strip()), "$options": "i"}
            if verification_state.strip() in VERIFICATION_LABELS:
                query["verification_state"] = verification_state.strip()

            if q.strip():
                q_norm = normalize_hassaniya(q, unify_gaf_to_kaf=flexible)
                field = "search_text_flex" if flexible else "search_text"
                query["$or"] = [
                    {field: {"$regex": re.escape(q_norm)}},
                    {"title": {"$regex": re.escape(q.strip()), "$options": "i"}},
                    {"poem_text": {"$regex": re.escape(q.strip()), "$options": "i"}},
                ]

            col = self.mongo_db["poems"]
            total = col.count_documents(query)
            projection = None if include_analysis else {"analysis": 0}
            cursor = col.find(query, projection).sort("created_at", DESCENDING).skip(offset).limit(limit)
            items = [self._normalize_poem_doc(doc, include_analysis=include_analysis) for doc in cursor]
        else:
            where_clauses: List[str] = []
            params: List[Any] = []
            if poem_type and poem_type != "الكل":
                where_clauses.append("poem_type LIKE ?")
                params.append(f"%{poem_type.strip()}%")
            if poet.strip():
                where_clauses.append("poet LIKE ?")
                params.append(f"%{poet.strip()}%")
            if meter.strip():
                where_clauses.append("dominant_meter LIKE ?")
                params.append(f"%{meter.strip()}%")
            if region.strip():
                where_clauses.append("region LIKE ?")
                params.append(f"%{region.strip()}%")
            if verification_state.strip() in VERIFICATION_LABELS:
                where_clauses.append("verification_state = ?")
                params.append(verification_state.strip())

            if q.strip():
                q_norm = normalize_hassaniya(q, unify_gaf_to_kaf=flexible)
                col_field = "search_text_flex" if flexible else "search_text"
                where_clauses.append(f"({col_field} LIKE ? OR title LIKE ? OR poem_text LIKE ? OR poet LIKE ?)")
                like_norm = f"%{q_norm}%"
                like_raw = f"%{q.strip()}%"
                params.extend([like_norm, like_raw, like_raw, like_raw])

            where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.row_factory = sqlite3.Row
                total = conn.execute(f"SELECT COUNT(*) FROM poems {where_sql}", params).fetchone()[0]
                rows = conn.execute(
                    f"SELECT * FROM poems {where_sql} ORDER BY created_at DESC LIMIT ? OFFSET ?",
                    [*params, limit, offset],
                ).fetchall()
                items = []
                for row in rows:
                    d = dict(row)
                    d["analysis"] = json.loads(d.pop("analysis_json", "{}") or "{}") if include_analysis else {}
                    items.append(self._normalize_poem_doc(d, include_analysis=include_analysis))

        total_pages = max(1, math.ceil(total / limit)) if total else 1
        return {
            "items": items,
            "total": total,
            "page": page,
            "limit": limit,
            "total_pages": total_pages,
            "has_more": page < total_pages,
        }

    def iter_all_poems(self, q: str = "", poem_type: str = "الكل") -> Iterator[Dict]:
        """مكرّر خفيف على الذاكرة للتصدير (DOCX / JSONL) والعمليات الإدارية."""
        page = 1
        while True:
            batch = self.list_poems_paginated(q=q, poem_type=poem_type, page=page, limit=100, include_analysis=True)
            for item in batch["items"]:
                yield item
            if not batch["has_more"]:
                break
            page += 1

    def list_poem_ids_for_sitemap(self) -> List[str]:
        """يجلب المعرّفات فقط لخريطة الموقع بدون تحميل النصوص والتحليلات."""
        if self.active_mode == "mongodb":
            return [str(d.get("id") or d["_id"]) for d in self.mongo_db["poems"].find({}, {"id": 1}).sort("created_at", DESCENDING)]
        with sqlite3.connect(self.sqlite_path) as conn:
            return [row[0] for row in conn.execute("SELECT id FROM poems ORDER BY created_at DESC").fetchall()]

    def clean_existing_archive(self, local_eng: LocalHassaniyaEngine) -> Dict[str, int]:
        scanned = fixed_poets = fixed_meters = updated_total = 0
        for rec in self.iter_all_poems():
            scanned += 1
            old_title, old_poet = rec.get("title", ""), rec.get("poet", "")
            old_meter = str(rec.get("dominant_meter", ""))
            new_title, new_poet = clean_title_and_poet(old_title, old_poet)

            needs_poet_fix = (new_title != old_title) or (new_poet != old_poet)
            old_meter_typo = ("السغير" in old_meter) or ("البت لكبير" in old_meter)
            is_broken_meter = (
                "وزن تقديري" in old_meter
                or any(f"({k} متحرك)" in old_meter for k in (1, 2, 3, 4))
                or old_meter_typo
                or not old_meter.strip()
            )
            is_local = not str(rec.get("source_engine", "")).startswith("AI")
            new_analysis = rec.get("analysis") or {}
            new_meter, new_type = old_meter, rec.get("poem_type", "نص شعري")

            fixed_this_meter = False
            if is_broken_meter and is_local and rec.get("poem_text"):
                reanalyzed = local_eng.analyze(rec["poem_text"])
                new_analysis, new_meter, new_type = reanalyzed, reanalyzed["dominant_meter"], reanalyzed["poem_type"]
                fixed_meters += 1
                fixed_this_meter = True
            elif old_meter_typo:
                new_meter = old_meter.replace("السغير", "اسغير").replace("البت لكبير", "البت الكبير")
                if isinstance(new_analysis, dict):
                    new_analysis["dominant_meter"] = new_meter
                fixed_meters += 1
                fixed_this_meter = True

            if needs_poet_fix:
                fixed_poets += 1

            # تحديث حقول البحث المزدوجة في كل الأحوال إن كانت ناقصة
            s_strict, s_flex = build_search_fields(new_title, new_poet, rec.get("region", DEFAULT_REGION_LABEL), rec.get("poem_text", ""), new_meter)
            if needs_poet_fix or fixed_this_meter:
                updated_total += 1
                if self.active_mode == "mongodb":
                    self.mongo_db["poems"].update_one(
                        {"$or": [{"_id": rec["id"]}, {"id": rec["id"]}]},
                        {"$set": {
                            "title": new_title, "poet": new_poet, "dominant_meter": new_meter,
                            "poem_type": new_type, "analysis": new_analysis,
                            "search_text": s_strict, "search_text_flex": s_flex,
                        }},
                    )
                else:
                    with sqlite3.connect(self.sqlite_path) as conn:
                        conn.execute(
                            "UPDATE poems SET title=?, poet=?, poem_type=?, dominant_meter=?, analysis_json=?, "
                            "search_text=?, search_text_flex=? WHERE id=?",
                            (new_title, new_poet, new_type, new_meter, json.dumps(new_analysis, ensure_ascii=False), s_strict, s_flex, rec["id"]),
                        )
                        self._sync_sqlite_fts(conn, rec["id"], s_strict, s_flex)
                        conn.commit()

        self._invalidate_stats_cache()
        return {"scanned": scanned, "updated_total": updated_total, "fixed_poets_and_titles": fixed_poets, "fixed_broken_meters": fixed_meters}

    def get_filters(self) -> Dict[str, List[str]]:
        if self.active_mode == "mongodb":
            col = self.mongo_db["poems"]
            poets = sorted({str(v).strip() for v in col.distinct("poet") if v and not is_default_poet(str(v))})
            meters = sorted({str(v).strip() for v in col.distinct("dominant_meter") if v})
            regions = sorted({str(v).strip() for v in col.distinct("region") if v})
            types_ = sorted({str(v).strip() for v in col.distinct("poem_type") if v})
            return {"poets": poets, "meters": meters, "regions": regions, "poem_types": types_}
        with sqlite3.connect(self.sqlite_path) as conn:
            poets = sorted({r[0].strip() for r in conn.execute("SELECT DISTINCT poet FROM poems").fetchall() if r[0] and not is_default_poet(r[0])})
            meters = sorted({r[0].strip() for r in conn.execute("SELECT DISTINCT dominant_meter FROM poems").fetchall() if r[0]})
            regions = sorted({r[0].strip() for r in conn.execute("SELECT DISTINCT region FROM poems").fetchall() if r[0]})
            types_ = sorted({r[0].strip() for r in conn.execute("SELECT DISTINCT poem_type FROM poems").fetchall() if r[0]})
            return {"poets": poets, "meters": meters, "regions": regions, "poem_types": types_}

    def delete_poem(self, record_id: str) -> bool:
        if self.active_mode == "mongodb":
            deleted = self.mongo_db["poems"].delete_one({"$or": [{"_id": record_id}, {"id": record_id}]}).deleted_count > 0
            if deleted:
                try:
                    self.mongo_db[RAG_COLLECTION_NAME].delete_many({"doc_id": f"poem::{record_id}"})
                except Exception:
                    pass
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                cur = conn.execute("DELETE FROM poems WHERE id = ?", (record_id,))
                if self._sqlite_fts_enabled:
                    conn.execute("DELETE FROM poems_fts WHERE poem_id = ?", (record_id,))
                conn.commit()
                deleted = cur.rowcount > 0
        if deleted:
            self._invalidate_stats_cache()
        return deleted

    def delete_poems_bulk(self, record_ids: List[str]) -> int:
        clean_ids = [str(rid).strip() for rid in record_ids if str(rid).strip()]
        if not clean_ids:
            return 0
        if self.active_mode == "mongodb":
            res = self.mongo_db["poems"].delete_many({"$or": [{"_id": {"$in": clean_ids}}, {"id": {"$in": clean_ids}}]})
            try:
                self.mongo_db[RAG_COLLECTION_NAME].delete_many({"doc_id": {"$in": [f"poem::{rid}" for rid in clean_ids]}})
            except Exception:
                pass
            count = int(res.deleted_count)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                placeholders = ",".join("?" for _ in clean_ids)
                cur = conn.execute(f"DELETE FROM poems WHERE id IN ({placeholders})", clean_ids)
                if self._sqlite_fts_enabled:
                    conn.execute(f"DELETE FROM poems_fts WHERE poem_id IN ({placeholders})", clean_ids)
                conn.commit()
                count = int(cur.rowcount)
        self._invalidate_stats_cache()
        return count

    def delete_all_data(self, include_glossary: bool = False, include_rag: bool = True) -> Dict[str, int]:
        deleted_poems = deleted_glossary = deleted_rag = 0
        if self.active_mode == "mongodb":
            deleted_poems = int(self.mongo_db["poems"].delete_many({}).deleted_count)
            if include_glossary:
                deleted_glossary = int(self.mongo_db["glossary"].delete_many({}).deleted_count)
            if include_rag:
                try:
                    rag_filter = {} if include_glossary else {"doc_id": {"$regex": "^poem::"}}
                    deleted_rag = int(self.mongo_db[RAG_COLLECTION_NAME].delete_many(rag_filter).deleted_count)
                except Exception:
                    pass
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                deleted_poems = int(conn.execute("DELETE FROM poems").rowcount)
                if self._sqlite_fts_enabled:
                    conn.execute("DELETE FROM poems_fts")
                if include_glossary:
                    deleted_glossary = int(conn.execute("DELETE FROM glossary_senses").rowcount)
                conn.commit()
        self._invalidate_stats_cache()
        return {"deleted_poems": deleted_poems, "deleted_glossary": deleted_glossary, "deleted_rag": deleted_rag}

    def list_glossary_paginated(self, q: str = "", status: str = "", page: int = 1, limit: int = 50) -> Dict[str, Any]:
        offset = (page - 1) * limit
        if self.active_mode == "mongodb":
            query: Dict[str, Any] = {}
            if status in VERIFICATION_LABELS:
                query["status"] = status
            if q.strip():
                q_norm = normalize_hassaniya(q)
                query["$or"] = [
                    {"normalized_term": {"$regex": re.escape(q_norm)}},
                    {"term": {"$regex": re.escape(q.strip()), "$options": "i"}},
                    {"arabic_meaning": {"$regex": re.escape(q.strip()), "$options": "i"}},
                    {"english_meaning": {"$regex": re.escape(q.strip()), "$options": "i"}},
                    {"source_poem_title": {"$regex": re.escape(q.strip()), "$options": "i"}},
                ]
            col = self.mongo_db["glossary"]
            total = col.count_documents(query)
            items = []
            for doc in col.find(query).sort("updated_at", DESCENDING).skip(offset).limit(limit):
                sense_id = doc.get("sense_id") or str(doc.get("_id", ""))
                st = doc.get("status") or "machine_generated"
                items.append({
                    "sense_id": sense_id,
                    "term": doc.get("term", ""),
                    "arabic_meaning": doc.get("arabic_meaning", ""),
                    "english_meaning": doc.get("english_meaning", ""),
                    "source_poem_id": doc.get("source_poem_id", ""),
                    "source_poem_title": doc.get("source_poem_title", ""),
                    "attestation": doc.get("attestation", ""),
                    "variants": doc.get("variants", []),
                    "status": st,
                    "status_label": VERIFICATION_LABELS.get(st, VERIFICATION_LABELS["machine_generated"]),
                    "editor": doc.get("editor", "AI"),
                    "updated_at": doc.get("updated_at", ""),
                })
        else:
            where, params = [], []
            if status in VERIFICATION_LABELS:
                where.append("status = ?")
                params.append(status)
            if q.strip():
                q_norm = normalize_hassaniya(q)
                where.append("(normalized_term LIKE ? OR term LIKE ? OR arabic_meaning LIKE ? OR english_meaning LIKE ? OR source_poem_title LIKE ?)")
                raw_like = f"%{q.strip()}%"
                params.extend([f"%{q_norm}%", raw_like, raw_like, raw_like, raw_like])
            where_sql = f"WHERE {' AND '.join(where)}" if where else ""
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.row_factory = sqlite3.Row
                total = conn.execute(f"SELECT COUNT(*) FROM glossary_senses {where_sql}", params).fetchone()[0]
                rows = conn.execute(
                    f"SELECT * FROM glossary_senses {where_sql} ORDER BY updated_at DESC LIMIT ? OFFSET ?",
                    [*params, limit, offset],
                ).fetchall()
                items = []
                for r in rows:
                    st = r["status"] or "machine_generated"
                    items.append({
                        "sense_id": r["sense_id"],
                        "term": r["term"],
                        "arabic_meaning": r["arabic_meaning"],
                        "english_meaning": r["english_meaning"],
                        "source_poem_id": r["source_poem_id"],
                        "source_poem_title": r["source_poem_title"],
                        "attestation": r["attestation"],
                        "variants": json.loads(r["variants_json"] or "[]"),
                        "status": st,
                        "status_label": VERIFICATION_LABELS.get(st, VERIFICATION_LABELS["machine_generated"]),
                        "editor": r["editor"],
                        "updated_at": r["updated_at"],
                    })

        total_pages = max(1, math.ceil(total / limit)) if total else 1
        return {"items": items, "total": total, "page": page, "limit": limit, "total_pages": total_pages, "has_more": page < total_pages}

    def create_glossary_sense(self, payload: GlossarySenseCreateRequest) -> str:
        sense_id = new_record_id(10)
        entry = {
            "sense_id": sense_id,
            "term": payload.term.strip(),
            "arabic_meaning": payload.arabic_meaning.strip(),
            "english_meaning": payload.english_meaning.strip(),
            "source_poem_id": payload.source_poem_id.strip(),
            "source_poem_title": payload.source_poem_title.strip(),
            "attestation": payload.attestation.strip(),
            "variants": [v.strip() for v in payload.variants if v.strip()],
            "status": payload.status,
            "editor": payload.editor.strip() or "هيئة تحرير ALBIDAN",
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        self._upsert_glossary_senses([entry], allow_overwrite_verified=True)
        self._invalidate_stats_cache()
        return sense_id

    def update_glossary_sense(self, sense_id: str, payload: GlossarySenseUpdateRequest) -> bool:
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M")
        updates: Dict[str, Any] = {"updated_at": now_iso}
        if payload.term is not None:
            updates["term"] = payload.term.strip()
            updates["normalized_term"] = normalize_hassaniya(payload.term)
        if payload.arabic_meaning is not None:
            updates["arabic_meaning"] = payload.arabic_meaning.strip()
        if payload.english_meaning is not None:
            updates["english_meaning"] = payload.english_meaning.strip()
        if payload.attestation is not None:
            updates["attestation"] = payload.attestation.strip()
        if payload.variants is not None:
            updates["variants"] = [v.strip() for v in payload.variants if v.strip()]
        if payload.status is not None:
            updates["status"] = payload.status
        if payload.editor is not None:
            updates["editor"] = payload.editor.strip()

        if self.active_mode == "mongodb":
            ok = self.mongo_db["glossary"].update_one({"$or": [{"sense_id": sense_id}, {"_id": sense_id}]}, {"$set": updates}).matched_count > 0
        else:
            sql_updates = dict(updates)
            if "variants" in sql_updates:
                sql_updates["variants_json"] = json.dumps(sql_updates.pop("variants"), ensure_ascii=False)
            cols = ", ".join(f"{k}=?" for k in sql_updates.keys())
            with sqlite3.connect(self.sqlite_path) as conn:
                cur = conn.execute(f"UPDATE glossary_senses SET {cols} WHERE sense_id=?", [*sql_updates.values(), sense_id])
                conn.commit()
                ok = cur.rowcount > 0
        if ok:
            self._invalidate_stats_cache()
        return ok

    def delete_glossary_term_or_sense(self, target: str) -> bool:
        """يدعم الحذف بالمعرّف sense_id أو بالمفردة term للحفاظ على توافق الواجهة الحالية."""
        if self.active_mode == "mongodb":
            ok = self.mongo_db["glossary"].delete_one({"$or": [{"sense_id": target}, {"_id": target}, {"term": target}]}).deleted_count > 0
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                cur = conn.execute("DELETE FROM glossary_senses WHERE sense_id = ? OR term = ?", (target, target))
                conn.commit()
                ok = cur.rowcount > 0
        if ok:
            self._invalidate_stats_cache()
        return ok

    def save_message(self, payload: ContactMessageRequest) -> str:
        msg_id = new_record_id(10)
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M")
        doc = {
            "id": msg_id, "name": payload.name.strip(), "contact": payload.contact.strip(),
            "category": payload.category.strip() or "تواصل بحثي",
            "message": payload.message.strip(), "created_at": now_iso,
        }
        if self.active_mode == "mongodb":
            self.mongo_db["messages"].insert_one(dict(doc, _id=msg_id))
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.execute(
                    "INSERT INTO messages (id, name, contact, category, message, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                    (doc["id"], doc["name"], doc["contact"], doc["category"], doc["message"], doc["created_at"]),
                )
                conn.commit()
        self._invalidate_stats_cache()
        return msg_id

    def list_messages(self) -> List[Dict]:
        if self.active_mode == "mongodb":
            items = []
            for doc in self.mongo_db["messages"].find().sort("created_at", DESCENDING).limit(200):
                doc["id"] = doc.get("id") or str(doc["_id"])
                doc.pop("_id", None)
                items.append(doc)
            return items
        with sqlite3.connect(self.sqlite_path) as conn:
            conn.row_factory = sqlite3.Row
            return [dict(r) for r in conn.execute("SELECT * FROM messages ORDER BY created_at DESC LIMIT 200").fetchall()]

    def delete_message(self, msg_id: str) -> bool:
        if self.active_mode == "mongodb":
            ok = self.mongo_db["messages"].delete_one({"$or": [{"_id": msg_id}, {"id": msg_id}]}).deleted_count > 0
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                cur = conn.execute("DELETE FROM messages WHERE id = ?", (msg_id,))
                conn.commit()
                ok = cur.rowcount > 0
        if ok:
            self._invalidate_stats_cache()
        return ok

    def get_stats(self) -> Dict:
        now = time.time()
        if self._stats_cache is not None and (now - self._stats_cached_at) < 30:
            return self._stats_cache

        if self.active_mode == "mongodb":
            poems_col = self.mongo_db["poems"]
            poems_count = poems_col.count_documents({})
            verified_count = poems_col.count_documents({"verification_state": "verified"})
            terms_count = self.mongo_db["glossary"].count_documents({})
            messages_count = self.mongo_db["messages"].count_documents({})
            total_lines = sum(
                len([l for l in str(doc.get("poem_text", "")).split("\n") if l.strip()])
                for doc in poems_col.find({}, {"poem_text": 1})
            )
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                poems_count = conn.execute("SELECT COUNT(*) FROM poems").fetchone()[0]
                verified_count = conn.execute("SELECT COUNT(*) FROM poems WHERE verification_state = 'verified'").fetchone()[0]
                terms_count = conn.execute("SELECT COUNT(*) FROM glossary_senses").fetchone()[0]
                messages_count = conn.execute("SELECT COUNT(*) FROM messages").fetchone()[0]
                total_lines = sum(
                    len([l for l in str(r[0] or "").split("\n") if l.strip()])
                    for r in conn.execute("SELECT poem_text FROM poems").fetchall()
                )

        res = {
            "poems_count": poems_count,
            "verified_poems_count": verified_count,
            "lines_count": total_lines,
            "terms_count": terms_count,
            "messages_count": messages_count,
            "db_mode": self.active_mode,
            "db_status": self.status_message,
            "ai_ready": bool(GEMINI_API_KEY),
            "default_model": DEFAULT_GEMINI_MODEL,
            "version": APP_VERSION,
        }
        self._stats_cache = res
        self._stats_cached_at = now
        return res


# =============================================================================
# 7. خدمات Gemini AI (التحليل العروضي، التفريغ، OCR، وتقسيم الدواوين)
# =============================================================================
def get_candidate_models(preferred: str) -> List[str]:
    ordered = [preferred] if preferred and preferred in ALLOWED_MODELS else []
    for m in FALLBACK_MODELS:
        if m not in ordered:
            ordered.append(m)
    return ordered


def generate_content_safe(client: genai.Client, preferred_model: str, contents, config=None) -> Tuple[object, str]:
    last_exc = None
    for model in get_candidate_models(preferred_model):
        try:
            kwargs = {"model": model, "contents": contents}
            if config is not None:
                kwargs["config"] = config
            return client.models.generate_content(**kwargs), model
        except Exception as exc:
            logger.error("Gemini call failed on model '%s': %s", model, exc)
            last_exc = exc
    raise last_exc if last_exc else RuntimeError("تعذر الاتصال بخدمة Gemini.")


def run_gemini_on_media_bytes(raw_bytes: bytes, mime_type: str, suffix: str, prompt: str, preferred_model: str) -> Tuple[str, str]:
    client = genai.Client(api_key=GEMINI_API_KEY)
    if len(raw_bytes) <= FILES_API_THRESHOLD:
        resp, used_model = generate_content_safe(client, preferred_model, [types.Part.from_bytes(data=raw_bytes, mime_type=mime_type), prompt])
        return (resp.text or "").strip(), used_model
    tmp_path, uploaded = None, None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix or ".bin") as tmp:
            tmp.write(raw_bytes)
            tmp_path = tmp.name
        uploaded = client.files.upload(file=tmp_path, config={"mime_type": mime_type})
        for _ in range(25):
            if "PROCESSING" not in str(getattr(uploaded, "state", "")).upper():
                break
            time.sleep(1)
            uploaded = client.files.get(name=uploaded.name)
        if "FAILED" in str(getattr(uploaded, "state", "")).upper():
            raise RuntimeError("تعذرت معالجة الملف على خادم الذكاء الاصطناعي، جرّب صيغة MP3 أو PDF أصغر.")
        resp, used_model = generate_content_safe(client, preferred_model, [uploaded, prompt])
        return (resp.text or "").strip(), used_model
    finally:
        if uploaded is not None:
            try:
                client.files.delete(name=uploaded.name)
            except Exception:
                pass
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def analyze_with_gemini(poem_text: str, model_name: str, local_eng: LocalHassaniyaEngine) -> Tuple[Dict, str]:
    client = genai.Client(api_key=GEMINI_API_KEY)
    sys_inst = """أنت باحث عروضي متخصص في الشعر الشعبي الحساني («لغن») وبحوره («لَبْتُوتَة»، ومفردها «بَتْ»). التزم بدقة بالمنهجية التالية:
1. اضبط كل «تافلويت» بالشكل الصوتي الحساني الكامل (الفتحة والضمة والكسرة والسكون)، مع استخدام حرف (گ) للقاف المعقودة.
2. أساس الوزن في «لغن»: عدّ الحروف المتحركة في «التافلويت» مع إهمال السواكن؛ مع مراعاة بنية البحر وأدائه (الظّهر).
3. وحدات «لغن»: الطلعة (3 تيفلواتن بروي واحد)، الگاف/الكاف (4 أو 6 أو 8 تيفلواتن)، التبراع (تافلويتين).
4. استخرج أهم المفردات الحسانية الأصيلة مع شرحها بالعربية الفصحى والإنجليزية.
5. قدّم ترجمة أدبية دقيقة للإنجليزية والفرنسية."""
    cfg = types.GenerateContentConfig(
        system_instruction=sys_inst,
        response_mime_type="application/json",
        response_schema=AIPoemAnalysis,
    )
    # محاولة ثانية تلقائية في حال أعاد النموذج JSON غير صالح
    last_err = None
    for _ in range(2):
        resp, used_model = generate_content_safe(client, model_name, f"حلّل هذا النص الشعري الحساني تحليلاً عروضياً ولغوياً كاملاً:\n\n{poem_text}", cfg)
        try:
            parsed = AIPoemAnalysis.model_validate_json(resp.text).model_dump()
            return local_eng.cross_check_ai_analysis(parsed), used_model
        except Exception as exc:
            logger.warning("Invalid JSON from Gemini, retrying once: %s", exc)
            last_err = exc
    raise last_err or RuntimeError("تعذر قراءة مخرجات JSON من Gemini.")


def split_poems_with_gemini(raw_text: str, model_name: str, default_poet: str, default_region: str) -> Tuple[List[Dict], str]:
    client = genai.Client(api_key=GEMINI_API_KEY)
    sys_inst = """أنت باحث متخصص في تحرير وفهرسة دواوين الشعر الحساني. افصل كل قصيدة في عنصر مستقل بدقة:
1. استخرج عنوان كل قصيدة نظيفاً بدون إلحاق اسم الشاعر بالعنوان.
2. استخرج اسم الشاعر أو الراوي وحده في حقل poet، وإن لم يُذكر اتركه فارغاً.
3. حافظ على النص الشعري الحساني حرفياً وضع كل شطر (تافلويت) في سطر مستقل."""
    cfg = types.GenerateContentConfig(
        system_instruction=sys_inst,
        response_mime_type="application/json",
        response_schema=BatchSplitResult,
    )
    resp, used_model = generate_content_safe(client, model_name, f"افصل جميع القصائد الحسانية في هذا النص إلى قائمة مهيكلة:\n\n{raw_text}", cfg)
    parsed = BatchSplitResult.model_validate_json(resp.text).model_dump()
    result = []
    for idx, p in enumerate(parsed.get("poems", []), start=1):
        txt = (p.get("poem_text") or "").strip()
        if txt:
            c_title, c_poet = clean_title_and_poet(
                (p.get("title") or "").strip() or f"قصيدة {idx}",
                (p.get("poet") or "").strip() or default_poet.strip() or DEFAULT_POET_LABEL,
            )
            result.append({
                "title": c_title,
                "poet": c_poet,
                "region": default_region.strip() or DEFAULT_REGION_LABEL,
                "poem_text": txt,
            })
    return result, used_model


def transcribe_audio_sync(audio_bytes: bytes, mime_type: str, suffix: str, model_name: str) -> Tuple[str, str]:
    prompt = "أنت خبير في اللهجة الحسانية والشعر الحساني. فرّغ هذا التسجيل الصوتي نصياً بدقة: اكتب الكلمات كما تُنطق بالحسانية مستخدماً حرف (گ) للقاف المعقودة، ضع كل شطر شعري (تافلويت) في سطر مستقل، وأخرج الأبيات الشعرية فقط دون عناوين أو تعليقات."
    return run_gemini_on_media_bytes(audio_bytes, mime_type, suffix, prompt, model_name)


def ocr_images_sync(images_data: List[Tuple[bytes, str]], model_name: str) -> Tuple[str, str]:
    client = genai.Client(api_key=GEMINI_API_KEY)
    prompt = "أنت خبير في رقمنة المخطوطات والكتب والأشعار الحسانية (OCR). اقرأ النص الشعري الحساني في هذه الصور واستخرجه بدقة: حافظ على الألفاظ الحسانية وحرف (گ) للقاف المعقودة، ضع كل شطر في سطر مستقل، وافصل بين القصائد بسطر فارغ، وأخرج النصوص الشعرية فقط بدون شروحات."
    parts = [types.Part.from_bytes(data=raw, mime_type=mime) for raw, mime in images_data]
    resp, used_model = generate_content_safe(client, model_name, [*parts, prompt])
    return (resp.text or "").strip(), used_model


def extract_pdf_sync(raw: bytes, model_name: str) -> Tuple[str, str]:
    extracted = ""
    try:
        extracted = "\n".join((page.extract_text() or "").strip() for page in PdfReader(io.BytesIO(raw)).pages).strip()
    except Exception:
        extracted = ""
    if len(extracted) >= 15:
        return extracted, "pdf-text"
    if not GEMINI_API_KEY:
        raise ValueError("ملف PDF مصوّر ويتطلب تفعيل مفتاح GEMINI_API_KEY.")
    prompt = "اقرأ جميع الأبيات والقصائد الحسانية في ملف PDF هذا واستخرجها بدقة مع الحفاظ على اللهجة الحسانية وحرف (گ)، وضع كل شطر في سطر مستقل وافصل بين القصائد بسطر فارغ، وأخرج النصوص فقط."
    text, used_model = run_gemini_on_media_bytes(raw, "application/pdf", ".pdf", prompt, model_name)
    return text, f"pdf-ocr ({used_model})"


# =============================================================================
# 8. تطبيق FastAPI وترويسات الأمان
# =============================================================================
app = FastAPI(title="ALBIDAN AI Lab — Academic Hassaniya Corpus API", version=APP_VERSION)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allow_headers=["Authorization", "Content-Type", "X-Admin-Pin"],
)


@app.middleware("http")
async def security_headers_middleware(request: Request, call_next):
    response = await call_next(request)
    response.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault(
        "Content-Security-Policy",
        (
            "default-src 'self'; "
            "script-src 'self' 'unsafe-inline' https://unpkg.com https://cdn.jsdelivr.net; "
            "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
            "font-src 'self' data: https://fonts.gstatic.com; "
            "img-src 'self' data: blob: https:; "
            "media-src 'self' data: blob: https:; "
            "connect-src 'self' data: blob:;"
        ),
    )
    return response


try:
    from hassaniya_rag_router import router as hassaniya_rag_router
    app.include_router(hassaniya_rag_router)
except ImportError:
    pass

archive = ArchiveManager(MONGODB_URI, MONGODB_DB_NAME)
local_engine = LocalHassaniyaEngine()
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
def serve_home():
    index_file = STATIC_DIR / "index.html"
    if not index_file.exists():
        raise HTTPException(status_code=404, detail="ملف الواجهة static/index.html غير موجود.")
    return HTMLResponse(content=index_file.read_text(encoding="utf-8").replace("__BASE_URL__", get_base_url()))


# -----------------------------------------------------------------------------
# صفحات SEO ومحركات البحث بالذكاء الاصطناعي (SSR Poems + Sitemap + llms.txt)
# -----------------------------------------------------------------------------
@app.get("/poem/{record_id}", response_class=HTMLResponse)
def serve_poem_ssr(record_id: str):
    rec = archive.get_poem(record_id)
    if not rec:
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    base_url = get_base_url()
    canonical = f"{base_url}/poem/{quote(rec['id'])}"
    analysis = rec.get("analysis") or {}
    title = rec.get("title") or "قصيدة حسانية"
    poet = rec.get("poet") or DEFAULT_POET_LABEL
    region = rec.get("region") or DEFAULT_REGION_LABEL
    meter = rec.get("dominant_meter") or "—"
    ptype = rec.get("poem_type") or "شعر حساني"
    poem_text = rec.get("poem_text") or ""
    v_status = rec.get("verification_status") or VERIFICATION_LABELS["machine_generated"]
    reviewer_info = f" • المراجع: {rec['reviewed_by']} ({rec['reviewed_at']})" if rec.get("reviewed_by") else ""
    excerpt = " / ".join([l.strip() for l in poem_text.split("\n") if l.strip()][:2])

    schema_obj = {
        "@context": "https://schema.org",
        "@type": "CreativeWork",
        "identifier": rec["id"],
        "name": title,
        "inLanguage": ["mey", "ar"],
        "genre": ptype,
        "spatialCoverage": region,
        "text": poem_text,
        "url": canonical,
        "publisher": {"@type": "Organization", "name": "ALBIDAN AI Lab — مختبر البيظان"},
    }
    if not is_default_poet(poet):
        schema_obj["author"] = {"@type": "Person", "name": poet}

    en_trans = analysis.get("english_translation") or ""
    fr_trans = analysis.get("french_translation") or ""
    t_label = "مراجَعة بشرياً" if analysis.get("translation_status") == "human_reviewed" else "ترجمة آلية (AI)"
    context = analysis.get("cultural_context") or ""

    page_html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{html.escape(title)} — {html.escape(poet)} | مختبر البيظان (ALBIDAN)</title>
  <meta name="description" content="{html.escape(f'{title} — {poet} ({ptype}، {meter}): {excerpt}')}" />
  <link rel="canonical" href="{canonical}" />
  <meta property="og:type" content="article" />
  <meta property="og:title" content="{html.escape(title)} — {html.escape(poet)}" />
  <meta property="og:description" content="{html.escape(excerpt)}" />
  <meta property="og:url" content="{canonical}" />
  <link href="https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=IBM+Plex+Sans+Arabic:wght@400;600;700&display=swap" rel="stylesheet" />
  <script type="application/ld+json">{safe_json_ld(schema_obj)}</script>
  <style>
    body{{margin:0;background:#0D0B09;color:#F4EFE6;font-family:'IBM Plex Sans Arabic',sans-serif;line-height:1.8;padding:1.5rem 1rem}}
    .wrap{{max-width:780px;margin:0 auto;background:#1A1612;border:1px solid rgba(216,178,84,.25);border-top:3px solid #D8B254;border-radius:16px;padding:1.8rem}}
    a{{color:#EBD07E;text-decoration:none}}
    .meta{{color:#A69C8E;font-size:.88rem;margin:.4rem 0 .8rem}}
    .badge{{display:inline-block;font-size:.78rem;padding:2px 10px;border-radius:999px;background:rgba(216,178,84,.14);color:#EBD07E;border:1px solid rgba(216,178,84,.35);margin-bottom:.8rem}}
    .poem{{font-family:'Amiri',serif;font-size:1.55rem;line-height:2.25;white-space:pre-line;text-align:center;background:#151310;border:1px solid rgba(216,178,84,.2);border-radius:12px;padding:1.4rem;margin:1rem 0}}
    .box{{background:#151310;border-right:3px solid #D8B254;border-radius:10px;padding:1rem;margin-top:1rem;font-size:.92rem}}
    .ltr{{direction:ltr;text-align:left;white-space:pre-line}}
    .actions{{display:flex;gap:8px;flex-wrap:wrap;margin-top:1.4rem}}
    .btn{{padding:.55rem 1rem;border-radius:9px;background:#D8B254;color:#14100C;font-weight:700;font-size:.86rem}}
    .btn-sec{{background:#151310;color:#F4EFE6;border:1px solid rgba(216,178,84,.3)}}
  </style>
</head>
<body>
  <article class="wrap">
    <div style="font-size:.82rem"><a href="{base_url}/">← العودة إلى ديوان ومختبر البيظان (ALBIDAN)</a></div>
    <h1 style="margin:.6rem 0 .2rem;font-family:'Amiri',serif;font-size:2rem;color:#EBD07E">{html.escape(title)}</h1>
    <div class="meta">المعرّف: <code>{html.escape(rec['id'])}</code> • الشاعر / الراوي: <strong>{html.escape(poet)}</strong> • المنطقة: {html.escape(region)} • الوحدة: {html.escape(ptype)} • البَتْ: {html.escape(meter)}</div>
    <div class="badge">حالة التوثيق: {html.escape(v_status)}{html.escape(reviewer_info)}</div>
    <div class="poem">{html.escape(poem_text)}</div>
    {f'<div class="box"><strong>السياق الثقافي والغرض:</strong><br>{html.escape(context)}</div>' if context else ''}
    {f'<div class="box ltr"><strong>English Translation ({html.escape(t_label)}):</strong><br>{html.escape(en_trans)}</div>' if en_trans else ''}
    {f'<div class="box ltr"><strong>Traduction Française ({html.escape(t_label)}):</strong><br>{html.escape(fr_trans)}</div>' if fr_trans else ''}
    <div class="actions">
      <a class="btn" href="{base_url}/">تصفح الديوان والمختبر التفاعلي</a>
      <a class="btn btn-sec" href="{base_url}/api/export/poem/{quote(rec['id'])}.docx">تنزيل الوثيقة (Word)</a>
    </div>
  </article>
</body>
</html>"""
    return HTMLResponse(content=page_html)


@app.get("/sitemap.xml")
def serve_sitemap():
    base_url = get_base_url()
    urls = [f"<url><loc>{base_url}/</loc><changefreq>daily</changefreq><priority>1.0</priority></url>"]
    for pid in archive.list_poem_ids_for_sitemap():
        urls.append(f"<url><loc>{base_url}/poem/{quote(pid)}</loc><changefreq>weekly</changefreq><priority>0.8</priority></url>")
    xml = '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' + "\n".join(urls) + "\n</urlset>"
    return Response(content=xml, media_type="application/xml")


@app.get("/robots.txt", response_class=PlainTextResponse)
def serve_robots():
    base_url = get_base_url()
    return (
        "User-agent: *\nAllow: /\nDisallow: /api/admin/\nDisallow: /api/messages\n\n"
        f"Sitemap: {base_url}/sitemap.xml\n"
    )


@app.get("/llms.txt", response_class=PlainTextResponse)
def serve_llms_txt():
    base_url = get_base_url()
    stats = archive.get_stats()
    return f"""# ALBIDAN — Digital Corpus & Prosody Lab of Hassaniya Poetry (مختبر وديوان البيظان)

> Open academic platform dedicated to Hassaniya Arabic (ISO 639-3: `mey`) poetry. Founded by Aggar Lahbib (أڭار الحبيب).

## Live Corpus Statistics
- Documented Poems: {stats['poems_count']} (Human-Verified: {stats['verified_poems_count']})
- Prosodic Hemistichs (Taflowit): {stats['lines_count']}
- Glossary Senses: {stats['terms_count']}

## Public Endpoints
- Web Portal: {base_url}/
- OpenAPI Docs: {base_url}/docs
- Paginated Poems API: {base_url}/api/poems?page=1&limit=24
- Paginated Glossary API: {base_url}/api/glossary?page=1&limit=50
- Streaming JSONL Export: {base_url}/api/export/jsonl
- Sitemap: {base_url}/sitemap.xml
"""


# -----------------------------------------------------------------------------
# مسارات الـ API التشغيلية والأكاديمية
# -----------------------------------------------------------------------------
@app.get("/health")
def health_check():
    return {"status": "ok", "db_mode": archive.active_mode, "version": APP_VERSION}


@app.get("/api/stats")
def api_stats():
    return archive.get_stats()


@app.get("/api/filters")
def api_filters():
    return archive.get_filters()


@app.post("/api/admin/verify")
def api_admin_verify(request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    return {"authorized": True}


@app.post("/api/admin/clean-archive")
def api_admin_clean_archive(request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    report = archive.clean_existing_archive(local_engine)
    return {
        **report,
        "message": (
            f"تم فحص {report['scanned']} قصيدة: إصلاح {report['fixed_poets_and_titles']} عنوان/شاعر "
            f"وإصلاح {report['fixed_broken_meters']} وزن عروضي."
        ),
    }


@app.post("/api/admin/enrich-glossary")
async def api_admin_enrich_glossary(
    req: EnrichGlossaryRequest,
    request: Request,
    x_admin_pin: Optional[str] = Header(default=None),
):
    require_admin(x_admin_pin, request)
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=400, detail="مفتاح GEMINI_API_KEY غير مضبوط على الخادم.")
    model = req.model_name if req.model_name in ALLOWED_MODELS else DEFAULT_GEMINI_MODEL
    candidates = []
    for p in archive.iter_all_poems():
        if (not str(p.get("source_engine", "")).startswith("AI") or not (p.get("analysis") or {}).get("glossary")) and (p.get("poem_text") or "").strip():
            candidates.append(p)
            if len(candidates) >= req.limit:
                break

    enriched, failed, added_terms = 0, 0, 0
    for rec in candidates:
        try:
            analysis, used_model = await run_in_threadpool(analyze_with_gemini, rec["poem_text"], model, local_engine)
            archive.update_poem_analysis(rec["id"], analysis, f"AI ({used_model})")
            enriched += 1
            added_terms += len(analysis.get("glossary") or [])
        except Exception as exc:
            logger.error("Failed enriching poem %s: %s", rec.get("id"), exc)
            failed += 1

    return {
        "processed": len(candidates),
        "enriched": enriched,
        "failed": failed,
        "terms_extracted": added_terms,
        "message": f"تم إثراء {enriched} قصيدة واستخراج {added_terms} مفردة للمعجم (فشل: {failed}).",
    }


@app.post("/api/analyze")
async def api_analyze(req: AnalyzeRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    if not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "ai_analyze", max_requests=20, window_seconds=3600)
    text = req.poem_text.strip()
    if not text:
        raise HTTPException(status_code=400, detail="يرجى إدخال بيت شعري واحد على الأقل.")
    model = req.model_name if req.model_name in ALLOWED_MODELS else DEFAULT_GEMINI_MODEL
    warning = None
    if req.use_ai and GEMINI_API_KEY:
        try:
            analysis, used_model = await run_in_threadpool(analyze_with_gemini, text, model, local_engine)
            source_engine = f"AI ({used_model})"
            if analysis.get("deterministic_mismatches", 0) > 0:
                warning = f"تنبيه عروضي: وُجد اختلاف في عدّ المتحركات في {analysis['deterministic_mismatches']} شطر؛ يُنصح بالمراجعة البشرية."
        except Exception as exc:
            logger.error("Analyze with Gemini failed: %s", exc)
            analysis, source_engine = local_engine.analyze(text), "Local (Fallback)"
            warning = f"تعذر الاتصال بـ Gemini ({exc})، تم العرض بالمحرك المحلي."
    else:
        analysis, source_engine = local_engine.analyze(text), "Local"
    return {
        "source_engine": source_engine,
        "verification_state": "machine_generated",
        "verification_status": VERIFICATION_LABELS["machine_generated"],
        "warning": warning,
        "analysis": analysis,
    }


@app.post("/api/transcribe")
async def api_transcribe(
    request: Request,
    file: UploadFile = File(...),
    model_name: str = Form(DEFAULT_GEMINI_MODEL),
    x_admin_pin: Optional[str] = Header(default=None),
):
    if not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "ai_media", max_requests=10, window_seconds=3600)
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=400, detail="مفتاح GEMINI_API_KEY غير مضبوط في الخادم.")
    audio_bytes = await file.read()
    if not audio_bytes:
        raise HTTPException(status_code=400, detail="الملف الصوتي فارغ.")
    if len(audio_bytes) > 20 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="حجم الصوت يتجاوز 20 ميجابايت.")
    ext = Path(file.filename or "audio.mp3").suffix.lower()
    mime_type = AUDIO_MIME_MAP.get(ext, file.content_type or "audio/mp3")
    if mime_type == "application/octet-stream":
        mime_type = "audio/mp3"
    try:
        text, used_model = await run_in_threadpool(transcribe_audio_sync, audio_bytes, mime_type, ext or ".mp3", model_name)
    except Exception as exc:
        logger.error("Audio transcription failed: %s", exc)
        raise HTTPException(status_code=500, detail=f"تعذر تفريغ الصوت: {exc}")
    finally:
        del audio_bytes
    if not text:
        raise HTTPException(status_code=422, detail="لم يتم التعرف على كلام واضح في هذا التسجيل.")
    return {"text": text, "model": used_model}


@app.post("/api/ocr")
async def api_ocr_images(
    request: Request,
    files: List[UploadFile] = File(...),
    model_name: str = Form(DEFAULT_GEMINI_MODEL),
    x_admin_pin: Optional[str] = Header(default=None),
):
    if not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "ai_media", max_requests=12, window_seconds=3600)
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=400, detail="يتطلب استخراج النص من الصور تفعيل مفتاح GEMINI_API_KEY.")
    if not files or len(files) > 6:
        raise HTTPException(status_code=400, detail="يمكن رفع من صورة واحدة إلى 6 صور في المرة الواحدة.")
    images_data, total_size = [], 0
    for f in files:
        raw = await f.read()
        if not raw:
            continue
        total_size += len(raw)
        if total_size > 20 * 1024 * 1024:
            raise HTTPException(status_code=400, detail="إجمالي حجم الصور كبير جداً (الحد الأقصى 20 ميجابايت).")
        ext = Path(f.filename or "img.jpg").suffix.lower()
        images_data.append((raw, IMAGE_MIME_MAP.get(ext, f.content_type or "image/jpeg")))
    if not images_data:
        raise HTTPException(status_code=400, detail="لم يتم استلام صور صالحة.")
    try:
        text, used_model = await run_in_threadpool(ocr_images_sync, images_data, model_name)
    except Exception as exc:
        logger.error("Image OCR failed: %s", exc)
        raise HTTPException(status_code=500, detail=f"تعذر قراءة الصورة: {exc}")
    finally:
        del images_data
    return {"text": text, "model": used_model}


@app.post("/api/import-file")
async def api_import_file(
    request: Request,
    file: UploadFile = File(...),
    model_name: str = Form(DEFAULT_GEMINI_MODEL),
    x_admin_pin: Optional[str] = Header(default=None),
):
    if not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "ai_media", max_requests=15, window_seconds=3600)
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="الملف المرفوع فارغ.")
    if len(raw) > 20 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="حجم الملف يتجاوز 20 ميجابايت.")
    ext = Path(file.filename or "").suffix.lower()
    try:
        if ext == ".docx":
            return {"text": await run_in_threadpool(extract_docx_plain_text, raw), "method": "docx"}
        if ext == ".doc":
            raise HTTPException(status_code=400, detail="صيغة Word القديمة (.doc) غير مدعومة؛ احفظ الملف بصيغة .docx.")
        if ext == ".pdf":
            text, method = await run_in_threadpool(extract_pdf_sync, raw, model_name)
            return {"text": text, "method": method}
        if ext == ".txt":
            return {"text": raw.decode("utf-8", errors="ignore").strip(), "method": "txt"}
        raise HTTPException(status_code=400, detail="الصيغ المدعومة هي: Word (.docx) و PDF (.pdf) و TXT.")
    except HTTPException:
        raise
    except Exception as exc:
        logger.error("File import failed: %s", exc)
        raise HTTPException(status_code=500, detail=f"تعذر استيراد الملف: {exc}")
    finally:
        del raw


@app.post("/api/import-batch-docx")
async def api_import_batch_docx(
    request: Request,
    file: UploadFile = File(...),
    use_ai_split: bool = Form(True),
    default_poet: str = Form(""),
    default_region: str = Form(DEFAULT_REGION_LABEL),
    model_name: str = Form(DEFAULT_GEMINI_MODEL),
    x_admin_pin: Optional[str] = Header(default=None),
):
    if PROTECT_PUBLIC_WRITES:
        require_admin(x_admin_pin, request)
    elif not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "batch_parse", max_requests=8, window_seconds=3600)

    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="الملف فارغ.")
    if len(raw) > 25 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="حجم الملف يتجاوز 25 ميجابايت.")
    ext = Path(file.filename or "").suffix.lower()
    if ext == ".doc":
        raise HTTPException(status_code=400, detail="احفظ الملف بصيغة (.docx) الحديثة بدلاً من (.doc).")
    if ext not in (".docx", ".txt"):
        raise HTTPException(status_code=400, detail="يدعم الاستيراد الجماعي ملفات Word (.docx) أو TXT فقط.")
    try:
        plain_text = await run_in_threadpool(extract_docx_plain_text, raw) if ext == ".docx" else raw.decode("utf-8", errors="ignore").strip()
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"تعذر قراءة الملف: {exc}")
    if not plain_text.strip():
        raise HTTPException(status_code=400, detail="الملف فارغ ولا يحتوي على نصوص.")

    warning, items, method = None, [], "structural"
    if use_ai_split and GEMINI_API_KEY and len(plain_text) <= 35000:
        try:
            items, used_model = await run_in_threadpool(split_poems_with_gemini, plain_text, model_name, default_poet, default_region)
            method = f"ai ({used_model})"
        except Exception as exc:
            logger.error("Batch AI split failed: %s", exc)
            warning = f"تعذر التقسيم الذكي ({exc})، تم التحويل للتقسيم الهيكلي السريع."
    if not items:
        if ext == ".docx":
            items = await run_in_threadpool(split_docx_structurally, raw, default_poet, default_region)
        else:
            blocks = [b.strip() for b in re.split(r"\n\s*\n", plain_text) if b.strip()]
            items = []
            for i, b in enumerate(blocks, start=1):
                c_title, c_poet = clean_title_and_poet(f"قصيدة {i}", default_poet.strip() or DEFAULT_POET_LABEL)
                items.append({"title": c_title, "poet": c_poet, "region": default_region.strip() or DEFAULT_REGION_LABEL, "poem_text": b})
        if use_ai_split and len(plain_text) > 35000:
            warning = "الملف طويل جداً، تم استخدام التقسيم الهيكلي السريع لضمان عدم فقدان أي قصيدة."
    del raw
    return {"count": len(items), "method": method, "warning": warning, "items": items}


# -----------------------------------------------------------------------------
# مسارات القصائد بترقيم الصفحات والتوثيق الأكاديمي
# -----------------------------------------------------------------------------
@app.get("/api/poems")
def api_list_poems(
    q: str = Query("", max_length=200),
    poem_type: str = Query("الكل", max_length=80),
    poet: str = Query("", max_length=120),
    meter: str = Query("", max_length=120),
    region: str = Query("", max_length=120),
    verification_state: str = Query("", max_length=40),
    flexible: bool = Query(False, description="وضع البحث المرن (يدمج ق/گ/ك)"),
    page: int = Query(1, ge=1),
    limit: int = Query(24, ge=1, le=200),
    include_analysis: bool = Query(True),
):
    return archive.list_poems_paginated(
        q=q,
        poem_type=poem_type,
        poet=poet,
        meter=meter,
        region=region,
        verification_state=verification_state,
        flexible=flexible,
        page=page,
        limit=limit,
        include_analysis=include_analysis,
    )


@app.get("/api/poems/{record_id}")
def api_get_single_poem(record_id: str):
    rec = archive.get_poem(record_id)
    if not rec:
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    return rec


@app.post("/api/poems")
def api_save_poem(req: SavePoemRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    if PROTECT_PUBLIC_WRITES:
        require_admin(x_admin_pin, request)
    elif not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "save_poem", max_requests=25, window_seconds=3600)
    if not req.poem_text.strip():
        raise HTTPException(status_code=400, detail="نص القصيدة فارغ.")
    return {"id": archive.save_poem(req), "message": "تم التوثيق في المدونة وتحديث المعجم بنجاح."}


@app.post("/api/poems/batch")
def api_save_poems_batch(req: BatchSaveRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    if PROTECT_PUBLIC_WRITES:
        require_admin(x_admin_pin, request)
    elif not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "save_batch", max_requests=10, window_seconds=3600)
    if not req.items:
        raise HTTPException(status_code=400, detail="لم يتم تحديد أي قصيدة للحفظ.")
    prepared = [
        SavePoemRequest(
            title=item.title, poet=item.poet, region=item.region,
            poem_text=item.poem_text.strip(),
            analysis=item.analysis or local_engine.analyze(item.poem_text.strip()),
            source_engine=item.source_engine,
        )
        for item in req.items if item.poem_text.strip()
    ]
    if not prepared:
        raise HTTPException(status_code=400, detail="جميع القصائد المحددة فارغة.")
    ids = archive.save_poems_batch(prepared)
    return {"saved_count": len(ids), "ids": ids, "message": f"تم حفظ {len(ids)} قصيدة دفعة واحدة بنجاح!"}


@app.post("/api/poems/bulk-delete")
def api_bulk_delete_poems(req: BulkDeletePoemsRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    if not req.ids:
        raise HTTPException(status_code=400, detail="لم يتم تحديد أي قصيدة للحذف.")
    deleted_count = archive.delete_poems_bulk(req.ids)
    return {"deleted_count": deleted_count, "message": f"تم حذف {deleted_count} قصيدة دفعة واحدة بنجاح."}


@app.post("/api/poems/delete-all")
def api_delete_all_poems(req: DeleteAllArchiveRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    stats = archive.delete_all_data(include_glossary=req.include_glossary, include_rag=req.include_rag)
    return {
        **stats,
        "message": (
            f"تم الحذف الشامل بنجاح (القصائد: {stats['deleted_poems']}، "
            f"المعجم: {stats['deleted_glossary']}، مقاطع البحث: {stats['deleted_rag']})."
        ),
    }


@app.put("/api/poems/{record_id}")
def api_update_poem(record_id: str, req: UpdatePoemRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    if not req.poem_text.strip():
        raise HTTPException(status_code=400, detail="نص القصيدة لا يمكن أن يكون فارغاً.")
    if not archive.update_poem(record_id, req):
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    return {"updated": True, "id": record_id}


@app.patch("/api/poems/{record_id}/verify")
def api_verify_poem(record_id: str, req: VerifyPoemRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    """مسار التحكيم والتوثيق البشري للقصيدة (مولّد آلياً -> قيد المراجعة -> موثّق ومحكّم)."""
    require_admin(x_admin_pin, request)
    updated = archive.verify_poem(record_id, req)
    if not updated:
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    return {"verified": True, "poem": updated, "message": f"تم تحديث حالة التوثيق إلى: {updated['verification_status']}"}


@app.delete("/api/poems/{record_id}")
def api_delete_poem(record_id: str, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    if not archive.delete_poem(record_id):
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    return {"deleted": True, "id": record_id}


# -----------------------------------------------------------------------------
# مسارات المعجم متعدد المعاني (Glossary Senses)
# -----------------------------------------------------------------------------
@app.get("/api/glossary")
def api_list_glossary(
    q: str = Query("", max_length=200),
    status: str = Query("", max_length=40),
    page: int = Query(1, ge=1),
    limit: int = Query(50, ge=1, le=200),
):
    return archive.list_glossary_paginated(q=q, status=status, page=page, limit=limit)


@app.post("/api/glossary")
def api_create_glossary_sense(req: GlossarySenseCreateRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    sense_id = archive.create_glossary_sense(req)
    return {"sense_id": sense_id, "message": "تمت إضافة المعنى المعجمي الموثّق بنجاح."}


@app.put("/api/glossary/{sense_id}")
def api_update_glossary_sense(sense_id: str, req: GlossarySenseUpdateRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    if not archive.update_glossary_sense(sense_id, req):
        raise HTTPException(status_code=404, detail="المدخل المعجمي غير موجود.")
    return {"updated": True, "sense_id": sense_id}


@app.delete("/api/glossary/{target}")
def api_delete_glossary(target: str, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    if not archive.delete_glossary_term_or_sense(target):
        raise HTTPException(status_code=404, detail="المفردة غير موجودة.")
    return {"deleted": True, "target": target}


# -----------------------------------------------------------------------------
# الرسائل والتصدير (Streaming JSONL + DOCX)
# -----------------------------------------------------------------------------
@app.post("/api/messages")
def api_save_message(req: ContactMessageRequest, request: Request):
    rate_limiter.check_limit(request, "messages", max_requests=8, window_seconds=3600)
    if not req.name.strip() or not req.message.strip():
        raise HTTPException(status_code=400, detail="يرجى كتابة الاسم والرسالة.")
    return {"id": archive.save_message(req), "message": "تم استلام مساهمتك بنجاح! شكراً لإثرائك مختبر البيظان."}


@app.get("/api/messages")
def api_list_messages(request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    return {"items": archive.list_messages()}


@app.delete("/api/messages/{msg_id}")
def api_delete_message(msg_id: str, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    if not archive.delete_message(msg_id):
        raise HTTPException(status_code=404, detail="الرسالة غير موجودة.")
    return {"deleted": True, "id": msg_id}


@app.get("/api/export/glossary.docx")
def api_export_glossary_docx(request: Request, q: str = Query("", max_length=200), x_admin_pin: Optional[str] = Header(default=None)):
    if not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "export_docx", max_requests=20, window_seconds=3600)
    items = archive.list_glossary_paginated(q=q, page=1, limit=2000)["items"]
    doc = create_base_arabic_doc("المعجم الحساني المتنامي", f"تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d')}  |  عدد المعاني الموثقة: {len(items)}")
    tbl = doc.add_table(rows=1, cols=5)
    tbl.style = "Table Grid"
    set_table_rtl(tbl)
    for i, h in enumerate(["المفردة الحسانية", "المعنى بالعربية الفصحى", "English Meaning", "المصدر", "حالة التوثيق"]):
        cell = tbl.rows[0].cells[i]
        set_cell_shading(cell, "1B365D")
        p = cell.paragraphs[0]
        set_paragraph_rtl(p)
        set_run_arabic(p.add_run(h), size_pt=10, bold=True, color_hex="FFFFFF")
    for g in items:
        cells = tbl.add_row().cells
        for i, v in enumerate([g.get("term", ""), g.get("arabic_meaning", ""), g.get("english_meaning", ""), g.get("source_poem_title", ""), g.get("status_label", "")]):
            p = cells[i].paragraphs[0]
            set_paragraph_rtl(p)
            set_run_arabic(p.add_run(v), size_pt=10, bold=(i == 0))
    return doc_to_stream(doc, "albidan_hassaniya_glossary.docx")


@app.post("/api/export/analysis.docx")
def api_export_analysis_docx(req: ExportAnalysisDocxRequest, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    if not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "export_docx", max_requests=20, window_seconds=3600)
    if len(json.dumps(req.analysis, ensure_ascii=False).encode("utf-8")) > MAX_ANALYSIS_JSON_BYTES:
        raise HTTPException(status_code=413, detail="حجم بيانات التحليل المرسلة للتصدير يتجاوز الحد المسموح.")
    doc = create_base_arabic_doc("تقرير التحليل العروضي واللساني", f"تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d')}")
    append_poem_section_to_doc(doc, req.title, req.poet, req.region, req.poem_text, req.analysis)
    return doc_to_stream(doc, "albidan_poem_analysis.docx")


@app.get("/api/export/poem/{record_id}.docx")
def api_export_single_poem_docx(record_id: str, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    if not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "export_docx", max_requests=40, window_seconds=3600)
    rec = archive.get_poem(record_id)
    if not rec:
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    doc = create_base_arabic_doc(
        "وثيقة شعرية حسانية",
        f"معرّف التوثيق: {rec['id']}  |  الحالة: {rec.get('verification_status', '')}  |  تاريخ الأرشفة: {rec.get('created_at', '')}",
    )
    append_poem_section_to_doc(doc, rec["title"], rec["poet"], rec["region"], rec["poem_text"], rec.get("analysis", {}))
    return doc_to_stream(doc, f"albidan_poem_{rec['id']}.docx")


@app.get("/api/export/poems.docx")
def api_export_all_poems_docx(
    request: Request,
    q: str = Query("", max_length=200),
    poem_type: str = Query("الكل", max_length=80),
    x_admin_pin: Optional[str] = Header(default=None),
):
    if not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "export_docx", max_requests=10, window_seconds=3600)
    poems = list(archive.iter_all_poems(q=q, poem_type=poem_type))
    doc = create_base_arabic_doc("المدونة الحسانية المفتوحة — ديوان النصوص الموثقة", f"تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d')}  |  عدد القصائد: {len(poems)}")
    for idx, rec in enumerate(poems):
        append_poem_section_to_doc(doc, rec["title"], rec["poet"], rec["region"], rec["poem_text"], rec.get("analysis", {}))
        if idx < len(poems) - 1:
            doc.add_page_break()
    return doc_to_stream(doc, "albidan_hassaniya_anthology.docx")


@app.get("/api/export/jsonl")
def api_export_jsonl(request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    if not is_valid_admin_pin(x_admin_pin):
        rate_limiter.check_limit(request, "export_jsonl", max_requests=15, window_seconds=3600)

    def stream_lines() -> Iterator[bytes]:
        for p in archive.iter_all_poems():
            yield (json.dumps(p, ensure_ascii=False) + "\n").encode("utf-8")

    return StreamingResponse(
        stream_lines(),
        media_type="application/jsonl",
        headers={"Content-Disposition": "attachment; filename=albidan_hassaniya_corpus.jsonl"},
    )
