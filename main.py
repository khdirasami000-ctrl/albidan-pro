import os
import json
from typing import Optional
from fastapi import FastAPI, Request, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, EmailStr, Field

# =====================================================================
# 1. الإعدادات الثابتة (تُقرأ من متغيرات البيئة في Render)
# =====================================================================
# النطاق الرسمي الثابت لمنع تزوير ترويسة Host
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "https://albidan.pro").rstrip("/")

# إغلاق الكتابة العامة افتراضياً (لا تُفتح إلا إذا وضعت false يدوياً)
PROTECT_PUBLIC_WRITES = os.getenv("PROTECT_PUBLIC_WRITES", "true").lower() == "true"

# قائمة النطاقات المسموح لها بالتواصل مع الخادم فقط (CORS)
ALLOWED_ORIGINS = [
    origin.strip()
    for origin in os.getenv(
        "ALLOWED_ORIGINS",
        "https://albidan.pro,https://www.albidan.pro"
    ).split(",")
    if origin.strip()
]

app = FastAPI(title="ALBIDAN API")

# =====================================================================
# 2. تفعيل حماية CORS (لائحة الضيوف المسموح لهم)
# =====================================================================
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE"],
    allow_headers=["Authorization", "Content-Type", "X-Admin-Pin"],
)

# =====================================================================
# 3. ترويسات الأمان الإجبارية لكل استجابة (حماية المتصفح)
# =====================================================================
@app.middleware("http")
async def add_security_headers(request: Request, call_next):
    response = await call_next(request)
    # إجبار المتصفح على استخدام HTTPS دائماً
    response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    # منع المتصفح من تخمين نوع الملفات بشكل خاطئ
    response.headers["X-Content-Type-Options"] = "nosniff"
    # منع عرض موقعك داخل iframe في مواقع أخرى (منع Clickjacking)
    response.headers["X-Frame-Options"] = "DENY"
    # حماية خصوصية الروابط المرجعية
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    # سياسة أمان المحتوى (CSP) لتحديد مصادر السكربتات والخطوط المسموحة
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' https://unpkg.com; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src 'self' https://fonts.gstatic.com; "
        "img-src 'self' data: https:; "
        "connect-src 'self';"
    )
    return response

# =====================================================================
# 4. الدوال المساعدة الآمنة (الرابط الثابت + مصفاة XSS)
# =====================================================================
def get_base_url() -> str:
    """ترجع الرابط الرسمي الثابت دائماً دون الاعتماد على طلب المستخدم."""
    return PUBLIC_BASE_URL

def safe_json_ld(data: dict) -> str:
    """
    مصفاة XSS: تحوّل أي علامة < أو > داخل النص إلى رمز آمن
    حتى لا يتمكن أي نص خبيث من كسر وسم <script> في صفحة القصيدة.
    """
    return (
        json.dumps(data, ensure_ascii=False)
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("&", "\\u0026")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )

# =====================================================================
# 5. نماذج المدخلات بحدود صارمة (منع إغراق الخادم)
# =====================================================================
class ContactMessageRequest(BaseModel):
    name: str = Field(..., min_length=2, max_length=100, description="اسم المرسل")
    email: Optional[EmailStr] = Field(default=None, description="البريد الإلكتروني")
    subject: str = Field(default="رسالة عامة", max_length=150, description="الموضوع")
    message: str = Field(..., min_length=10, max_length=2000, description="نص الرسالة")

class ExportDocxRequest(BaseModel):
    title: str = Field(..., min_length=1, max_length=200, description="عنوان القصيدة")
    poem_text: str = Field(..., min_length=5, max_length=15000, description="نص القصيدة")
    analysis_text: str = Field(default="", max_length=30000, description="نص التحليل")

# =====================================================================
# 6. أمثلة تطبيقية لكيفية استخدام الدوال داخل مساراتك
# =====================================================================
@app.get("/poem/{poem_id}", response_class=HTMLResponse)
async def render_poem_page(poem_id: str):
    # مثال لبيانات القصيدة القادمة من قاعدة البيانات
    base_url = get_base_url()
    schema_obj = {
        "@context": "https://schema.org",
        "@type": "CreativeWork",
        "genre": "Hassani Poetry",
        "url": f"{base_url}/poem/{poem_id}",
        "name": "عنوان القصيدة",
    }

    # نستخدم safe_json_ld بدلاً من json.dumps مباشرة
    safe_schema = safe_json_ld(schema_obj)

    html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <link rel="canonical" href="{base_url}/poem/{poem_id}" />
    <script type="application/ld+json">{safe_schema}</script>
</head>
<body>
    <h1>صفحة القصيدة</h1>
</body>
</html>"""
    return HTMLResponse(content=html)

@app.post("/api/contact")
async def submit_contact(payload: ContactMessageRequest):
    # إذا تجاوزت الرسالة 2000 حرف، يرفضها FastAPI تلقائياً بخطأ 422 قبل الوصول إلى هنا
    return {"status": "received", "sender": payload.name}

@app.post("/api/export/analysis.docx")
async def export_analysis_docx(payload: ExportDocxRequest):
    # النص هنا مضمون ألا يتجاوز الحدود المسموحة (15 ألف حرف للقصيدة، 30 ألف للتحليل)
    return {"status": "ready_to_export", "title": payload.title}
import os
import re
import io
import hmac
import time
import json
import uuid
import html
import sqlite3
import tempfile
from collections import defaultdict, Counter
from pathlib import Path
from datetime import datetime
from urllib.parse import quote
from typing import List, Dict, Optional, Tuple

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

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
STATIC_DIR.mkdir(exist_ok=True)

MONGODB_URI = os.getenv("MONGODB_URI", "").strip()
MONGODB_DB_NAME = os.getenv("MONGODB_DB_NAME", "albidan_corpus").strip()
RAG_COLLECTION_NAME = os.getenv("RAG_COLLECTION_NAME", "hassaniya_rag_chunks").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
DEFAULT_GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip() or "gemini-3.5-flash-lite"
ADMIN_PIN = os.getenv("ADMIN_PIN", "").strip()
PROTECT_PUBLIC_WRITES = os.getenv("PROTECT_PUBLIC_WRITES", "false").strip().lower() == "true"

ALLOWED_MODELS = {
    "gemini-3.5-flash-lite",
    "gemini-3.8-flash",
    "gemini-3.1-pro-preview",
}
FALLBACK_MODELS = ["gemini-3.5-flash-lite", "gemini-3.8-flash"]

AUDIO_MIME_MAP = {
    ".mp3": "audio/mp3", ".wav": "audio/wav", ".m4a": "audio/mp4", ".mp4": "audio/mp4",
    ".aac": "audio/aac", ".ogg": "audio/ogg", ".oga": "audio/ogg", ".opus": "audio/ogg",
    ".webm": "audio/webm", ".flac": "audio/flac",
}
IMAGE_MIME_MAP = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}
FILES_API_THRESHOLD = 2 * 1024 * 1024

DEFAULT_POET_LABEL = "تراث شفهي / غير معروف"
DEFAULT_REGION_LABEL = "الصحراء المغربية"


# =============================================================================
# 1. التطبيع اللغوي الحساني وتنظيف العناوين وأسماء الشعراء
# =============================================================================
ARABIC_DIACRITICS_RE = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u0640]")
POET_EMBED_PATTERNS = [
    re.compile(r"\s*[—–\-|•]\s*(?:الشاعر\s*/\s*الراوي|الشاعر|الراوي|قائلها|نظم)\s*[:：]\s*(.+)$"),
    re.compile(r"\s*\((?:الشاعر\s*/\s*الراوي|الشاعر|الراوي)\s*[:：]?\s*([^)]+)\)\s*$"),
    re.compile(r"\s+للشاعر\s+(.+)$"),
]
REDUNDANT_POET_PREFIX_RE = re.compile(r"^(?:الشاعر\s*/\s*الراوي|الشاعر|الراوي)\s*[:：]?\s*")


def normalize_hassaniya(text: str, unify_gaf_to_kaf: bool = False) -> str:
    """تطبيع النص الحساني لأغراض البحث والفهرسة مع الحفاظ على النص الأصلي للعرض."""
    if not text:
        return ""
    t = ARABIC_DIACRITICS_RE.sub("", str(text))
    replacements = {
        "أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا",
        "ى": "ي", "ة": "ه",
        "ݣ": "گ", "ڭ": "گ", "ڨ": "گ", "ڤ": "گ",
    }
    for src, dst in replacements.items():
        t = t.replace(src, dst)
    if unify_gaf_to_kaf:
        t = t.replace("گ", "ك").replace("ق", "ك")
    return re.sub(r"\s+", " ", t).strip().lower()


def is_default_poet(poet: str) -> bool:
    p = (poet or "").strip()
    return not p or p in {
        DEFAULT_POET_LABEL, "غير معروف", "تراث شفهي", "مجهول", "—", "-"
    }


def clean_poet_name(poet: str) -> str:
    p = (poet or "").strip()
    if not p:
        return DEFAULT_POET_LABEL
    # إزالة التكرار مثل "الشاعر/الراوي: الشاعر حمداتي..." -> "حمداتي..."
    for _ in range(3):
        cleaned = REDUNDANT_POET_PREFIX_RE.sub("", p).strip()
        if cleaned == p:
            break
        p = cleaned
    p = re.sub(r"^[:：\-—–\s]+|[:：\-—–\s]+$", "", p).strip()
    return p or DEFAULT_POET_LABEL


def clean_title_and_poet(title: str, poet: str) -> Tuple[str, str]:
    """يفصل اسم الشاعر المدمج خطأً داخل العنوان وينظف الحقلين معاً."""
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

    final_title = raw_title or "قصيدة بدون عنوان"
    return final_title, final_poet


def build_poem_search_text(title: str, poet: str, region: str, poem_text: str, dominant_meter: str) -> str:
    joined = f"{title} {poet} {region} {dominant_meter} {poem_text}"
    return f"{normalize_hassaniya(joined)} {normalize_hassaniya(joined, unify_gaf_to_kaf=True)}"


def get_base_url(request: Request) -> str:
    base_url = str(request.base_url).rstrip("/")
    if base_url.startswith("http://") and "localhost" not in base_url and "127.0.0.1" not in base_url:
        base_url = base_url.replace("http://", "https://", 1)
    return base_url


# =============================================================================
# 2. حماية الطلبات (Rate Limiting) ومكافحة تخمين رمز المشرف
# =============================================================================
class SimpleRateLimiter:
    def __init__(self):
        self._hits: Dict[str, List[float]] = defaultdict(list)
        self._admin_failures: Dict[str, List[float]] = defaultdict(list)

    def _client_ip(self, request: Request) -> str:
        fwd = request.headers.get("x-forwarded-for", "")
        if fwd:
            return fwd.split(",")[0].strip()
        return request.client.host if request.client else "unknown"

    def check_limit(self, request: Request, bucket: str, max_requests: int, window_seconds: int):
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
            detail="لم يتم ضبط متغير البيئة ADMIN_PIN على الخادم؛ يرجى إضافته في إعدادات Render لتفعيل صلاحيات المشرف.",
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
    poem_type: str = Field(description="وحدة/جنس «لغن»: طلعة (3 تيفلواتن برويّ واحد)، گاف/كاف (4 أو 6 أو 8 تيفلواتن)، تبراع (تافلويتين)، طلعة وگاف، أو نص موزون")
    dominant_meter: str = Field(description="البَتْ (البحر الحساني من لَبْتُوتَة) السائد في القصيدة مثل: التيدوم، البت الكبير، بوعمران، المزارگ، المصارع، إملخ، اسغير، الواكدي، أحويويص، الدگسري، المشكل، إمريميدة، لبتيت التام، لبتيت الناقص، السروزي، الرسم، لبير، بغزز، بت اثنين، بت اثلاثة، حثو اجراد")
    dominant_rhyme: str = Field(description="بنية القافية والروي العام في القطعة")
    zahr_mode: str = Field(default="", description="الظّهر (المقام أو النغم الموسيقي المرافق لأداء البحر في «لغن» إن أمكن تحديده)")
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
    poem_text: str
    model_name: str = DEFAULT_GEMINI_MODEL
    use_ai: bool = True


class SavePoemRequest(BaseModel):
    title: str
    poet: str = DEFAULT_POET_LABEL
    region: str = DEFAULT_REGION_LABEL
    poem_text: str
    analysis: Dict
    source_engine: str = "AI"


class BatchSaveItem(BaseModel):
    title: str = ""
    poet: str = DEFAULT_POET_LABEL
    region: str = DEFAULT_REGION_LABEL
    poem_text: str
    analysis: Optional[Dict] = None
    source_engine: str = "Local (Batch Word)"


class BatchSaveRequest(BaseModel):
    items: List[BatchSaveItem]


class BulkDeletePoemsRequest(BaseModel):
    ids: List[str] = Field(default_factory=list)


class DeleteAllArchiveRequest(BaseModel):
    include_glossary: bool = False
    include_rag: bool = True


class UpdatePoemRequest(BaseModel):
    title: str
    poet: str
    region: str
    poem_text: str


class ExportAnalysisDocxRequest(BaseModel):
    title: str = "تقرير تحليل عروضي حساني"
    poet: str = DEFAULT_POET_LABEL
    region: str = DEFAULT_REGION_LABEL
    poem_text: str
    analysis: Dict


class ContactMessageRequest(BaseModel):
    name: str
    contact: str = ""
    category: str = "تواصل بحثي"
    message: str


class EnrichGlossaryRequest(BaseModel):
    limit: int = Field(default=10, ge=1, le=25)
    model_name: str = DEFAULT_GEMINI_MODEL


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
            for i, v in enumerate([row_item.get("original_line", ""), row_item.get("phonetic_diacritized", ""), str(row_item.get("mutaharrikat_count", "")), row_item.get("estimated_meter", ""), row_item.get("rhyme_ending", "")]):
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
# 5. المحرك العروضي المحلي المحسّن (يحل خطأ 1 أو 3 متحركات في النصوص جزئية التشكيل)
# =============================================================================
class LocalHassaniyaEngine:
    short_vowels = set("َُِ")
    tanween = set("ًٌٍ")
    long_vowels = set("اويى")
    arabic_letters_re = re.compile(r"[\u0621-\u064Aگݣڭڨڤ]")
    meters_map = {
        5: "اسغير / إملخ (5-6 متحركات)",
        6: "اسغير / لبير (6 متحركات)",
        7: "لبتيت الناقص / بوعمران (7 متحركات)",
        8: "لبتيت التام (8 متحركات)",
        9: "البت الكبير (9 متحركات)",
        10: "البت الكبير (10 متحركات)",
        11: "البت الكبير (11 متحرك)",
        12: "البت الكبير (12 متحرك)",
    }

    def _estimate_line_mutaharrikat(self, line: str) -> int:
        norm = re.sub(r"[ݣڭڨڤ]", "گ", line)
        letters = self.arabic_letters_re.findall(norm)
        consonants = [ch for ch in letters if ch not in self.long_vowels]
        total_consonants = max(1, len(consonants))

        vowel_marks = sum(1 for ch in norm if ch in self.short_vowels or ch in self.tanween)
        # لا نعتمد عدّ الحركات إلا إذا كان الشطر مشكولاً تشكيلاً كثيفاً (وليس حركة أو حركتين عارضتين)
        if vowel_marks >= 5 and (vowel_marks / total_consonants) >= 0.38:
            return vowel_marks

        # تقدير صوتي حساني للنصوص غير المشكولة أو جزئية التشكيل (مع مراعاة الابتداء بالساكن)
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
                # في الحسانية يسقط متحرك الابتداء غالباً وتُسكن نهايات الكلمات
                est += max(1.5, c_count * 0.62)

        return max(4, min(14, int(round(est))))

    def _meter_label(self, count: int) -> str:
        if count in self.meters_map:
            return self.meters_map[count]
        if count > 12:
            return "البت الكبير (ممتد)"
        return "قيد التدقيق العروضي"

    def analyze(self, poem_text: str) -> Dict:
        lines = [l.strip() for l in poem_text.strip().split("\n") if l.strip()]
        details = []
        counts = []

        for line in lines:
            norm = re.sub(r"[ݣڭڨڤ]", "گ", line)
            count = self._estimate_line_mutaharrikat(norm)
            counts.append(count)
            clean = re.sub(r"[\u0610-\u061A\u064B-\u065F\u0670\u0640\s\W]", "", norm)
            rhyme = clean[-2:] if len(clean) >= 2 else (clean or "—")
            details.append({
                "original_line": line,
                "phonetic_diacritized": "تحليل أولي — افتح القصيدة في المختبر للتشكيل الصوتي الكامل",
                "mutaharrikat_count": count,
                "estimated_meter": self._meter_label(count),
                "rhyme_ending": rhyme,
            })

        if counts:
            sorted_counts = sorted(counts)
            median_count = sorted_counts[len(sorted_counts) // 2]
            dominant_meter = self._meter_label(median_count)
        else:
            dominant_meter = "قيد التدقيق العروضي"

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
            "zahr_mode": "يُحدَّد بالتحليل الأكاديمي (الظّهر)",
            "cultural_context": "قصيدة محفوظة بالتقطيع العروضي الأولي؛ افتحها في المختبر للتحليل الأكاديمي الكامل وتحديد البَتْ والظّهر والمفردات.",
            "linguistic_notes": "تقدير محلي لعدّ المتحركات في كل «تافلويت» مع إهمال السواكن وفق قواعد «لغن»؛ يُنصح بالتدقيق عبر محرك Gemini.",
            "lines": details,
            "glossary": [],
            "english_translation": "",
            "french_translation": "",
        }


# =============================================================================
# 6. مدير قواعد البيانات والأرشيف (MongoDB Atlas + SQLite + Indexes + BulkWrite)
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

        if mongo_uri and PYMONGO_AVAILABLE:
            try:
                client = MongoClient(mongo_uri, serverSelectionTimeoutMS=5000)
                client.admin.command("ping")
                self.mongo_db = client[db_name]
                self.active_mode = "mongodb"
                self.status_message = f"متصل سحابياً بـ MongoDB Atlas ({db_name})"
                self._ensure_mongo_indexes()
            except Exception as exc:
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
            self.mongo_db["poems"].create_index([("created_at", DESCENDING)])
            self.mongo_db["poems"].create_index([("poet", ASCENDING), ("dominant_meter", ASCENDING)])
            self.mongo_db["poems"].create_index([("poem_type", ASCENDING)])
            self.mongo_db["glossary"].create_index([("term", ASCENDING)])
            self.mongo_db["glossary"].create_index([("normalized_term", ASCENDING)])
            self.mongo_db["messages"].create_index([("created_at", DESCENDING)])
        except Exception:
            pass

    def _init_sqlite(self):
        with sqlite3.connect(self.sqlite_path) as conn:
            conn.execute(
                "CREATE TABLE IF NOT EXISTS poems ("
                "id TEXT PRIMARY KEY, title TEXT NOT NULL, poet TEXT, region TEXT, "
                "poem_text TEXT NOT NULL, poem_type TEXT, dominant_meter TEXT, "
                "azawan_mode TEXT, source_engine TEXT, analysis_json TEXT NOT NULL, "
                "music_json TEXT NOT NULL, created_at TEXT NOT NULL)"
            )
            conn.execute(
                "CREATE TABLE IF NOT EXISTS glossary ("
                "term TEXT PRIMARY KEY, arabic_meaning TEXT, english_meaning TEXT, "
                "source_poem_title TEXT, updated_at TEXT NOT NULL)"
            )
            conn.execute(
                "CREATE TABLE IF NOT EXISTS messages ("
                "id TEXT PRIMARY KEY, name TEXT NOT NULL, contact TEXT, category TEXT, "
                "message TEXT NOT NULL, created_at TEXT NOT NULL)"
            )
            conn.commit()

    def _prepare_poem_record(self, payload: SavePoemRequest, fallback_title: str = "قصيدة بدون عنوان") -> Dict:
        record_id = str(uuid.uuid4())[:8]
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M")
        clean_title, clean_poet = clean_title_and_poet(payload.title or fallback_title, payload.poet)
        region = (payload.region or "").strip() or DEFAULT_REGION_LABEL
        poem_text = payload.poem_text.strip()
        analysis = payload.analysis or {}
        dominant_meter = analysis.get("dominant_meter", "قيد التدقيق العروضي")
        poem_type = analysis.get("poem_type", "غير محدد")
        verification_status = "تحليل أكاديمي آلي" if str(payload.source_engine).startswith("AI") else "قيد التدقيق العروضي"

        return {
            "id": record_id,
            "title": clean_title,
            "poet": clean_poet,
            "region": region,
            "poem_text": poem_text,
            "poem_type": poem_type,
            "dominant_meter": dominant_meter,
            "source_engine": payload.source_engine,
            "verification_status": verification_status,
            "search_text": build_poem_search_text(clean_title, clean_poet, region, poem_text, dominant_meter),
            "analysis": analysis,
            "created_at": now_iso,
        }

    def _upsert_glossary_items_mongo(self, glossary_entries: List[Dict]):
        if not glossary_entries or self.mongo_db is None:
            return
        ops = []
        for g in glossary_entries:
            term = str(g.get("term", "")).strip()
            if not term:
                continue
            norm_term = normalize_hassaniya(term)
            update_doc: Dict = {
                "$set": {
                    "term": term,
                    "normalized_term": norm_term,
                    "arabic_meaning": g.get("arabic_meaning", ""),
                    "english_meaning": g.get("english_meaning", ""),
                    "source_poem_title": g.get("source_poem_title", "قصيدة موثّقة"),
                    "source_poem_id": g.get("source_poem_id", ""),
                    "updated_at": g.get("updated_at", datetime.now().strftime("%Y-%m-%d %H:%M")),
                }
            }
            if g.get("source_poem_id"):
                update_doc["$addToSet"] = {
                    "sources": {
                        "poem_id": g["source_poem_id"],
                        "poem_title": g.get("source_poem_title", ""),
                    }
                }
            ops.append(UpdateOne({"term": term}, update_doc, upsert=True))
        if ops:
            self.mongo_db["glossary"].bulk_write(ops, ordered=False)

    def save_poem(self, payload: SavePoemRequest) -> str:
        record = self._prepare_poem_record(payload)
        record_id = record["id"]
        now_iso = record["created_at"]
        glossary_items = [
            {
                "term": str(g.get("term", "")).strip(),
                "arabic_meaning": g.get("arabic_meaning", ""),
                "english_meaning": g.get("english_meaning", ""),
                "source_poem_title": record["title"],
                "source_poem_id": record_id,
                "updated_at": now_iso,
            }
            for g in (record["analysis"].get("glossary") or [])
            if str(g.get("term", "")).strip()
        ]

        if self.active_mode == "mongodb":
            self.mongo_db["poems"].insert_one(dict(record, _id=record_id))
            self._upsert_glossary_items_mongo(glossary_items)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.execute(
                    "INSERT INTO poems (id, title, poet, region, poem_text, poem_type, dominant_meter, azawan_mode, source_engine, analysis_json, music_json, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (
                        record["id"], record["title"], record["poet"], record["region"],
                        record["poem_text"], record["poem_type"], record["dominant_meter"],
                        "", record["source_engine"], json.dumps(record["analysis"], ensure_ascii=False),
                        "{}", record["created_at"],
                    ),
                )
                for g in glossary_items:
                    conn.execute(
                        "INSERT OR REPLACE INTO glossary (term, arabic_meaning, english_meaning, source_poem_title, updated_at) VALUES (?, ?, ?, ?, ?)",
                        (g["term"], g["arabic_meaning"], g["english_meaning"], g["source_poem_title"], now_iso),
                    )
                conn.commit()
        self._invalidate_stats_cache()
        return record_id

    def save_poems_batch(self, items: List[SavePoemRequest]) -> List[str]:
        saved_ids, docs_to_insert, glossary_updates = [], [], []
        for idx, payload in enumerate(items, start=1):
            if not payload.poem_text.strip():
                continue
            record = self._prepare_poem_record(payload, fallback_title=f"قصيدة {idx}")
            saved_ids.append(record["id"])
            docs_to_insert.append(record)
            for g in (record["analysis"].get("glossary") or []):
                term = str(g.get("term", "")).strip()
                if term:
                    glossary_updates.append({
                        "term": term,
                        "arabic_meaning": g.get("arabic_meaning", ""),
                        "english_meaning": g.get("english_meaning", ""),
                        "source_poem_title": record["title"],
                        "source_poem_id": record["id"],
                        "updated_at": record["created_at"],
                    })

        if not docs_to_insert:
            return []

        if self.active_mode == "mongodb":
            self.mongo_db["poems"].insert_many([dict(d, _id=d["id"]) for d in docs_to_insert])
            self._upsert_glossary_items_mongo(glossary_updates)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.executemany(
                    "INSERT INTO poems (id, title, poet, region, poem_text, poem_type, dominant_meter, azawan_mode, source_engine, analysis_json, music_json, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    [
                        (
                            d["id"], d["title"], d["poet"], d["region"], d["poem_text"],
                            d["poem_type"], d["dominant_meter"], "", d["source_engine"],
                            json.dumps(d["analysis"], ensure_ascii=False), "{}", d["created_at"],
                        )
                        for d in docs_to_insert
                    ],
                )
                if glossary_updates:
                    conn.executemany(
                        "INSERT OR REPLACE INTO glossary (term, arabic_meaning, english_meaning, source_poem_title, updated_at) VALUES (?, ?, ?, ?, ?)",
                        [(g["term"], g["arabic_meaning"], g["english_meaning"], g["source_poem_title"], g["updated_at"]) for g in glossary_updates],
                    )
                conn.commit()
        self._invalidate_stats_cache()
        return saved_ids

    def get_poem(self, record_id: str) -> Optional[Dict]:
        if self.active_mode == "mongodb":
            doc = self.mongo_db["poems"].find_one({"$or": [{"_id": record_id}, {"id": record_id}]})
            if not doc:
                return None
            doc["id"] = doc.get("id") or str(doc["_id"])
            doc.pop("_id", None)
            return doc
        with sqlite3.connect(self.sqlite_path) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT * FROM poems WHERE id = ?", (record_id,)).fetchone()
            if not row:
                return None
            return {
                "id": row["id"], "title": row["title"], "poet": row["poet"], "region": row["region"],
                "poem_text": row["poem_text"], "poem_type": row["poem_type"],
                "dominant_meter": row["dominant_meter"], "source_engine": row["source_engine"],
                "analysis": json.loads(row["analysis_json"]), "created_at": row["created_at"],
            }

    def update_poem(self, record_id: str, payload: UpdatePoemRequest) -> bool:
        clean_title, clean_poet = clean_title_and_poet(payload.title, payload.poet)
        region = payload.region.strip() or DEFAULT_REGION_LABEL
        poem_text = payload.poem_text.strip()
        existing = self.get_poem(record_id)
        if not existing:
            return False
        dominant_meter = existing.get("dominant_meter", "قيد التدقيق العروضي")
        updates = {
            "title": clean_title,
            "poet": clean_poet,
            "region": region,
            "poem_text": poem_text,
            "search_text": build_poem_search_text(clean_title, clean_poet, region, poem_text, dominant_meter),
        }
        if self.active_mode == "mongodb":
            ok = self.mongo_db["poems"].update_one({"$or": [{"_id": record_id}, {"id": record_id}]}, {"$set": updates}).matched_count > 0
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                cur = conn.execute(
                    "UPDATE poems SET title = ?, poet = ?, region = ?, poem_text = ? WHERE id = ?",
                    (updates["title"], updates["poet"], updates["region"], updates["poem_text"], record_id),
                )
                conn.commit()
                ok = cur.rowcount > 0
        if ok:
            self._invalidate_stats_cache()
        return ok

    def update_poem_analysis(self, record_id: str, analysis: Dict, source_engine: str) -> bool:
        rec = self.get_poem(record_id)
        if not rec:
            return False
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M")
        dominant_meter = analysis.get("dominant_meter", rec.get("dominant_meter", "قيد التدقيق العروضي"))
        poem_type = analysis.get("poem_type", rec.get("poem_type", "نص شعري"))
        search_text = build_poem_search_text(rec["title"], rec["poet"], rec["region"], rec["poem_text"], dominant_meter)
        glossary_items = [
            {
                "term": str(g.get("term", "")).strip(),
                "arabic_meaning": g.get("arabic_meaning", ""),
                "english_meaning": g.get("english_meaning", ""),
                "source_poem_title": rec["title"],
                "source_poem_id": record_id,
                "updated_at": now_iso,
            }
            for g in (analysis.get("glossary") or [])
            if str(g.get("term", "")).strip()
        ]

        if self.active_mode == "mongodb":
            self.mongo_db["poems"].update_one(
                {"$or": [{"_id": record_id}, {"id": record_id}]},
                {
                    "$set": {
                        "analysis": analysis,
                        "dominant_meter": dominant_meter,
                        "poem_type": poem_type,
                        "source_engine": source_engine,
                        "verification_status": "تحليل أكاديمي آلي" if source_engine.startswith("AI") else "قيد التدقيق العروضي",
                        "search_text": search_text,
                    }
                },
            )
            self._upsert_glossary_items_mongo(glossary_items)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.execute(
                    "UPDATE poems SET poem_type = ?, dominant_meter = ?, source_engine = ?, analysis_json = ? WHERE id = ?",
                    (poem_type, dominant_meter, source_engine, json.dumps(analysis, ensure_ascii=False), record_id),
                )
                for g in glossary_items:
                    conn.execute(
                        "INSERT OR REPLACE INTO glossary (term, arabic_meaning, english_meaning, source_poem_title, updated_at) VALUES (?, ?, ?, ?, ?)",
                        (g["term"], g["arabic_meaning"], g["english_meaning"], g["source_poem_title"], now_iso),
                    )
                conn.commit()
        self._invalidate_stats_cache()
        return True

    def clean_existing_archive(self, local_eng: LocalHassaniyaEngine) -> Dict[str, int]:
        """تنظيف شامل للأرشيف الحالي: فصل الشاعر عن العنوان وإصلاح الأوزان التقديرية المكسورة."""
        poems = self.list_poems()
        fixed_poets = 0
        fixed_meters = 0
        updated_total = 0

        for rec in poems:
            old_title = rec.get("title", "")
            old_poet = rec.get("poet", "")
            old_meter = str(rec.get("dominant_meter", ""))
            new_title, new_poet = clean_title_and_poet(old_title, old_poet)

            needs_poet_fix = (new_title != old_title) or (new_poet != old_poet)
            old_meter_typo = ("السغير" in old_meter) or ("البت لكبير" in old_meter)
            is_broken_meter = (
                "وزن تقديري" in old_meter
                or "(1 متحرك)" in old_meter
                or "(2 متحرك)" in old_meter
                or "(3 متحرك)" in old_meter
                or "(4 متحرك)" in old_meter
                or old_meter_typo
                or not old_meter.strip()
            )
            is_local = not str(rec.get("source_engine", "")).startswith("AI")

            new_analysis = rec.get("analysis") or {}
            new_meter = old_meter
            new_type = rec.get("poem_type", "نص شعري")

            fixed_this_meter = False
            if is_broken_meter and is_local and rec.get("poem_text"):
                reanalyzed = local_eng.analyze(rec["poem_text"])
                new_analysis = reanalyzed
                new_meter = reanalyzed["dominant_meter"]
                new_type = reanalyzed["poem_type"]
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

            if needs_poet_fix or fixed_this_meter:
                updated_total += 1
                search_text = build_poem_search_text(
                    new_title, new_poet, rec.get("region", DEFAULT_REGION_LABEL), rec.get("poem_text", ""), new_meter
                )
                if self.active_mode == "mongodb":
                    self.mongo_db["poems"].update_one(
                        {"$or": [{"_id": rec["id"]}, {"id": rec["id"]}]},
                        {
                            "$set": {
                                "title": new_title,
                                "poet": new_poet,
                                "dominant_meter": new_meter,
                                "poem_type": new_type,
                                "analysis": new_analysis,
                                "search_text": search_text,
                            }
                        },
                    )
                else:
                    with sqlite3.connect(self.sqlite_path) as conn:
                        conn.execute(
                            "UPDATE poems SET title = ?, poet = ?, poem_type = ?, dominant_meter = ?, analysis_json = ? WHERE id = ?",
                            (new_title, new_poet, new_type, new_meter, json.dumps(new_analysis, ensure_ascii=False), rec["id"]),
                        )
                        conn.commit()

        self._invalidate_stats_cache()
        return {
            "scanned": len(poems),
            "updated_total": updated_total,
            "fixed_poets_and_titles": fixed_poets,
            "fixed_broken_meters": fixed_meters,
        }

    def list_poems(self, q: str = "", poem_type: str = "الكل", poet: str = "", meter: str = "") -> List[Dict]:
        records: List[Dict] = []
        if self.active_mode == "mongodb":
            for doc in self.mongo_db["poems"].find().sort("created_at", -1):
                doc["id"] = doc.get("id") or str(doc["_id"])
                doc.pop("_id", None)
                records.append(doc)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.row_factory = sqlite3.Row
                for row in conn.execute("SELECT * FROM poems ORDER BY created_at DESC").fetchall():
                    records.append({
                        "id": row["id"], "title": row["title"], "poet": row["poet"], "region": row["region"],
                        "poem_text": row["poem_text"], "poem_type": row["poem_type"],
                        "dominant_meter": row["dominant_meter"], "source_engine": row["source_engine"],
                        "analysis": json.loads(row["analysis_json"]), "created_at": row["created_at"],
                    })

        if poem_type and poem_type != "الكل":
            pt_norm = normalize_hassaniya(poem_type, unify_gaf_to_kaf=True)
            records = [r for r in records if pt_norm in normalize_hassaniya(r.get("poem_type", ""), unify_gaf_to_kaf=True)]

        if poet.strip():
            poet_norm = normalize_hassaniya(poet, unify_gaf_to_kaf=True)
            records = [r for r in records if poet_norm in normalize_hassaniya(r.get("poet", ""), unify_gaf_to_kaf=True)]

        if meter.strip():
            meter_norm = normalize_hassaniya(meter, unify_gaf_to_kaf=True)
            records = [r for r in records if meter_norm in normalize_hassaniya(r.get("dominant_meter", ""), unify_gaf_to_kaf=True)]

        if q.strip():
            q_strict = normalize_hassaniya(q)
            q_flex = normalize_hassaniya(q, unify_gaf_to_kaf=True)
            filtered = []
            for r in records:
                hay = r.get("search_text") or build_poem_search_text(
                    r.get("title", ""), r.get("poet", ""), r.get("region", ""), r.get("poem_text", ""), r.get("dominant_meter", "")
                )
                if q_strict in hay or q_flex in hay:
                    filtered.append(r)
            records = filtered
        return records

    def get_filters(self) -> Dict[str, List[str]]:
        poems = self.list_poems()
        poets = sorted({p.get("poet", "").strip() for p in poems if p.get("poet") and not is_default_poet(p.get("poet"))})
        meters = sorted({p.get("dominant_meter", "").strip() for p in poems if p.get("dominant_meter")})
        regions = sorted({p.get("region", "").strip() for p in poems if p.get("region")})
        types_ = sorted({p.get("poem_type", "").strip() for p in poems if p.get("poem_type")})
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
            res = self.mongo_db["poems"].delete_many(
                {"$or": [{"_id": {"$in": clean_ids}}, {"id": {"$in": clean_ids}}]}
            )
            try:
                rag_doc_ids = [f"poem::{rid}" for rid in clean_ids]
                self.mongo_db[RAG_COLLECTION_NAME].delete_many({"doc_id": {"$in": rag_doc_ids}})
            except Exception:
                pass
            count = int(res.deleted_count)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                placeholders = ",".join("?" for _ in clean_ids)
                cur = conn.execute(f"DELETE FROM poems WHERE id IN ({placeholders})", clean_ids)
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
                if include_glossary:
                    deleted_glossary = int(conn.execute("DELETE FROM glossary").rowcount)
                conn.commit()
        self._invalidate_stats_cache()
        return {
            "deleted_poems": deleted_poems,
            "deleted_glossary": deleted_glossary,
            "deleted_rag": deleted_rag,
        }

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
            q_flex = normalize_hassaniya(q, unify_gaf_to_kaf=True)
            items = [
                g for g in items
                if any(
                    q_flex in normalize_hassaniya(str(g.get(k, "")), unify_gaf_to_kaf=True)
                    for k in ("term", "arabic_meaning", "english_meaning", "source_poem_title")
                )
            ]
        return items

    def delete_glossary_term(self, term: str) -> bool:
        if self.active_mode == "mongodb":
            ok = self.mongo_db["glossary"].delete_one({"term": term}).deleted_count > 0
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                cur = conn.execute("DELETE FROM glossary WHERE term = ?", (term,))
                conn.commit()
                ok = cur.rowcount > 0
        if ok:
            self._invalidate_stats_cache()
        return ok

    def save_message(self, payload: ContactMessageRequest) -> str:
        msg_id = str(uuid.uuid4())[:8]
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
        items: List[Dict] = []
        if self.active_mode == "mongodb":
            for doc in self.mongo_db["messages"].find().sort("created_at", -1):
                doc["id"] = doc.get("id") or str(doc["_id"])
                doc.pop("_id", None)
                items.append(doc)
        else:
            with sqlite3.connect(self.sqlite_path) as conn:
                conn.row_factory = sqlite3.Row
                items = [dict(r) for r in conn.execute("SELECT * FROM messages ORDER BY created_at DESC").fetchall()]
        return items

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
            poems_count = self.mongo_db["poems"].count_documents({})
            terms_count = self.mongo_db["glossary"].count_documents({})
            messages_count = self.mongo_db["messages"].count_documents({})
            total_lines = sum(
                len(doc.get("analysis", {}).get("lines", [])) or len([l for l in str(doc.get("poem_text", "")).split("\n") if l.strip()])
                for doc in self.mongo_db["poems"].find({}, {"analysis.lines": 1, "poem_text": 1})
            )
        else:
            poems = self.list_poems()
            poems_count, terms_count, messages_count = len(poems), len(self.list_glossary()), len(self.list_messages())
            total_lines = sum(
                len(p.get("analysis", {}).get("lines", [])) or len([l for l in str(p.get("poem_text", "")).split("\n") if l.strip()])
                for p in poems
            )

        res = {
            "poems_count": poems_count,
            "lines_count": total_lines,
            "terms_count": terms_count,
            "messages_count": messages_count,
            "db_mode": self.active_mode,
            "db_status": self.status_message,
            "ai_ready": bool(GEMINI_API_KEY),
            "default_model": DEFAULT_GEMINI_MODEL,
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


def analyze_with_gemini(poem_text: str, model_name: str) -> Tuple[Dict, str]:
    client = genai.Client(api_key=GEMINI_API_KEY)
    sys_inst = """أنت باحث أكاديمي وعروضي متخصص في الشعر الشعبي الحساني («لغن») وبحوره («لَبْتُوتَة»، ومفردها «بَتْ»). التزم بدقة بالمنهجية التالية:
1. أساس الوزن في «لغن»: عدّ الحروف المتحركة في الوحدة الشعرية المسماة «تافلويت» (جمعها «تيفلواتن») مع إهمال السواكن؛ فالحرف المتحرك يُعدّ واحداً، وإذا تبعه ساكن يُضمّ إليه في التقطيع الإيقاعي، مع مراعاة الابتداء بالساكن واستخدام حرف (گ) للقاف المعقودة.
2. وحدات «لغن» وأجناسه:
   - الطلعة: وحدة شائعة تتكوّن من ثلاث «تيفلواتن» برويّ واحد.
   - الكاف (الگاف): وحدة تتكوّن غالباً من أربع أو ست أو ثماني «تيفلواتن».
   - التبراع: فن نسائي يقوم غالباً على تافلويتين متفقتي القافية والوزن.
   - الظّهر: المقام أو النغم الموسيقي الذي يرافق أداء البحر ويمنحه لحنه الخاص (اذكره في حقل zahr_mode إن أمكن تحديده).
3. أبرز بحور «لَبْتُوتَة» المعتمدة في التصنيف:
   - التيدوم (من أقدم البحور المعروفة في الموزون الحساني).
   - البت الكبير (من الأبحر الجامعة، يجمع أوزاناً من مراحل سابقة، 9 متحركات فأكثر).
   - بوعمران (7 متحركات — مشهور في الفخر والحماسة).
   - المزارگ (موضوعات اجتماعية وشعبية).
   - المصارع (الفخر والمآثر).
   - إملخ واسغير (القصائد القصيرة والمواضيع العاطفية؛ اسغير 6 متحركات).
   - الواكدي (الغزل والحب).
   - أحويويص (بحر حماسي له تقطيع ومقاطع مميزة).
   - الدگسري (إيقاع سريع يناسب الاحتفالات).
   - المشكل (الألغاز الشعرية).
   - بحور أخرى: إمريميدة، لبتيت التام (8 متحركات)، لبتيت الناقص (7 متحركات)، السروزي، الرسم، لبير (6 متحركات)، بغزز، بت اثنين، بت اثلاثة، وحثو اجراد.
4. استخرج أهم المفردات الحسانية الأصيلة مع شرحها بالعربية الفصحى والإنجليزية.
5. قدّم ترجمة أدبية دقيقة للإنجليزية والفرنسية."""
    cfg = types.GenerateContentConfig(
        system_instruction=sys_inst,
        response_mime_type="application/json",
        response_schema=AIPoemAnalysis,
    )
    resp, used_model = generate_content_safe(client, model_name, f"حلّل هذا النص الشعري الحساني تحليلاً أكاديمياً وعروضياً كاملاً:\n\n{poem_text}", cfg)
    return AIPoemAnalysis.model_validate_json(resp.text).model_dump(), used_model


def split_poems_with_gemini(raw_text: str, model_name: str, default_poet: str, default_region: str) -> Tuple[List[Dict], str]:
    client = genai.Client(api_key=GEMINI_API_KEY)
    sys_inst = """أنت باحث متخصص في تحرير وفهرسة دواوين الشعر الحساني. افصل كل قصيدة في عنصر مستقل بدقة:
1. استخرج عنوان كل قصيدة نظيفاً بدون إلحاق اسم الشاعر بالعنوان، وإن لم يُذكر اقترح عنواناً موجزاً من مضمونها.
2. استخرج اسم الشاعر أو الراوي وحده في حقل poet (بدون تكرار كلمة الشاعر/الراوي)، وإن لم يُذكر اتركه فارغاً.
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
# 8. تطبيق FastAPI وترويسات الأمان والـ SEO/GEO
# =============================================================================
app = FastAPI(title="ALBIDAN AI Lab — Academic Hassaniya Corpus API", version="4.1.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


@app.middleware("http")
async def security_headers_middleware(request: Request, call_next):
    response = await call_next(request)
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    return response


# دمج محرك البحث المعرفي الهجين السحابي (Cloud Hybrid RAG)
try:
    from hassaniya_rag_router import router as hassaniya_rag_router
    app.include_router(hassaniya_rag_router)
except ImportError:
    pass

archive = ArchiveManager(MONGODB_URI, MONGODB_DB_NAME)
local_engine = LocalHassaniyaEngine()
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
def serve_home(request: Request):
    index_file = STATIC_DIR / "index.html"
    if not index_file.exists():
        raise HTTPException(status_code=404, detail="ملف الواجهة static/index.html غير موجود.")
    base_url = get_base_url(request)
    return HTMLResponse(content=index_file.read_text(encoding="utf-8").replace("__BASE_URL__", base_url))


# -----------------------------------------------------------------------------
# صفحات SEO ومحركات البحث بالذكاء الاصطناعي (SSR Poems + Sitemap + llms.txt)
# -----------------------------------------------------------------------------
@app.get("/poem/{record_id}", response_class=HTMLResponse)
def serve_poem_ssr(record_id: str, request: Request):
    rec = archive.get_poem(record_id)
    if not rec:
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    base_url = get_base_url(request)
    canonical = f"{base_url}/poem/{quote(rec['id'])}"
    analysis = rec.get("analysis") or {}
    title = rec.get("title") or "قصيدة حسانية"
    poet = rec.get("poet") or DEFAULT_POET_LABEL
    region = rec.get("region") or DEFAULT_REGION_LABEL
    meter = rec.get("dominant_meter") or "—"
    ptype = rec.get("poem_type") or "شعر حساني"
    poem_text = rec.get("poem_text") or ""
    excerpt = " / ".join([l.strip() for l in poem_text.split("\n") if l.strip()][:2])

    schema_obj = {
        "@context": "https://schema.org",
        "@type": "CreativeWork",
        "additionalType": "https://schema.org/Poem",
        "name": title,
        "author": {"@type": "Person", "name": poet},
        "inLanguage": ["mey", "ar"],
        "genre": ptype,
        "spatialCoverage": region,
        "text": poem_text,
        "url": canonical,
        "publisher": {"@type": "Organization", "name": "ALBIDAN AI Lab — مختبر البيظان"},
    }
    en_trans = analysis.get("english_translation") or ""
    fr_trans = analysis.get("french_translation") or ""
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
  <script type="application/ld+json">{json.dumps(schema_obj, ensure_ascii=False)}</script>
  <style>
    body{{margin:0;background:#0D0B09;color:#F4EFE6;font-family:'IBM Plex Sans Arabic',sans-serif;line-height:1.8;padding:1.5rem 1rem}}
    .wrap{{max-width:780px;margin:0 auto;background:#1A1612;border:1px solid rgba(216,178,84,.25);border-top:3px solid #D8B254;border-radius:16px;padding:1.8rem}}
    a{{color:#EBD07E;text-decoration:none}}
    .meta{{color:#A69C8E;font-size:.88rem;margin:.4rem 0 1.2rem}}
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
    <div class="meta">الشاعر / الراوي: <strong>{html.escape(poet)}</strong> • المنطقة: {html.escape(region)} • الوحدة: {html.escape(ptype)} • البَتْ (البحر): {html.escape(meter)}{f" • الظّهر: {html.escape(analysis.get('zahr_mode'))}" if analysis.get('zahr_mode') else ''}</div>
    <div class="poem">{html.escape(poem_text)}</div>
    {f'<div class="box"><strong>السياق الثقافي والغرض:</strong><br>{html.escape(context)}</div>' if context else ''}
    {f'<div class="box ltr"><strong>English Literary Translation:</strong><br>{html.escape(en_trans)}</div>' if en_trans else ''}
    {f'<div class="box ltr"><strong>Traduction Française:</strong><br>{html.escape(fr_trans)}</div>' if fr_trans else ''}
    <div class="actions">
      <a class="btn" href="{base_url}/">تصفح الديوان والمختبر التفاعلي</a>
      <a class="btn btn-sec" href="{base_url}/api/export/poem/{quote(rec['id'])}.docx">تنزيل الوثيقة (Word)</a>
    </div>
  </article>
</body>
</html>"""
    return HTMLResponse(content=page_html)


@app.get("/sitemap.xml")
def serve_sitemap(request: Request):
    base_url = get_base_url(request)
    poems = archive.list_poems()
    urls = [f"<url><loc>{base_url}/</loc><changefreq>daily</changefreq><priority>1.0</priority></url>"]
    for p in poems:
        urls.append(f"<url><loc>{base_url}/poem/{quote(p['id'])}</loc><changefreq>weekly</changefreq><priority>0.8</priority></url>")
    xml = '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' + "\n".join(urls) + "\n</urlset>"
    return Response(content=xml, media_type="application/xml")


@app.get("/robots.txt", response_class=PlainTextResponse)
def serve_robots(request: Request):
    base_url = get_base_url(request)
    return (
        "User-agent: *\n"
        "Allow: /\n"
        "Disallow: /api/admin/\n"
        "Disallow: /api/messages\n\n"
        "User-agent: GPTBot\nAllow: /\n\n"
        "User-agent: PerplexityBot\nAllow: /\n\n"
        "User-agent: ClaudeBot\nAllow: /\n\n"
        "User-agent: Google-Extended\nAllow: /\n\n"
        f"Sitemap: {base_url}/sitemap.xml\n"
    )


@app.get("/llms.txt", response_class=PlainTextResponse)
def serve_llms_txt(request: Request):
    base_url = get_base_url(request)
    stats = archive.get_stats()
    return f"""# ALBIDAN — Digital Corpus & Prosody Lab of Hassaniya Poetry (مختبر وديوان البيظان)

> Open academic platform dedicated to Hassaniya Arabic (ISO 639-3: `mey`) oral and written poetry from the Moroccan Sahara, Mauritania, and the Hassaniya cultural sphere. Founded and developed by Aggar Lahbib (أڭار الحبيب).

## Live Corpus Statistics
- Documented Poems: {stats['poems_count']}
- Prosodic Hemistichs (Teflouit / تافلويت): {stats['lines_count']}
- Growing Hassaniya Glossary Terms: {stats['terms_count']}

## Core Prosody Reference (عروض الشعر الحساني «لغن» — لَبْتُوتَة)
- **Lebtouta (لَبْتُوتَة, sg. Bett / بَتْ):** The poetic meters of Hassaniya poetry («لغن»), measured by counting vocalized consonants (Mutaharrikat / المتحركات) while disregarding unvocalized consonants (which attach rhythmically to the preceding vocalized letter).
- **Taflowit (تافلويت, pl. Tiflouaten / تيفلواتن):** The base poetic hemistich / prosodic unit on which meter is measured.
- **Tal'aa (طلعة):** A common poetic unit of three Tiflouaten sharing a single rhyme.
- **Gaf / Kaf (الكاف / الگاف):** A poetic stanza typically composed of four, six, or eight Tiflouaten.
- **Tebraa (تبراع):** Women's lyrical form, typically two rhyming Tiflouaten.
- **Zahr (الظّهر):** The musical mode/melody accompanying the performance of a meter and giving it its distinct tune.
- **Primary & Classical Meters (لَبْتُوتَة):** Et-Tidoum (التيدوم), El-Bett El-Kbir (البت الكبير), Bouamrane (بوعمران — 7 mutaharrikat, pride & valor), El-Mzareg (المزارگ — social), El-Msare' (المصارع — valor), Imlakh & Esghir (إملخ واسغير — short & emotional), El-Wakedi (الواكدي — love/ghazal), Ahwiwis (أحويويص — enthusiastic), Ed-Degsri (الدگسري — fast celebratory), El-Moshkel (المشكل — riddles), plus Imrimida, Lebteit Et-tam (8), Lebteit En-naqes (7), Es-Serouzi, Er-Rasm, Lbeyr (6), Baghzez, Bett Ethnayn, Bett Ethlatha, and Hathw Ajrad.
- **Orthography:** Uses `گ` consistently for the voiced velar Hassaniya /g/ (القاف المعقودة).

## Public Academic Endpoints
- Interactive Web Portal: {base_url}/
- Full Poems JSON API: {base_url}/api/poems
- Growing Hassaniya Glossary JSON API: {base_url}/api/glossary
- Open JSONL Corpus Download: {base_url}/api/export/jsonl
- XML Sitemap of Individual Poem Pages: {base_url}/sitemap.xml
"""


# -----------------------------------------------------------------------------
# مسارات الـ API التشغيلية والأكاديمية
# -----------------------------------------------------------------------------
@app.get("/health")
def health_check():
    return {"status": "ok", "db_mode": archive.active_mode, "version": "4.1.0"}


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
    """تنظيف تلقائي لكل القصائد المؤرشفة: فصل اسم الشاعر عن العنوان وإصلاح الأوزان المكسورة."""
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
    """تحليل دفعة من القصائد المحلية بـ Gemini AI لاستخراج المفردات وتغذية المعجم المتنامي."""
    require_admin(x_admin_pin, request)
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=400, detail="مفتاح GEMINI_API_KEY غير مضبوط على الخادم.")
    model = req.model_name if req.model_name in ALLOWED_MODELS else DEFAULT_GEMINI_MODEL
    all_poems = archive.list_poems()
    candidates = [
        p for p in all_poems
        if (not str(p.get("source_engine", "")).startswith("AI") or not (p.get("analysis") or {}).get("glossary"))
        and (p.get("poem_text") or "").strip()
    ][: req.limit]

    enriched, failed, added_terms = 0, 0, 0
    for rec in candidates:
        try:
            analysis, used_model = await run_in_threadpool(analyze_with_gemini, rec["poem_text"], model)
            archive.update_poem_analysis(rec["id"], analysis, f"AI ({used_model})")
            enriched += 1
            added_terms += len(analysis.get("glossary") or [])
        except Exception:
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
            analysis, used_model = await run_in_threadpool(analyze_with_gemini, text, model)
            source_engine = f"AI ({used_model})"
        except Exception as exc:
            analysis, source_engine = local_engine.analyze(text), "Local (Fallback)"
            warning = f"تعذر الاتصال بـ Gemini ({exc})، تم العرض بالمحرك المحلي."
    else:
        analysis, source_engine = local_engine.analyze(text), "Local"
    return {"source_engine": source_engine, "warning": warning, "analysis": analysis}


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
        raise HTTPException(status_code=400, detail="حجم الصوت يتجاوز 20 ميجابايت. يُفضّل تحويله لـ MP3 مضغوط أولاً.")
    ext = Path(file.filename or "audio.mp3").suffix.lower()
    mime_type = AUDIO_MIME_MAP.get(ext, file.content_type or "audio/mp3")
    if mime_type == "application/octet-stream":
        mime_type = "audio/mp3"
    try:
        text, used_model = await run_in_threadpool(transcribe_audio_sync, audio_bytes, mime_type, ext or ".mp3", model_name)
    except Exception as exc:
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
            raise HTTPException(status_code=400, detail="صيغة Word القديمة (.doc) غير مدعومة؛ احفظ الملف بصيغة .docx ثم أعد رفعه.")
        if ext == ".pdf":
            text, method = await run_in_threadpool(extract_pdf_sync, raw, model_name)
            return {"text": text, "method": method}
        if ext == ".txt":
            return {"text": raw.decode("utf-8", errors="ignore").strip(), "method": "txt"}
        raise HTTPException(status_code=400, detail="الصيغ المدعومة للاستيراد هي: Word (.docx) و PDF (.pdf) و TXT.")
    except HTTPException:
        raise
    except Exception as exc:
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
        raise HTTPException(status_code=400, detail="احفظ الملف من Word بصيغة (.docx) الحديثة بدلاً من (.doc) القديمة.")
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
            warning = f"تعذر التقسيم الذكي ({exc})، تم التحويل للتقسيم الهيكلي السريع."
    if not items:
        if ext == ".docx":
            items = await run_in_threadpool(split_docx_structurally, raw, default_poet, default_region)
        else:
            blocks = [b.strip() for b in re.split(r"\n\s*\n", plain_text) if b.strip()]
            items = []
            for i, b in enumerate(blocks, start=1):
                c_title, c_poet = clean_title_and_poet(f"قصيدة {i}", default_poet.strip() or DEFAULT_POET_LABEL)
                items.append({
                    "title": c_title,
                    "poet": c_poet,
                    "region": default_region.strip() or DEFAULT_REGION_LABEL,
                    "poem_text": b,
                })
        if use_ai_split and len(plain_text) > 35000:
            warning = "الملف طويل جداً، تم استخدام التقسيم الهيكلي السريع لضمان عدم فقدان أي قصيدة."
    del raw
    return {"count": len(items), "method": method, "warning": warning, "items": items}


@app.get("/api/poems")
def api_list_poems(
    q: str = Query(""),
    poem_type: str = Query("الكل"),
    poet: str = Query(""),
    meter: str = Query(""),
):
    return {"items": archive.list_poems(q=q, poem_type=poem_type, poet=poet, meter=meter)}


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
            title=item.title,
            poet=item.poet,
            region=item.region,
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


@app.delete("/api/poems/{record_id}")
def api_delete_poem(record_id: str, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    if not archive.delete_poem(record_id):
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    return {"deleted": True, "id": record_id}


@app.get("/api/glossary")
def api_list_glossary(q: str = Query("")):
    return {"items": archive.list_glossary(q=q)}


@app.delete("/api/glossary/{term}")
def api_delete_glossary(term: str, request: Request, x_admin_pin: Optional[str] = Header(default=None)):
    require_admin(x_admin_pin, request)
    if not archive.delete_glossary_term(term):
        raise HTTPException(status_code=404, detail="المفردة غير موجودة.")
    return {"deleted": True, "term": term}


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
def api_export_glossary_docx(q: str = Query("")):
    items = archive.list_glossary(q=q)
    doc = create_base_arabic_doc("المعجم الحساني المتنامي", f"تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d')}  |  عدد المفردات الموثقة: {len(items)}")
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    set_table_rtl(tbl)
    for i, h in enumerate(["المفردة الحسانية", "المعنى بالعربية الفصحى", "English Meaning", "المصدر (القصيدة)"]):
        cell = tbl.rows[0].cells[i]
        set_cell_shading(cell, "1B365D")
        p = cell.paragraphs[0]
        set_paragraph_rtl(p)
        set_run_arabic(p.add_run(h), size_pt=10.5, bold=True, color_hex="FFFFFF")
    for g in items:
        cells = tbl.add_row().cells
        for i, v in enumerate([g.get("term", ""), g.get("arabic_meaning", ""), g.get("english_meaning", ""), g.get("source_poem_title", "")]):
            p = cells[i].paragraphs[0]
            set_paragraph_rtl(p)
            set_run_arabic(p.add_run(v), size_pt=10.5, bold=(i == 0))
    return doc_to_stream(doc, "albidan_hassaniya_glossary.docx")


@app.post("/api/export/analysis.docx")
def api_export_analysis_docx(req: ExportAnalysisDocxRequest):
    doc = create_base_arabic_doc("تقرير التحليل العروضي واللساني", f"تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d')}")
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
    doc = create_base_arabic_doc("المدونة الحسانية المفتوحة — ديوان النصوص الموثقة", f"تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d')}  |  عدد القصائد: {len(poems)}")
    for idx, rec in enumerate(poems):
        append_poem_section_to_doc(doc, rec["title"], rec["poet"], rec["region"], rec["poem_text"], rec.get("analysis", {}))
        if idx < len(poems) - 1:
            doc.add_page_break()
    return doc_to_stream(doc, "albidan_hassaniya_anthology.docx")


@app.get("/api/export/jsonl")
def api_export_jsonl():
    content = "\n".join(json.dumps(p, ensure_ascii=False) for p in archive.list_poems())
    return StreamingResponse(
        io.BytesIO(content.encode("utf-8")),
        media_type="application/jsonl",
        headers={"Content-Disposition": "attachment; filename=albidan_hassaniya_corpus.jsonl"},
    )
