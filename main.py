import os
import re
import io
import time
import json
import uuid
import sqlite3
import tempfile
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple

from dotenv import load_dotenv
from fastapi import FastAPI, Request, UploadFile, File, Form, HTTPException, Header, Query
from fastapi.concurrency import run_in_threadpool
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
DEFAULT_GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip() or "gemini-3.5-flash-lite"
ADMIN_PIN = os.getenv("ADMIN_PIN", "").strip()

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
    linguistic_notes: str = Field(description="ملاحظات لغوية وعروضية منهجية حول النص")
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
    poet: str = "تراث شفهي / غير معروف"
    region: str = "الصحراء المغربية"
    poem_text: str
    analysis: Dict
    source_engine: str = "AI"


class BatchSaveItem(BaseModel):
    title: str = ""
    poet: str = "تراث شفهي / غير معروف"
    region: str = "الصحراء المغربية"
    poem_text: str
    analysis: Optional[Dict] = None
    source_engine: str = "Local (Batch Word)"


class BatchSaveRequest(BaseModel):
    items: List[BatchSaveItem]


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
    category: str = "تواصل بحثي
