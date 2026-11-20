import os
import re
import io
import csv
import json
import uuid
import sqlite3
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional

from dotenv import load_dotenv
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Header, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

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
# 1. هياكل البيانات (Pydantic Schemas)
# ==========================================
class AILineAnalysis(BaseModel):
    original_line: str = Field(description="الشطر الأصلي كما أدخله المستخدم")
    phonetic_diacritized: str = Field(description="الشطر مضبوطاً بالشكل الصوتي الحساني")
    mutaharrikat_count: int = Field(description="العدد الدقيق للمتحركات في الشطر")
    estimated_meter: str = Field(description="اسم البحر الحساني")
    rhyme_ending: str = Field(description="الروي أو القافية الصوتية للشطر")


class GlossaryTerm(BaseModel):
    term: str = Field(description="المفردة الحسانية")
    arabic_meaning: str = Field(description="معناها بالعربية الفصحى")
    english_meaning: str = Field(description="معناها المختصر بالإنجليزية")


class AIPoemAnalysis(BaseModel):
    poem_type: str = Field(description="نوع القطعة: تبراع، كاف، طلعة، أو شعر حر")
    dominant_meter: str = Field(description="البحر السائد في القصيدة")
    cultural_context: str = Field(description="شرح موجز للغرض الشعري والرمزية الثقافية الحسانية")
    recommended_azawan: str = Field(description="المقام الموسيقي الأنسب من أزوان مع السبب")
    lines: List[AILineAnalysis]
    glossary: List[GlossaryTerm] = Field(description="شرح لأهم المفردات الحسانية التراثية")
    english_translation: str = Field(description="ترجمة شعرية وأدبية رفيعة للإنجليزية")
    french_translation: str = Field(description="ترجمة شعرية وأدبية رفيعة للفرنسية")


class AnalyzeRequest(BaseModel):
    poem_text: str
    model_name: str = DEFAULT_GEMINI_MODEL
    use_ai: bool = True
    azawan_mode: str = "لبياظ (Labyad) — الحنين والغزل"
    fusion_style: str = "World Music Fusion"
    vocal_type: str = "Deep soulful male Hassaniya vocals"


class AzawanRequest(BaseModel):
    poem_text: str
    mode: str = "لبياظ (Labyad) — الحنين والغزل"
    fusion_style: str = "World Music Fusion"
    vocal_type: str = "Deep soulful male Hassaniya vocals"


class SavePoemRequest(BaseModel):
    title: str
    poet: str = "تراث شفهي / غير معروف"
    region: str = "الصحراء المغربية"
    poem_text: str
    analysis: Dict
    music: Dict
    azawan_mode: str = "لبياظ (Labyad) — الحنين والغزل"
    source_engine: str = "AI"


# ==========================================
# 2. مدير قاعدة البيانات (MongoDB Atlas + SQLite احتياطي)
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
            "azawan_mode": payload.azawan_mode,
            "source_engine": payload.source_engine,
            "analysis": payload.analysis,
            "music": payload.music,
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
                        record["azawan_mode"],
                        record["source_engine"],
                        json.dumps(record["analysis"], ensure_ascii=False),
                        json.dumps(record["music"], ensure_ascii=False),
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

    def list_poems(self, q: str = "", poem_type: str = "الكل") -> List[Dict]:
        records: List[Dict] = []
        if self.active_mode == "mongodb":
            for doc in self.mongo_db["poems"].find().sort("created_at", -1):
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
                        "azawan_mode": row["azawan_mode"],
                        "source_engine": row["source_engine"],
                        "analysis": json.loads(row["analysis_json"]),
                        "music": json.loads(row["music_json"]),
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
            res = self.mongo_db["poems"].delete_one({"_id": record_id})
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

    def get_stats(self) -> Dict:
        if self.active_mode == "mongodb":
            poems_count = self.mongo_db["poems"].count_documents({})
            terms_count = self.mongo_db["glossary"].count_documents({})
            cursor = self.mongo_db["poems"].find({}, {"analysis.lines": 1})
            total_lines = sum(len(doc.get("analysis", {}).get("lines", [])) for doc in cursor)
        else:
            poems = self.list_poems()
            glossary = self.list_glossary()
            poems_count = len(poems)
            terms_count = len(glossary)
            total_lines = sum(len(p.get("analysis", {}).get("lines", [])) for p in poems)

        return {
            "poems_count": poems_count,
            "lines_count": total_lines,
            "terms_count": terms_count,
            "db_mode": self.active_mode,
            "db_status": self.status_message,
            "ai_ready": bool(GEMINI_API_KEY),
            "default_model": DEFAULT_GEMINI_MODEL,
        }


# ==========================================
# 3. المحلل الاحتياطي ومحرك أزوان وGemini
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
            "cultural_context": "تحليل أولي بالمحرك المحلي. أضف GEMINI_API_KEY في متغيرات الخادم لتفعيل التحليل الأكاديمي والترجمة الأدبية.",
            "recommended_azawan": "لبياظ (تقديري)",
            "lines": details,
            "glossary": [],
            "english_translation": "Configure GEMINI_API_KEY on the server for poetic English translation.",
            "french_translation": "Configurez GEMINI_API_KEY sur le serveur pour la traduction poétique en français.",
        }


class AzawanMusicDirector:
    azawan_modes = {
        "كر (Karr) — الفخر والوقار": {
            "mood": "Epic, noble, deep desert resonance, proud",
            "instruments": "Tidinit (Saharan lute), deep Tbal drums, warm ambient synth",
            "bpm": 85,
        },
        "فاغو (Faghu) — الحماس والإيقاع": {
            "mood": "Energetic, triumphant, rhythmic, uplifting",
            "instruments": "Fast Tidinit riffs, Ardin harp, handclaps, modern Afro-desert percussion",
            "bpm": 115,
        },
        "لبياظ (Labyad) — الحنين والغزل": {
            "mood": "Melancholic, romantic, nostalgic, soulful",
            "instruments": "Soft Ardin harp, acoustic Tidinit, warm cello pad, emotional vocals",
            "bpm": 75,
        },
        "لكحال (Lakhhal) — العمق والروحانية": {
            "mood": "Mystical, hypnotic, deep night desert atmosphere",
            "instruments": "Reverb-drenched Tidinit, subtle oud, slow hypnotic percussion",
            "bpm": 70,
        },
    }

    def generate_song_structure(
        self, poem_lines: List[str], mode: str, fusion_style: str, vocal_type: str
    ) -> Dict[str, str]:
        selected = self.azawan_modes.get(mode, list(self.azawan_modes.values())[2])
        style_prompt = (
            f"Hassaniya Moroccan Saharan {fusion_style}, {selected['mood']}, "
            f"{vocal_type}, featuring {selected['instruments']}, {selected['bpm']} BPM, "
            f"authentic Saharan vocal ornamentation, studio quality mastering, crystal clear mix."
        )
        lyrics = "[Intro: Solo Tidinit melody with gentle desert ambience]\n\n"
        if len(poem_lines) <= 2:
            lyrics += "[Verse 1: Intimate, expressive vocal]\n" + "\n".join(poem_lines) + "\n\n"
            lyrics += "[Instrumental Interlude: Emotional Ardin harp solo]\n\n"
            lyrics += "[Outro: Soft vocal echo fading with Tidinit]"
        elif len(poem_lines) <= 4:
            lyrics += "[Verse 1]\n" + "\n".join(poem_lines[:2]) + "\n\n"
            lyrics += "[Chorus: Full rhythm with Tbal and handclaps]\n" + "\n".join(poem_lines[2:]) + "\n\n"
            lyrics += "[Outro: Fading Tidinit and rhythmic drums]"
        else:
            lyrics += "[Verse 1]\n" + "\n".join(poem_lines[:4]) + "\n\n"
            lyrics += "[Chorus: Energetic and soulful]\n" + "\n".join(poem_lines[4:6]) + "\n\n"
            if len(poem_lines) > 6:
                lyrics += "[Bridge / Verse 2]\n" + "\n".join(poem_lines[6:]) + "\n\n"
            lyrics += "[Outro: Traditional Saharan instrumental fade-out]"

        return {"style_prompt": style_prompt, "formatted_lyrics": lyrics, "bpm": str(selected["bpm"])}


def analyze_with_gemini(poem_text: str, model_name: str) -> Dict:
    client = genai.Client(api_key=GEMINI_API_KEY)
    system_instruction = """
    أنت خبير أكاديمي وعروضي متخصص في الأدب والشعر الشعبي الحساني المغربي والموريتاني (لغنى والتبراع) وموسيقى أزوان.
    مهمتك تحليل النص الشعري الحساني بدقة لغوية وعروضية فائقة:
    1. راعِ ظاهرة "الابتداء بالساكن" في اللهجة الحسانية عند التشكيل الصوتي وعدّ المتحركات.
    2. عدّ المتحركات هو أساس الوزن الحساني (السغير = 6، لبتيت الناقص = 7، لبتيت التام = 8، البت لكبير = 9+).
    3. ميّز بدقة بين التبراع، الكاف، والطلعة.
    4. استخرج المفردات الحسانية الأصيلة واشرحها بوضوح بالعربية الفصحى والإنجليزية.
    5. قدم ترجمة شعرية عالية الجودة للإنجليزية والفرنسية تحافظ على الروح الصحراوية.
    """
    response = client.models.generate_content(
        model=model_name,
        contents=f"حلّل هذا النص الشعري الحساني تحليلاً كاملاً:\n\n{poem_text}",
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
    أنت خبير في اللهجة الحسانية والشعر الحساني (لغنى، الكاف، الطلعة، التبراع).
    فرّغ التسجيل الصوتي نصياً وفق القواعد التالية:
    1. اكتب الكلمات كما تُنطق باللهجة الحسانية الأصيلة دون تحويلها للفصحى.
    2. استخدم الحرف (گ) للقاف المعقودة الحسانية (مثل: گلب، گال، اگليب).
    3. افصل كل شطر شعري (تافلويت) في سطر مستقل تماماً.
    4. أخرج الأبيات الشعرية فقط دون عناوين أو تعليقات إضافية.
    """
    response = client.models.generate_content(
        model=model_name,
        contents=[
            types.Part.from_bytes(data=audio_bytes, mime_type=mime_type),
            prompt,
        ],
    )
    return (response.text or "").strip()


# ==========================================
# 4. تطبيق FastAPI ونقاط الاتصال (API Routes)
# ==========================================
app = FastAPI(title="ALBIDAN AI Lab API", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

archive = ArchiveManager(MONGODB_URI, MONGODB_DB_NAME)
local_engine = LocalHassaniyaEngine()
director = AzawanMusicDirector()

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/")
def serve_home():
    index_file = STATIC_DIR / "index.html"
    if not index_file.exists():
        raise HTTPException(status_code=404, detail="ملف الواجهة static/index.html غير موجود.")
    return FileResponse(str(index_file))


@app.get("/api/stats")
def api_stats():
    return archive.get_stats()


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

    lines_list = [item["original_line"] for item in analysis.get("lines", [])]
    music = director.generate_song_structure(
        lines_list, req.azawan_mode, req.fusion_style, req.vocal_type
    )
    return {
        "source_engine": source_engine,
        "warning": warning,
        "analysis": analysis,
        "music": music,
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
    mime_map = {
        ".mp3": "audio/mp3",
        ".wav": "audio/wav",
        ".m4a": "audio/mp4",
        ".ogg": "audio/ogg",
        ".flac": "audio/flac",
    }
    mime_type = mime_map.get(ext, file.content_type or "audio/mp3")
    model = model_name if model_name in ALLOWED_MODELS else DEFAULT_GEMINI_MODEL

    try:
        text = transcribe_with_gemini(audio_bytes, mime_type, model)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"خطأ أثناء التفريغ الصوتي: {exc}")

    return {"text": text, "model": model}


@app.post("/api/azawan")
def api_azawan(req: AzawanRequest):
    lines = [l.strip() for l in req.poem_text.split("\n") if l.strip()]
    if not lines:
        raise HTTPException(status_code=400, detail="أدخل شطراً واحداً على الأقل.")
    return director.generate_song_structure(lines, req.mode, req.fusion_style, req.vocal_type)


@app.get("/api/poems")
def api_list_poems(q: str = Query(""), poem_type: str = Query("الكل")):
    return {"items": archive.list_poems(q=q, poem_type=poem_type)}


@app.post("/api/poems")
def api_save_poem(req: SavePoemRequest):
    if not req.poem_text.strip():
        raise HTTPException(status_code=400, detail="نص القصيدة فارغ.")
    new_id = archive.save_poem(req)
    return {"id": new_id, "message": "تم الحفظ في المدونة وتحديث المعجم بنجاح."}


@app.delete("/api/poems/{record_id}")
def api_delete_poem(record_id: str, x_admin_pin: Optional[str] = Header(default=None)):
    if ADMIN_PIN and x_admin_pin != ADMIN_PIN:
        raise HTTPException(status_code=403, detail="يتطلب رمز المشرف الصحيح لحذف القصيدة.")
    deleted = archive.delete_poem(record_id)
    if not deleted:
        raise HTTPException(status_code=404, detail="القصيدة غير موجودة.")
    return {"deleted": True, "id": record_id}


@app.get("/api/glossary")
def api_list_glossary(q: str = Query("")):
    return {"items": archive.list_glossary(q=q)}


@app.get("/api/export/jsonl")
def api_export_jsonl():
    poems = archive.list_poems()
    content = "\n".join(json.dumps(p, ensure_ascii=False) for p in poems)
    return StreamingResponse(
        io.BytesIO(content.encode("utf-8")),
        media_type="application/jsonl",
        headers={"Content-Disposition": "attachment; filename=albidan_hassaniya_corpus.jsonl"},
    )


@app.get("/api/export/glossary.csv")
def api_export_glossary_csv():
    items = archive.list_glossary()
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["المفردة الحسانية", "المعنى بالعربية الفصحى", "English Meaning", "المصدر", "تاريخ التوثيق"])
    for g in items:
        writer.writerow([
            g.get("term", ""),
            g.get("arabic_meaning", ""),
            g.get("english_meaning", ""),
            g.get("source_poem_title", ""),
            g.get("updated_at", ""),
        ])
    return StreamingResponse(
        io.BytesIO(buf.getvalue().encode("utf-8-sig")),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=albidan_hassaniya_glossary.csv"},
    )
