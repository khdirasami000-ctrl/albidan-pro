import hashlib
import os
import re
import time
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, HTTPException
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from pymongo import MongoClient, UpdateOne
from pymongo.collection import Collection
from pymongo.errors import OperationFailure

# ==============================================================================
# 1. الإعدادات والمتغيرات البيئية (مضبوطة تلقائياً على قاعدة albidan_corpus)
# ==============================================================================
MONGODB_URI = (
    os.getenv("MONGODB_URI")
    or os.getenv("MONGO_URI")
    or os.getenv("MONGODB_URL")
    or ""
)
MONGODB_DB_NAME = (
    os.getenv("MONGODB_DB_NAME")
    or os.getenv("MONGO_DB_NAME")
    or "albidan_corpus"
)
RAG_COLLECTION_NAME = os.getenv("RAG_COLLECTION_NAME", "hassaniya_rag_chunks")

VECTOR_INDEX_NAME = os.getenv("VECTOR_INDEX_NAME", "hassaniya_vector_index")
TEXT_INDEX_NAME = os.getenv("TEXT_INDEX_NAME", "hassaniya_text_index")

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or ""
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
EMBEDDING_MODEL = os.getenv("GEMINI_EMBEDDING_MODEL", "gemini-embedding-001")
VECTOR_DIMS = int(os.getenv("VECTOR_DIMS", "768"))

router = APIRouter(prefix="/api/rag", tags=["Hassaniya Cloud Hybrid RAG"])

_mongo_client: Optional[MongoClient] = None
_genai_client: Optional[genai.Client] = None


# ==============================================================================
# 2. إدارة الاتصالات المستقرة بين Render وMongoDB Atlas
# ==============================================================================
def get_rag_collection() -> Collection:
    global _mongo_client
    if not MONGODB_URI:
        raise HTTPException(
            status_code=500,
            detail="متغير البيئة MONGODB_URI أو MONGO_URI غير مضبوط في إعدادات Render.",
        )

    if _mongo_client is None:
        _mongo_client = MongoClient(
            MONGODB_URI,
            serverSelectionTimeoutMS=5000,
            connectTimeoutMS=10000,
            socketTimeoutMS=25000,
            maxPoolSize=20,
            retryWrites=True,
        )
    return _mongo_client[MONGODB_DB_NAME][RAG_COLLECTION_NAME]


def get_source_collection(collection_name: str) -> Collection:
    get_rag_collection()
    assert _mongo_client is not None
    return _mongo_client[MONGODB_DB_NAME][collection_name]


def get_genai_client() -> genai.Client:
    global _genai_client
    if not GEMINI_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="متغير البيئة GEMINI_API_KEY غير مضبوط في إعدادات Render.",
        )
    if _genai_client is None:
        _genai_client = genai.Client(api_key=GEMINI_API_KEY)
    return _genai_client


# ==============================================================================
# 3. التطبيع اللغوي للشعر الحساني والتقطيع الذكي
# ==============================================================================
def normalize_hassaniya(text: str) -> str:
    """توحيد الحرف الحساني (گ / ڨ / ڭ -> ݣ) وإزالة التشكيل للبحث المعجمي الدقيق."""
    if not text:
        return ""
    text = re.sub(r"[\u064B-\u065F\u0670\u0640]", "", text)
    text = re.sub(r"[گڨڭ]", "ݣ", text)
    text = re.sub(r"[إأآٱ]", "ا", text)
    text = text.replace("ى", "ي").replace("ة", "ه")
    text = re.sub(r"[^\w\sݣ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def compute_content_hash(payload_text: str) -> str:
    return hashlib.sha256(payload_text.encode("utf-8")).hexdigest()


def chunk_hassaniya_text(
    text: str,
    chunk_size: int = 750,
    chunk_overlap: int = 120,
) -> List[str]:
    """تقطيع النص الشعري مع احترام حدود الأبيات والأسطر."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return []

    chunks: List[str] = []
    current_lines: List[str] = []
    current_len = 0

    for line in lines:
        line_len = len(line) + 1
        if current_len + line_len <= chunk_size or not current_lines:
            current_lines.append(line)
            current_len += line_len
        else:
            chunks.append("\n".join(current_lines))
            overlap_lines: List[str] = []
            overlap_len = 0
            for prev_line in reversed(current_lines):
                if overlap_len + len(prev_line) <= chunk_overlap:
                    overlap_lines.insert(0, prev_line)
                    overlap_len += len(prev_line) + 1
                else:
                    break
            current_lines = overlap_lines + [line]
            current_len = sum(len(x) + 1 for x in current_lines)

    if current_lines:
        chunks.append("\n".join(current_lines))

    return chunks


def _pick_first_text(raw: Dict[str, Any], candidate_keys: List[str], default: str = "") -> str:
    """استخراج أول حقل غير فارغ من قائمة أسماء محتملة لتفادي اختلاف مسميات الحقول."""
    for key in candidate_keys:
        val = raw.get(key)
        if isinstance(val, str) and val.strip():
            return val.strip()
        if isinstance(val, list) and val:
            joined = "\n".join(str(x).strip() for x in val if str(x).strip())
            if joined:
                return joined
    return default


def _fallback_serialize(raw: Dict[str, Any], exclude_keys: set) -> str:
    """تجميع كل الحقول النصية المتاحة في الوثيقة إذا لم تطابق الأسماء القياسية."""
    parts: List[str] = []
    for k, v in raw.items():
        if k in exclude_keys or v is None:
            continue
        if isinstance(v, (str, int, float)) and str(v).strip():
            parts.append(f"{k}: {str(v).strip()}")
        elif isinstance(v, list) and v:
            parts.append(f"{k}: " + "، ".join(str(x).strip() for x in v if str(x).strip()))
    return "\n".join(parts).strip()


# ==============================================================================
# 4. توليد المتجهات الرقمية (Gemini Embeddings)
# ==============================================================================
def embed_texts(texts: List[str], task_type: str = "RETRIEVAL_DOCUMENT") -> List[List[float]]:
    client = get_genai_client()
    all_vectors: List[List[float]] = []
    batch_size = 40

    for i in range(0, len(texts), batch_size):
        batch = texts[i : i + batch_size]
        response = client.models.embed_content(
            model=EMBEDDING_MODEL,
            contents=batch,
            config=types.EmbedContentConfig(
                task_type=task_type,
                output_dimensionality=VECTOR_DIMS,
            ),
        )
        if not response.embeddings:
            raise HTTPException(status_code=502, detail="فشل توليد المتجهات من Gemini API.")
        all_vectors.extend([list(emb.values) for emb in response.embeddings])

    return all_vectors


# ==============================================================================
# 5. نماذج البيانات (Pydantic Schemas)
# ==============================================================================
class PoetryDocumentInput(BaseModel):
    doc_id: str = Field(..., description="المعرّف الفريد للقصيدة أو المادة المعجمية")
    title: str = Field(..., description="العنوان أو الكلمة الحسانية")
    content: str = Field(..., description="النص الكامل للقصيدة أو الشرح")
    poet: str = Field(default="غير محدد", description="اسم الشاعر أو الراوي")
    meter: str = Field(default="غير محدد", description="البحر الحساني / البتّ")
    genre: str = Field(default="شعر حساني", description="النوع: طلعة، ݣاف، تبراع، معجم حساني")
    source: str = Field(default="أرشيف ALBIDAN", description="المصدر الموثق")
    lexical_notes: str = Field(default="", description="شرح المفردات إن وُجد")


class BulkIndexRequest(BaseModel):
    documents: List[PoetryDocumentInput]


class HybridQueryRequest(BaseModel):
    query: str = Field(..., min_length=2, description="سؤال الباحث حول الشعر الحساني أو المعجم")
    poet: Optional[str] = Field(default=None, description="فلترة اختيارية باسم الشاعر")
    meter: Optional[str] = Field(default=None, description="فلترة اختيارية بالبحر / البتّ")
    genre: Optional[str] = Field(default=None, description="فلترة اختيارية بالنوع")
    top_k: int = Field(default=5, ge=1, le=12)
    enable_reflection: bool = Field(
        default=True,
        description="تفعيل التحقق الذاتي الأكاديمي لمنع تحريف الأبيات أو خلط البحور",
    )


# ==============================================================================
# 6. الفهرسة الذكية (تتخطى النصوص غير المعدّلة بالبصمة الرقمية)
# ==============================================================================
def upsert_poetry_documents(docs: List[PoetryDocumentInput]) -> Dict[str, int]:
    col = get_rag_collection()
    stats = {"indexed_docs": 0, "skipped_docs": 0, "total_chunks": 0}

    for doc in docs:
        header = (
            f"العنوان: {doc.title} | الشاعر: {doc.poet} | "
            f"البحر (البتّ): {doc.meter} | النوع: {doc.genre} | المصدر: {doc.source}"
        )
        full_body = doc.content.strip()
        if doc.lexical_notes.strip():
            full_body += f"\n\nشرح المفردات:\n{doc.lexical_notes.strip()}"

        content_hash = compute_content_hash(f"{header}\n{full_body}")
        existing = col.find_one({"doc_id": doc.doc_id}, {"content_hash": 1})
        if existing and existing.get("content_hash") == content_hash:
            stats["skipped_docs"] += 1
            continue

        col.delete_many({"doc_id": doc.doc_id})

        chunks = chunk_hassaniya_text(full_body)
        if not chunks:
            continue

        texts_for_embedding = [f"{header}\n{chunk}" for chunk in chunks]
        vectors = embed_texts(texts_for_embedding, task_type="RETRIEVAL_DOCUMENT")

        now_iso = datetime.now(timezone.utc).isoformat()
        operations: List[UpdateOne] = []

        for idx, (chunk_text, vector) in enumerate(zip(chunks, vectors)):
            chunk_id = f"{doc.doc_id}::chunk_{idx}"
            record = {
                "chunk_id": chunk_id,
                "doc_id": doc.doc_id,
                "content_hash": content_hash,
                "title": doc.title,
                "poet": doc.poet,
                "meter": doc.meter,
                "genre": doc.genre,
                "source": doc.source,
                "text": chunk_text,
                "normalized_text": normalize_hassaniya(chunk_text),
                "normalized_meta": normalize_hassaniya(header),
                "embedding": vector,
                "updated_at": now_iso,
            }
            operations.append(UpdateOne({"chunk_id": chunk_id}, {"$set": record}, upsert=True))

        if operations:
            col.bulk_write(operations, ordered=False)
            stats["indexed_docs"] += 1
            stats["total_chunks"] += len(operations)

    return stats


# ==============================================================================
# 7. محرك البحث الهجين (Vector Search + Atlas Search + RRF)
# ==============================================================================
def _build_pre_filter(req: HybridQueryRequest) -> Dict[str, Any]:
    filt: Dict[str, Any] = {}
    if req.poet:
        filt["poet"] = req.poet
    if req.meter:
        filt["meter"] = req.meter
    if req.genre:
        filt["genre"] = req.genre
    return filt


def _vector_search(col: Collection, query_vector: List[float], req: HybridQueryRequest, limit: int) -> List[Dict[str, Any]]:
    pre_filter = _build_pre_filter(req)
    vs_stage: Dict[str, Any] = {
        "index": VECTOR_INDEX_NAME,
        "path": "embedding",
        "queryVector": query_vector,
        "numCandidates": max(60, limit * 10),
        "limit": limit,
    }
    if pre_filter:
        vs_stage["filter"] = pre_filter

    pipeline = [
        {"$vectorSearch": vs_stage},
        {
            "$project": {
                "_id": 0,
                "chunk_id": 1,
                "doc_id": 1,
                "title": 1,
                "poet": 1,
                "meter": 1,
                "genre": 1,
                "source": 1,
                "text": 1,
                "score": {"$meta": "vectorSearchScore"},
            }
        },
    ]
    try:
        return list(col.aggregate(pipeline))
    except OperationFailure:
        return []


def _atlas_lexical_search(col: Collection, query: str, req: HybridQueryRequest, limit: int) -> List[Dict[str, Any]]:
    norm_query = normalize_hassaniya(query)
    pre_filter = _build_pre_filter(req)

    pipeline: List[Dict[str, Any]] = [
        {
            "$search": {
                "index": TEXT_INDEX_NAME,
                "compound": {
                    "should": [
                        {
                            "text": {
                                "query": norm_query,
                                "path": ["normalized_text", "normalized_meta"],
                            }
                        },
                        {
                            "text": {
                                "query": query,
                                "path": ["text", "poet", "meter", "title"],
                            }
                        },
                    ]
                },
            }
        }
    ]
    if pre_filter:
        pipeline.append({"$match": pre_filter})

    pipeline.extend([
        {"$limit": limit},
        {
            "$project": {
                "_id": 0,
                "chunk_id": 1,
                "doc_id": 1,
                "title": 1,
                "poet": 1,
                "meter": 1,
                "genre": 1,
                "source": 1,
                "text": 1,
                "score": {"$meta": "searchScore"},
            }
        },
    ])

    try:
        return list(col.aggregate(pipeline))
    except OperationFailure:
        tokens = [t for t in norm_query.split() if len(t) >= 2][:6]
        if not tokens:
            return []
        regex_pattern = "|".join(re.escape(t) for t in tokens)
        mongo_filter: Dict[str, Any] = {
            "$or": [
                {"normalized_text": {"$regex": regex_pattern}},
                {"normalized_meta": {"$regex": regex_pattern}},
            ]
        }
        mongo_filter.update(pre_filter)
        cursor = col.find(
            mongo_filter,
            {
                "_id": 0,
                "chunk_id": 1,
                "doc_id": 1,
                "title": 1,
                "poet": 1,
                "meter": 1,
                "genre": 1,
                "source": 1,
                "text": 1,
            },
        ).limit(limit)
        return [{**doc, "score": 0.5} for doc in cursor]


def _fuse_rrf(
    dense_hits: List[Dict[str, Any]],
    lexical_hits: List[Dict[str, Any]],
    rrf_k: int = 60,
    dense_weight: float = 0.62,
    lexical_weight: float = 0.38,
) -> List[Dict[str, Any]]:
    fused_scores: Dict[str, float] = {}
    doc_map: Dict[str, Dict[str, Any]] = {}

    for rank, item in enumerate(dense_hits, start=1):
        cid = item["chunk_id"]
        doc_map[cid] = item
        fused_scores[cid] = fused_scores.get(cid, 0.0) + dense_weight * (1.0 / (rrf_k + rank))

    for rank, item in enumerate(lexical_hits, start=1):
        cid = item["chunk_id"]
        if cid not in doc_map:
            doc_map[cid] = item
        fused_scores[cid] = fused_scores.get(cid, 0.0) + lexical_weight * (1.0 / (rrf_k + rank))

    sorted_ids = sorted(fused_scores.keys(), key=lambda x: fused_scores[x], reverse=True)
    max_score = fused_scores[sorted_ids[0]] if sorted_ids else 1.0

    results: List[Dict[str, Any]] = []
    for cid in sorted_ids:
        entry = dict(doc_map[cid])
        entry["rrf_score"] = round(fused_scores[cid], 6)
        entry["confidence"] = round(min(0.99, fused_scores[cid] / max_score), 3)
        results.append(entry)
    return results


# ==============================================================================
# 8. التوليد الأكاديمي والتحقق الذاتي عبر Gemini API
# ==============================================================================
def generate_hassaniya_verified_answer(
    query: str,
    sources: List[Dict[str, Any]],
    enable_reflection: bool = True,
) -> Dict[str, Any]:
    if not sources:
        return {
            "answer": "لم يُعثر في أرشيف منصة ALBIDAN الموثّق على نصوص أو شواهد كافية للإجابة على هذا السؤال.",
            "verification_status": "بدون مصادر",
            "verification_notes": "قاعدة المعرفة لا تحتوي على مقاطع مطابقة حالياً.",
        }

    client = get_genai_client()

    context_blocks = []
    for i, s in enumerate(sources, start=1):
        block = (
            f"[الشاهد {i}]\n"
            f"- العنوان: {s.get('title', 'غير محدد')}\n"
            f"- الشاعر: {s.get('poet', 'غير محدد')}\n"
            f"- البحر (البتّ): {s.get('meter', 'غير محدد')}\n"
            f"- النوع: {s.get('genre', 'غير محدد')}\n"
            f"- المصدر: {s.get('source', 'أرشيف ALBIDAN')}\n"
            f"- النص:\n{s.get('text', '')}"
        )
        context_blocks.append(block)
    context_text = "\n\n---\n\n".join(context_blocks)

    system_prompt = (
        "أنت باحث أكاديمي متخصص في الأدب والشعر الشعبي الحساني (الطلعة، الݣيفان، التبراع، "
        "البتوت الحسانية، والمعجم الحساني) ضمن منصة ALBIDAN.\n"
        "قواعد علمية صارمة:\n"
        "1. اعتمد حصراً على الشواهد المرفقة في السياق ولا تؤلّف أبياتاً من عندك.\n"
        "2. انقل الشواهد الحسانية بلفظها وحروفها كما وردت في المصدر.\n"
        "3. لا تنسب أي بيت لشاعر أو بحر لم يرد صراحةً في الشواهد.\n"
        "4. وثّق إجابتك بذكر رقم الشاهد والمصدر بين قوسين."
    )

    draft_resp = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=f"الشواهد الموثّقة من أرشيف ALBIDAN:\n{context_text}\n\nسؤال الباحث:\n{query}",
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.15,
        ),
    )
    draft_answer = (draft_resp.text or "").strip()

    if not enable_reflection:
        return {
            "answer": draft_answer,
            "verification_status": "موثّقة (الوضع السريع)",
            "verification_notes": "تم التوليد من الشواهد الموثقة دون جولة مراجعة ثانية.",
        }

    critic_prompt = (
        "أنت مدقق أكاديمي صارم في التراث الحساني. راجع المسودة التالية مقابل الشواهد الموثّقة فقط:\n"
        "- تأكد من تطابق الأبيات الحسانية حرفياً مع الشواهد دون تحريف.\n"
        "- تأكد من صحة نسبة كل بيت لشاعره وبحره ومصدره.\n"
        "- احذف أي استنتاج غير مسنود بنص صريح في الشواهد.\n"
        "التزم حرفياً بهذا القالب:\n"
        "[[STATUS]]: PASS أو REVISED\n"
        "[[NOTES]]: سطر واحد يوضح نتيجة التدقيق\n"
        "[[FINAL_ANSWER]]:\n"
        "الإجابة النهائية المعتمدة فقط."
    )

    critic_resp = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=(
            f"سؤال الباحث:\n{query}\n\n"
            f"الشواهد الموثقة:\n{context_text}\n\n"
            f"المسودة المطلوب تدقيقها:\n{draft_answer}"
        ),
        config=types.GenerateContentConfig(
            system_instruction=critic_prompt,
            temperature=0.0,
        ),
    )
    raw_critic = (critic_resp.text or "").strip()

    status_m = re.search(r"\[\[STATUS\]\]\s*:\s*(PASS|REVISED)", raw_critic, re.IGNORECASE)
    notes_m = re.search(r"\[\[NOTES\]\]\s*:\s*(.+?)(?=\[\[FINAL_ANSWER\]\]|$)", raw_critic, re.DOTALL)
    ans_m = re.search(r"\[\[FINAL_ANSWER\]\]\s*:?\s*(.+)$", raw_critic, re.DOTALL)

    final_answer = ans_m.group(1).strip() if ans_m and ans_m.group(1).strip() else draft_answer
    status_code = status_m.group(1).upper() if status_m else "PASS"
    notes = notes_m.group(1).strip() if notes_m else "تمت مطابقة الأبيات ونسبتها للمصادر المرفقة."

    return {
        "answer": final_answer,
        "verification_status": "مُصحّحة أكاديمياً" if status_code == "REVISED" else "موثّقة ومُدقّقة",
        "verification_notes": notes,
    }


# ==============================================================================
# 9. نقاط النهاية البرمجية (FastAPI Endpoints)
# ==============================================================================
@router.get("/health")
def rag_health() -> Dict[str, Any]:
    """فحص جاهزية المحرك السحابي وعدد الوثائق والمقاطع المفهرسة."""
    col = get_rag_collection()
    return {
        "status": "ready",
        "database": MONGODB_DB_NAME,
        "collection": RAG_COLLECTION_NAME,
        "indexed_documents": len(col.distinct("doc_id")),
        "total_chunks": col.count_documents({}),
        "model": GEMINI_MODEL,
        "embedding_model": EMBEDDING_MODEL,
    }


@router.post("/index")
def index_documents_endpoint(payload: BulkIndexRequest) -> Dict[str, Any]:
    """فهرسة قصيدة أو نص جديد مباشرة."""
    t0 = time.perf_counter()
    stats = upsert_poetry_documents(payload.documents)
    return {
        "message": "تمت الفهرسة بنجاح.",
        "elapsed_seconds": round(time.perf_counter() - t0, 2),
        **stats,
    }


@router.post("/sync-albidan")
def sync_albidan_archive() -> Dict[str, Any]:
    """
    مزامنة تلقائية شاملة لمجموعتي poems و glossary داخل قاعدة albidan_corpus
    مع التعرف الذكي على أسماء الحقول وتحويلها إلى متجهات بحث هجينة.
    """
    t0 = time.perf_counter()
    batch_docs: List[PoetryDocumentInput] = []

    # 1. قراءة مجموعة القصائد (poems)
    poems_col = get_source_collection("poems")
    poems_count = 0
    for raw in poems_col.find({}):
        doc_id = f"poem::{raw.get('_id')}"
        title = _pick_first_text(raw, ["title", "name", "heading", "poem_title"], "قصيدة حسانية")
        poet = _pick_first_text(raw, ["poet", "author", "poet_name", "shair"], "غير محدد")
        meter = _pick_first_text(raw, ["meter", "bahr", "batt", "rhythm", "wazn"], "غير محدد")
        genre = _pick_first_text(raw, ["genre", "type", "category", "form"], "شعر حساني")
        source = _pick_first_text(raw, ["source", "reference", "ref"], "أرشيف ALBIDAN - القصائد")
        notes = _pick_first_text(raw, ["lexical_notes", "analysis", "explanation", "notes", "summary"], "")
        content = _pick_first_text(raw, ["content", "text", "body", "poem", "lyrics", "verses"], "")

        if not content:
            content = _fallback_serialize(raw, {"_id", "embedding", "created_at", "updated_at"})
        if not content:
            continue

        batch_docs.append(
            PoetryDocumentInput(
                doc_id=doc_id,
                title=title,
                content=content,
                poet=poet,
                meter=meter,
                genre=genre,
                source=source,
                lexical_notes=notes,
            )
        )
        poems_count += 1

    # 2. قراءة مجموعة المعجم الحساني (glossary)
    glossary_col = get_source_collection("glossary")
    glossary_count = 0
    for raw in glossary_col.find({}):
        doc_id = f"glossary::{raw.get('_id')}"
        term = _pick_first_text(raw, ["term", "word", "kalima", "title", "entry", "name"], "مفردة حسانية")
        meaning = _pick_first_text(
            raw,
            ["meaning", "definition", "explanation", "description", "content", "text", "sharh"],
            "",
        )
        example = _pick_first_text(raw, ["example", "usage", "verse", "shahid"], "")
        source = _pick_first_text(raw, ["source", "reference"], "المعجم الحساني - ALBIDAN")

        body_parts = [f"المفردة الحسانية: {term}"]
        if meaning:
            body_parts.append(f"المعنى والشرح: {meaning}")
        if example:
            body_parts.append(f"مثال وشاهد: {example}")
        if len(body_parts) == 1:
            fallback = _fallback_serialize(raw, {"_id", "embedding", "created_at", "updated_at"})
            if fallback:
                body_parts.append(fallback)

        content = "\n".join(body_parts).strip()
        if not content:
            continue

        batch_docs.append(
            PoetryDocumentInput(
                doc_id=doc_id,
                title=term,
                content=content,
                poet="المعجم الحساني",
                meter="معجم",
                genre="معجم حساني",
                source=source,
            )
        )
        glossary_count += 1

    stats = upsert_poetry_documents(batch_docs)
    return {
        "message": "تمت مزامنة أرشيف ALBIDAN بنجاح.",
        "poems_scanned": poems_count,
        "glossary_scanned": glossary_count,
        "elapsed_seconds": round(time.perf_counter() - t0, 2),
        **stats,
    }


@router.post("/ask")
def ask_hassaniya_rag(req: HybridQueryRequest) -> Dict[str, Any]:
    """البحث الهجين (الدلالي + المعجمي الحساني) وتوليد إجابة أكاديمية موثّقة."""
    t_search = time.perf_counter()
    col = get_rag_collection()

    query_vector = embed_texts([req.query], task_type="RETRIEVAL_QUERY")[0]
    candidate_limit = max(12, req.top_k * 3)

    dense_hits = _vector_search(col, query_vector, req, limit=candidate_limit)
    lexical_hits = _atlas_lexical_search(col, req.query, req, limit=candidate_limit)
    fused_sources = _fuse_rrf(dense_hits, lexical_hits)[: req.top_k]
    search_seconds = round(time.perf_counter() - t_search, 2)

    t_gen = time.perf_counter()
    gen_result = generate_hassaniya_verified_answer(
        query=req.query,
        sources=fused_sources,
        enable_reflection=req.enable_reflection,
    )
    gen_seconds = round(time.perf_counter() - t_gen, 2)

    return {
        "query": req.query,
        "answer": gen_result["answer"],
        "verification": {
            "status": gen_result["verification_status"],
            "notes": gen_result["verification_notes"],
        },
        "timing": {
            "search_seconds": search_seconds,
            "generation_seconds": gen_seconds,
        },
        "sources": [
            {
                "doc_id": s.get("doc_id"),
                "title": s.get("title"),
                "poet": s.get("poet"),
                "meter": s.get("meter"),
                "genre": s.get("genre"),
                "source": s.get("source"),
                "confidence": s.get("confidence"),
                "excerpt": s.get("text"),
            }
            for s in fused_sources
        ],
    }
