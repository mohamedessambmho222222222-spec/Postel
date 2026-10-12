#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
أداة تسجيل البوستة: تقرأ الجواب المسكّن (OCR)، تقترح البيانات، وأنت تراجع وتحفظ في الإكسل.
كل حاجة بتشتغل على جهازك بس (127.0.0.1) ومفيش ملفات بتروح أي مكان.
"""
import copy
import difflib
import os
import re
import json
import math
import shutil
import sys
import tempfile
import threading
import time
import uuid
import urllib.parse
import webbrowser
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = Path(__file__).resolve().parent
_NEW_REG = BASE / "البوستة.xlsx"
XLSX = Path(os.environ.get("POSTAL_XLSX",
                           str(_NEW_REG)))
SCANS = Path(os.environ.get("POSTAL_SCANS", BASE / "scans"))
LOGO = BASE / "شركة_ميناء_القاهرة_الجوي.png"
BACKUPS = Path(os.environ.get("POSTAL_BACKUPS", BASE / "backups"))
INBOX = Path(os.environ.get("POSTAL_INBOX", BASE / "inbox"))
MODELS = BASE / "models"
AR_MODEL = MODELS / "arabic_PP-OCRv5_rec_mobile.onnx"
AR_DICT = MODELS / "ppocrv5_arabic_dict.txt"
# لا نعامل الخطاب كصورة واحدة؛ الحد يحمي الجهاز من ملف PDF ضخم بالخطأ ويمكن
# زيادته من البيئة عند الحاجة. 20 صفحة تغطي خطابات ADF المعتادة.
MAX_PAGES = int(os.environ.get("POSTAL_PAGES", "20"))
PORT = int(os.environ.get("POSTAL_PORT", "8765"))
SHEET = os.environ.get("POSTAL_SHEET", "")   # فاضي = أول ورقة في الملف
LISTS = "القوائم"
DATA_SHEETS = ("البوسطه الداخليه", "البوسطه الداخلية", "سجل البوستة", "الوارد")
DEC_SHEETS = ("القرارات",)
OUT_SHEETS = ("الصادر",)
EXTS = (".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff")
FREEZE = os.environ.get("POSTAL_FREEZE", "D2")   # تثبيت تاريخ الاستلام ورقم الوارد وتاريخ الختم
TMP = Path(tempfile.gettempdir()) / "postal_ocr_tmp"
TMP.mkdir(exist_ok=True)
INBOX.mkdir(exist_ok=True)
LOCK = threading.Lock()      # لحماية الكتابة على الإكسل
OCR_LOCK = threading.Lock()  # المحرك مش آمن للـ threads

AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
# Scale factor for digit cropping (default)
DEFAULT_SCALE = 5

# تطبيع خفيف (بنفس طول النص عشان المواضع تفضل صح) لأسماء الشهور.
_AR_LIGHT = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ة": "ه", "ى": "ي"})
AR_MONTHS = (("يناير", 1), ("فبراير", 2), ("مارس", 3), ("ابريل", 4), ("مايو", 5),
             ("يونيو", 6), ("يوليو", 7), ("اغسطس", 8), ("سبتمبر", 9), ("سبتمر", 9),
             ("سنتمبر", 9), ("اكتوبر", 10), ("نوفمبر", 11), ("ديسمبر", 12))
_MONTH_NUM = dict(AR_MONTHS)
MONTH_DATE_RE = re.compile(
    r"(\d{1,2})\s*[/\-.]?\s*(" + "|".join(m for m, _n in AR_MONTHS) + r")\s*[،,]?\s*(\d{4})")

# أقصى فرق أيام لتاريخ «من المتن»: الخطابات اللي بتتسجل كلها حديثة، وأي تاريخ
# قديم (تواريخ أحكام/قضايا) بيتساب فاضي بدل ما يتسجل غلط.
DATE_WINDOW_DAYS = int(os.environ.get("POSTAL_DATE_WINDOW", "400"))


# ------------------------------------------------------------------ OCR (RapidOCR + ONNX)
_OCR = None
_OCR_NOTE = []


class DocumentAIEngine:
    """واجهة ثابتة للمحركات: فصل قراءة المستند عن قواعد الحقول والواجهة.

    المحرك الافتراضي محلي بالكامل. يمكن لاحقًا إضافة PaddleOCR-VL أو GLM-OCR
    كتطبيق آخر لهذه الواجهة من غير تغيير الحفظ أو واجهة المراجعة.
    """
    name = "unknown"
    mode = "local"

    def analyze(self, paths):
        raise NotImplementedError


class RapidOCRDocumentAIEngine(DocumentAIEngine):
    name = "RapidOCR / PP-OCRv5 Arabic"

    def analyze(self, paths):
        return ocr_pages(paths)


_DOCUMENT_ENGINE = None


def document_ai_engine():
    global _DOCUMENT_ENGINE
    if _DOCUMENT_ENGINE is None:
        # لا نقبل اختيار محرك غير مثبّت بصمت؛ تجنّب أن تصبح نتيجة السجل
        # مرتبطة بخدمة Cloud أو نموذج مختلف بالخطأ.
        requested = os.environ.get("POSTAL_DOCUMENT_ENGINE", "rapidocr").lower()
        if requested != "rapidocr":
            _OCR_NOTE.append(
                f"المحرك المطلوب ({requested}) غير مفعّل بعد؛ تم استخدام RapidOCR المحلي.")
        _DOCUMENT_ENGINE = RapidOCRDocumentAIEngine()
    return _DOCUMENT_ENGINE


def ocr_engine():
    """محرك RapidOCR: نموذج العربية PP-OCRv5 من فولدر models (لو موجود).

    الدقة العالية هي الافتراضية: نحافظ على دقة صفحات السكانر قبل الكشف.
    POSTAL_ACCURACY=normal يتيح الوضع الأخف عند الحاجة.
    """
    global _OCR
    with OCR_LOCK:
        if _OCR is not None:
            return _OCR
        from rapidocr import RapidOCR
        from rapidocr.utils.typings import LangRec, ModelType, OCRVersion
        high = os.environ.get("POSTAL_ACCURACY", "high").lower() in ("high", "1", "yes")
        params = {"Global.use_cls": False,
                  "Global.text_score": 0.20 if high else 0.30,
                  "Global.max_side_len": 2600}
        if AR_MODEL.exists() and AR_DICT.exists():
            params.update({
                "Rec.model_path": str(AR_MODEL),
                "Rec.rec_keys_path": str(AR_DICT),
                "Rec.lang_type": LangRec.ARABIC,
                "Rec.ocr_version": OCRVersion.PPOCRV5,
                "Rec.model_type": ModelType.MOBILE,
            })
        else:
            _OCR_NOTE.append(
                "نموذج العربية مش موجود في فولدر models — القراءة هتستخدم النموذج "
                "الافتراضي (ضعيف مع العربي).")
        _OCR = RapidOCR(params=params)
        return _OCR


def try_qr(path: Path):
    """لو فيه QR/باركود على الورقة بنقراه -> رقم الوارد مضمون 100%.

    ملحوظة: cv2.imread بيفشل مع المسارات العربية على ويندوز، فبنقرأ بـ PIL
    (اللي بيتعامل مع يونيكود كويس) ونحولها لـ numpy.
    """
    try:
        import cv2
        import numpy as np
        from PIL import Image
        # قراءة آمنة للملفات (المسار العربي بيسبب مشاكل مع imread)
        with Image.open(path) as im:
            arr = np.array(im.convert("RGB"))
        det = cv2.QRCodeDetector()
        data, _pts, _ = det.detectAndDecode(arr)
        if not data:
            for k in (1.6, 2.5):        # تانيات بتكبير الصورة
                big = cv2.resize(arr, None, fx=k, fy=k)
                data, _pts, _ = det.detectAndDecode(big)
                if data:
                    break
        if data:
            return _parse_qr(data)
    except Exception:
        pass
    return None


def find_qr(paths):
    """أول QR موجود في أي صورة من الصور دي (الأصلية أو صفحات الـ PDF)."""
    seen = set()
    for p in paths or []:
        if not p:
            continue
        key = str(p)
        if key in seen:
            continue
        seen.add(key)
        try:
            res = try_qr(Path(p))
        except Exception:
            res = None
        if res:
            return res
    return None


def _parse_qr(raw):
    """تحويل نص الـ QR لأرقام.

    «وارد-2026-0863» -> الرقم التسلسلي (0863)
    «6024»            -> رقم وارد مباشر
    «2026-09-22»      -> تاريخ، مش رقم -> نتجاهله
    """
    s = (raw or "").strip()
    if not s:
        return None
    # تاريخ كامل مش رقم وارد
    if re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", s) or \
       re.fullmatch(r"\d{1,2}[-/]\d{1,2}[-/]\d{4}", s):
        return None
    # 1) صيغة الرقم التسلسلي: وارد-2026-0863  -> ناخد آخر جزء
    m = re.search(r"(?:وارد|صادر|ورود|رقم)\s*[-_]?\s*\d{0,4}[-_]\s*(\d{1,6})", s)
    if m:
        num = m.group(1)
        return {"number": num, "_number_conf": "qr", "_qr_raw": s,
                "_qr_kind": "serial"}
    # 2) رقم مباشر
    m = re.search(r"(\d{2,8})", s)
    if not m:
        return None
    num = m.group(1)
    if len(num) == 4 and re.fullmatch(r"(19|20)\d{2}", num):
        return None                      # سنة لوحدها مش رقم وارد
    return {"number": num, "_number_conf": "qr", "_qr_raw": s,
            "_qr_kind": "plain"}


def _deskew_angle(gray):
    """زاوية ميل الصفحة عبر minAreaRect لأكبر كتلة نص (0 لو مفيش ميل واضح)."""
    try:
        import cv2
        import numpy as np
        _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        coords = np.column_stack(np.where(bw > 0))
        if len(coords) < 1000:
            return 0.0
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = 90 + angle
        if abs(angle) < 0.5 or abs(angle) > 15:
            return 0.0
        return -angle
    except Exception:
        return 0.0


def _trim_scanner_blank(im):
    """Trim only large white outer bands, preserving page edges and all ink."""
    import numpy as np
    thumb = im.convert('L')
    thumb.thumbnail((1400, 1400))
    pixels = np.asarray(thumb)
    rows = np.flatnonzero((pixels < 220).mean(axis=1) >= .005)
    if not len(rows):
        return im
    padding = max(30, int(im.width * .025))
    top = max(0, int(rows[0] * im.height / thumb.height) - padding)
    bottom = min(im.height, int((rows[-1]+1) * im.height / thumb.height) + padding)
    if bottom-top >= im.height * .2 and im.height-(bottom-top) > im.height * .25:
        return im.crop((0, top, im.width, bottom))
    return im


def prepare_image(src: Path, out: Path):
    """تنظيف محسّن: EXIF + إزالة ميل + CLAHE + إزالة ضوضاء + تكبير + تباين.

    كله محلي بـ OpenCV/Pillow، ومن غير تحويل ثنائي قاسي عشان خط اليد الرفيع.
    يُتجاوز بالكامل عبر POSTAL_PREPROCESS=off عند الحاجة.
    """
    from PIL import Image, ImageOps
    im = Image.open(src)
    im = ImageOps.exif_transpose(im).convert("L")
    # Some scanners append several blank page lengths to a single image.
    # Trim only large, nearly white outer bands before OCR rescales the page.
    try:
        import numpy as np
        thumb = im.copy()
        thumb.thumbnail((1400, 1400))
        pixels = np.asarray(thumb)
        rows = np.flatnonzero((pixels < 220).mean(axis=1) >= 0.005)
        if len(rows):
            scale_y = im.height / thumb.height
            padding = max(30, int(im.width * 0.025))
            top = max(0, int(rows[0] * scale_y) - padding)
            bottom = min(im.height, int((rows[-1] + 1) * scale_y) + padding)
            if bottom - top >= im.height * 0.2 and im.height - (bottom - top) > im.height * 0.25:
                im = im.crop((0, top, im.width, bottom))
    except Exception:
        pass
    if os.environ.get("POSTAL_PREPROCESS", "on").lower() not in ("0", "off", "no"):
        try:
            import cv2
            import numpy as np
            arr = np.array(im)
            # إنقاذ الصفحات المحروقة من السكانر: 97% أبيض والحبر في أغمق 2%.
            # autocontrast العادي مابيقدرش (min=0 من الحواف) — نمدّ يدوي
            # بين الحبر (p1) والورقة (p50). البوابة mean>245 تحمي الصفحات
            # الطبيعية (≈182) والمتوسطة (Scan_0002≈235) من أي تغيير.
            if arr.mean() > 245 or float((arr >= 245).mean()) > 0.90:
                lo = float(np.percentile(arr, 1))
                hi = float(np.percentile(arr, 50))
                if hi - lo >= 60:
                    arr = np.clip(
                        (arr.astype(np.float32) - lo) * (255.0 / (hi - lo)),
                        0, 255).astype(np.uint8)
            ang = _deskew_angle(arr)
            if ang:
                h, w = arr.shape[:2]
                m = cv2.getRotationMatrix2D((w / 2, h / 2), ang, 1.0)
                arr = cv2.warpAffine(arr, m, (w, h), flags=cv2.INTER_LINEAR,
                                     borderMode=cv2.BORDER_REPLICATE)
            arr = cv2.fastNlMeansDenoising(arr, None, 7, 7, 21)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            arr = clahe.apply(arr)
            im = Image.fromarray(arr)
        except Exception:
            pass
    if im.width < 2600:
        k = 2600 / im.width
        im = im.resize((int(im.width * k), int(im.height * k)), Image.LANCZOS)
    im = ImageOps.autocontrast(im)
    im.save(out)


def document_quality(path: Path):
    """مؤشرات تحذيرية فقط لجودة الصفحة، من غير تغيير الأصل أو منع المراجعة."""
    try:
        from PIL import Image, ImageStat
        with Image.open(path) as im:
            gray = im.convert("L")
            stat = ImageStat.Stat(gray)
            mean = float(stat.mean[0])
            warnings = []
            if min(im.width, im.height) < 1000:
                warnings.append("دقة الصفحة منخفضة وقد تؤثر على القراءة.")
            if mean < 45:
                warnings.append("الصفحة داكنة جدًا وقد تؤثر على القراءة.")
            elif mean > 215:
                # عتبة عملية: سحبة محروقة (mean≈235، أبيض 86%) لم تُحذَّر عند 245،
                # بينما الطبيعية (mean≈182) تبقى بعيدة — فنحذّر مبكرًا بدل الصمت.
                try:
                    import numpy as _np
                    _arr = _np.asarray(gray)
                    _white = float((_arr >= 245).sum()) / float(_arr.size or 1)
                except Exception:
                    _white = 0.0
                if _white > 0.55:
                    warnings.append("الصفحة فاتحة جدًا وقد تؤثر على القراءة.")
                elif mean > 245:
                    warnings.append("الصفحة فاتحة جدًا وقد تؤثر على القراءة.")
            return {"width": im.width, "height": im.height,
                    "mean_luminance": round(mean, 1), "warnings": warnings}
    except Exception:
        return {"warnings": []}


def pdf_pages(src: Path, outdir: Path):
    """Render every PDF page, rejecting oversized files instead of truncating."""
    import fitz  # PyMuPDF
    outs = []
    with fitz.open(str(src)) as doc:
        if doc.needs_pass:
            raise ValueError("الملف محمي بكلمة مرور؛ ارفع نسخة قابلة للقراءة.")
        if len(doc) > MAX_PAGES:
            raise ValueError(f"المستند أكبر من {MAX_PAGES} صفحة؛ قسمه إلى ملفات أصغر.")
        for i in range(len(doc)):
            p = outdir / f"page{i}.png"
            doc[i].get_pixmap(dpi=200).save(str(p))
            outs.append(p)
    return outs


def image_pages(src: Path, outdir: Path):
    """Expand all pages of a multipage TIFF before OCR."""
    from PIL import Image, ImageSequence
    with Image.open(src) as image:
        count = getattr(image, "n_frames", 1)
        if count > MAX_PAGES:
            raise ValueError(f"المستند أكبر من {MAX_PAGES} صفحة؛ قسمه إلى ملفات أصغر.")
        if count == 1:
            return [src]
        paths = []
        for i, frame in enumerate(ImageSequence.Iterator(image)):
            path = outdir / f"page{i}.png"
            frame.convert("RGB").save(path)
            paths.append(path)
        return paths


def _lines_of(res):
    t = getattr(res, "txts", None)
    if t is None and isinstance(res, (list, tuple)):
        t = [it[1] for it in res if isinstance(it, (list, tuple)) and len(it) > 1]
    return [str(x).strip() for x in (t or []) if x]


def _res_lines(res):
    """أسطر OCR مع المواضع: [{text,x0,x1,y0,y1}]."""
    t = getattr(res, "txts", None)
    items = list(res) if isinstance(res, (list, tuple)) else []
    if t is None and isinstance(res, (list, tuple)):
        t = [it[1] for it in items if isinstance(it, (list, tuple)) and len(it) > 1]
    boxes = getattr(res, "boxes", None)
    boxes = list(boxes) if boxes is not None else []
    scores = getattr(res, "scores", None)
    scores = list(scores) if scores is not None else []
    out = []
    for i, s in enumerate(t or []):
        s = str(s or "").strip()
        if not s:
            continue
        d = {"text": s}
        b = boxes[i] if i < len(boxes) else None
        if b is not None:
            try:
                xs = [float(p[0]) for p in b]
                ys = [float(p[1]) for p in b]
                d.update(x0=min(xs), x1=max(xs), y0=min(ys), y1=max(ys))
            except Exception:
                pass
        score = scores[i] if i < len(scores) else None
        if score is None and i < len(items) and isinstance(items[i], (list, tuple)) \
                and len(items[i]) > 2:
            score = items[i][2]
        try:
            d["ocr_score"] = float(score)
        except (TypeError, ValueError):
            pass
        out.append(d)
    return out


def _run_engine(engine, image, **options):
    """Serialize inference and restore RapidOCR's mutable per-call switches."""
    with OCR_LOCK:
        switches = {name: getattr(engine, name) for name in ("use_det", "use_cls", "use_rec")
                    if hasattr(engine, name)}
        try:
            return engine(image, **options)
        finally:
            for name, value in switches.items():
                setattr(engine, name, value)


def _decision_body_readings(engine, image, top, bottom):
    """Read a printed decision body locally, retaining original page positions.

    A diagonal watermark can fragment otherwise clear printed rows. Horizontal
    ink concentration measures local skew; a light despeckle is an additional
    reading only, never a replacement for the archived image.
    """
    import cv2
    import numpy as np
    from PIL import ImageOps
    width, height = image.size
    body_top = max(0, int(top))
    body_bottom = min(height, int(bottom))
    crop = image.crop((0, body_top, width, body_bottom))
    if crop.height < 50:
        return []
    pad = max(16, int(width * 0.012))
    pixels = np.asarray(ImageOps.expand(crop, border=pad, fill=255))
    probe = np.asarray(crop)[:max(50, int(crop.height * 0.75))]
    scale = min(1, 1200 / max(width, 1))
    probe = cv2.resize(probe, None, fx=scale, fy=scale)
    ph, pw = probe.shape
    scored = []
    for angle in np.arange(-4, 4.01, 0.25):
        matrix = cv2.getRotationMatrix2D((pw / 2, ph / 2), float(angle), 1)
        rotated = cv2.warpAffine(probe, matrix, (pw, ph), borderValue=255)
        projection = (rotated < 150).sum(axis=1).astype(float)
        scored.append((float(np.mean(projection ** 2)), float(angle)))
    best_score, angle = max(scored)
    zero_score = next(score for score, candidate in scored if candidate == 0)
    if best_score < zero_score * 1.05 or abs(angle) < 0.5:
        angle = 0.0
    h, w = pixels.shape
    matrix = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1)
    inverse = cv2.invertAffineTransform(matrix)
    aligned = cv2.warpAffine(pixels, matrix, (w, h), borderValue=255)
    variants = [("local", aligned)]
    # Three pixels remove stippled watermark dots while substantial printed
    # strokes remain. The unchanged reading remains available beside it.
    ink = (aligned < 180).astype("uint8") * 255
    cleaned = 255 - cv2.morphologyEx(ink, cv2.MORPH_OPEN,
                                    np.ones((3, 3), dtype="uint8"))
    variants.append(("despeckled", cleaned))
    readings = []
    for variant, pixels in variants:
        for line in _res_lines(_run_engine(engine, pixels)):
            corners = np.array([[line["x0"], line["y0"]],
                                [line["x1"], line["y0"]],
                                [line["x1"], line["y1"]],
                                [line["x0"], line["y1"]]], dtype=float)
            original = corners @ inverse[:, :2].T + inverse[:, 2]
            line.update(x0=max(0, float(original[:, 0].min()) - pad),
                        x1=min(width, float(original[:, 0].max()) - pad),
                        y0=max(0, float(original[:, 1].min()) - pad + body_top),
                        y1=min(height, float(original[:, 1].max()) - pad + body_top),
                        _body_local_variant=variant, _body_skew_angle=angle)
            readings.append(line)
    return readings


def _refine_body_rows(engine, img, lines):
    """Re-read complete Arabic body rows when detection fragments a sentence."""
    from PIL import Image
    import numpy as np
    greeting = next((line for line in lines
                     if re.search(r"تحية|تحيه|السلام عليكم", line.get("text", ""))), None)
    # Decisions do not have a greeting. Re-read their operative articles,
    # rather than the legal preamble or the incoming stamp below the signature.
    article = next((line for line in lines
                    if re.search(r"مادة\s*(?:[أاإ]ول[ىي]|1|١)",
                                 line.get("text", ""))), None) if greeting is None else None
    if greeting is None and article is None:
        return lines
    with Image.open(img) as source:
        image = source.convert("L")
    width, height = image.size
    is_decision = greeting is None
    top = (float(article.get("y0", 0)) - max(5, (article["y1"] - article["y0"]) * 0.55) if is_decision
           else float(greeting.get("y1", 0)))
    closing = next((line for line in lines
                    if float(line.get("y0", 0)) > top
                    and re.search(r"وتفضلوا|مع خالص|التوقيع", line.get("text", ""))), None)
    bottom = (float(closing["y0"]) if closing
              else min(height, top + height * (0.35 if is_decision else 0.22)))
    body = [line for line in lines if top < float(line.get("y0", 0)) < bottom]
    if not body:
        return lines
    left = max(0, int(min(line["x0"] for line in body) - width * 0.02))
    right = min(width, int(max(line["x1"] for line in body) + width * 0.02))
    local_readings = (_decision_body_readings(engine, image, top, bottom)
                      if is_decision else [])
    slope = math.tan(math.radians(local_readings[0].get("_body_skew_angle", 0))) if local_readings else 0

    def row_center(line):
        return ((line["y0"] + line["y1"]) / 2
                - slope * (line.get("x0", 0) + line.get("x1", 0)) / 2)

    groups = []
    heights = sorted(max(12, line["y1"] - line["y0"]
                         - abs(slope) * (line["x1"] - line["x0"])) for line in body)
    row_tolerance = (max(8, min(44, heights[len(heights) // 2] * 0.45))
                     if is_decision else height * 0.013)
    for line in body:
        center = row_center(line)
        group = next((group for group in groups if abs(group[0] - center) < row_tolerance), None)
        if group is None:
            groups.append([center, [line]])
        else:
            group[1].append(line)
    replacements = {}
    removed = set()
    alternative_lines = []
    if any(line.get("ocr_score", 1) < 0.8 for line in body):
        body_top = max(0, int(top - 5))
        body_image = image.crop((0, body_top, width, min(height, int(bottom + 5))))
        alternative_lines = _res_lines(_run_engine(engine, np.asarray(body_image)))
        for alternative in alternative_lines:
            alternative["y0"] += body_top
            alternative["y1"] += body_top
    alternative_lines.extend(local_readings)
    for _, group in groups[:16 if is_decision else 6]:
        y0, y1 = min(line["y0"] for line in group), max(line["y1"] for line in group)
        crop = image.crop((left, int(y0), right, int(y1)))
        result = _run_engine(engine, np.asarray(crop), use_det=False, use_cls=False)
        readings = _res_lines(result)
        if len(readings) != 1:
            continue
        reading = readings[0]
        phrase = reading["text"]
        previous = " ".join(line["text"] for line in group)
        score = reading.get("ocr_score", 0)
        previous_score = sum(line.get("ocr_score", 0) for line in group) / len(group)
        alternatives = [candidate for candidate in alternative_lines
                        if abs(row_center(candidate) - sum(row_center(line) for line in group) / len(group)) < row_tolerance
                        and (candidate.get("ocr_score", 0) >= previous_score + 0.08
                             or (is_decision and candidate.get("ocr_score", 0) >= 0.85
                                 and len(candidate["text"]) >= len(previous) * 1.5))
                        and len(candidate["text"]) >= len(previous) * 0.85]
        if alternatives:
            candidate = max(alternatives, key=lambda item: (
                len(item["text"]) if is_decision else 0,
                item.get("ocr_score", 0)))
            if (candidate.get("ocr_score", 0) > score
                    or (is_decision and len(candidate["text"]) >= len(phrase) * 1.5)):
                reading = candidate
                phrase = reading["text"]
                score = reading.get("ocr_score", 0)
        recovered_fragment = is_decision and len(phrase) >= len(previous) * 1.5
        if (score < 0.85 or (score < previous_score - 0.04 and not recovered_fragment)
                or len(phrase) < len(previous) * 0.85):
            if is_decision and len(group) > 1:
                # Keep the recognized fragments in Arabic reading order even
                # when a new complete-row reading is not clear enough to use.
                first = group[0]
                replacements[id(first)] = {**first,
                    "text": " ".join(line["text"] for line in sorted(
                        group, key=lambda item: -item.get("x1", 0))),
                    "x0": min(line["x0"] for line in group),
                    "x1": max(line["x1"] for line in group),
                    "y0": y0, "y1": y1, "_body_row_joined": True}
                removed.update(id(line) for line in group[1:])
            continue
        # A complete row must recover missing text, or improve a fragmented row.
        if len(group) == 1 and len(phrase) < len(previous) + 8 and score < previous_score + 0.08:
            continue
        first = group[0]
        replacements[id(first)] = {**first, "text": phrase, "ocr_score": score,
                                   "x0": reading.get("x0", left), "x1": reading.get("x1", right),
                                   "y0": reading.get("y0", y0), "y1": reading.get("y1", y1),
                                   "_body_row_reread": True}
        removed.update(id(line) for line in group[1:])
    refined = [replacements.get(id(line), line) for line in lines if id(line) not in removed]
    if is_decision:
        # A crop can reveal a complete row that the full-page detector omitted.
        # Add only clear Arabic rows from the same body pixels and avoid a second
        # version of an already represented row.
        for candidate in sorted(alternative_lines, key=lambda item: -item.get("ocr_score", 0)):
            center = (candidate["y0"] + candidate["y1"]) / 2
            if (candidate.get("ocr_score", 0) < 0.90
                    or len(candidate["text"]) < 24
                    or _arabic_ratio(candidate["text"]) < 0.75
                    or not top < center < bottom):
                continue
            if any(abs(row_center(line) - row_center(candidate)) < row_tolerance
                   for line in refined if top < line.get("y0", 0) < bottom):
                continue
            refined.append({**candidate, "_body_row_reread": True})
        refined.sort(key=lambda line: (line.get("y0", 0), -line.get("x1", 0)))
    # The Arabic recognizer can omit Latin digits embedded in an Arabic row.
    # Require both a very clear isolated digit and the word رقم immediately
    # to its right before inserting anything into the extracted text.
    import cv2
    for line in refined:
        phrase = line.get('text', '')
        if not (top < float(line.get('y0', 0)) < bottom
                and re.search(r'رقم\s*[،,]', phrase)):
            continue
        lx, ly = max(0, int(line['x0']) - 5), max(0, int(line['y0']) - 8)
        row = image.crop((lx, ly, min(width, int(line['x1']) + 5), int(line['y1']) + 8))
        pixels = np.asarray(row)
        _, _, stats, _ = cv2.connectedComponentsWithStats((pixels < 170).astype('uint8'))
        digit_engine = digits_engine()
        candidates = []
        for x, y, w, h, area in stats[1:]:
            if not (.22 * row.height < h < .85 * row.height and .2 < w / h < .9):
                continue
            chip = row.crop((max(0, x-8), max(0, y-8), x+w+8, y+h+8))
            result = _run_engine(digit_engine, np.asarray(chip), use_det=False, use_cls=False)
            readings = _res_lines(result)
            if len(readings) != 1 or not re.fullmatch(r'[0-9]', readings[0]['text']):
                continue
            if readings[0].get('ocr_score', 0) < .995:
                continue
            prefix = row.crop((x+w+5, 0, row.width, row.height))
            if prefix.width < row.width * .12:
                continue
            reading = _run_engine(engine, np.asarray(prefix), use_det=False, use_cls=False)
            if any(re.search(r'رقم\s*$', str(tx)) for tx in getattr(reading, 'txts', None) or []):
                candidates.append(readings[0]['text'])
        if len(candidates) == 1:
            line['text'] = re.sub(r'رقم\s*(?=[،,])', 'رقم ' + candidates[0] + ' ', phrase, count=1)
            line['_body_digit_recovered'] = True
    return refined


def ocr_page(img: Path):
    """قراءة صفحة واحدة -> (نص, تحذيرات, أسطر بمواضع)."""
    warnings = list(_OCR_NOTE)
    try:
        engine = ocr_engine()
    except Exception as e:
        return "", [f"مفيش محرك OCR جاهز: {e}", "جرّب:  pip install rapidocr onnxruntime"], []
    try:
        res = _run_engine(engine, str(img))
    except Exception as e:
        return "", [f"فشلت القراءة: {e}"], []
    lines = _res_lines(res)
    try:
        lines = _refine_body_rows(engine, img, lines)
    except Exception as e:
        warnings.append(f"تعذرت مراجعة سطور المتن: {e}")
    return "\n".join(l["text"] for l in lines), warnings, lines


def run_ocr(img: Path):
    """قراءة صورة واحدة -> (نص, تحذيرات)."""
    text, warnings, _ = ocr_page(img)
    return text, warnings


def ocr_pages(paths):
    """قراءة أكتر من صفحة ودمج النص + كل الأسطر."""
    texts, warnings, all_lines = [], [], []
    for i, p in enumerate(paths):
        t0 = time.time()
        t, w, lines = ocr_page(p)
        dt = time.time() - t0
        if dt > 45:
            warnings.append(f"القراءة خدت {dt:.0f} ثانية — ممكن تقلل عدد الصفحات.")
        warnings += [f"صفحة {i + 1}: {x}" for x in w]
        if t.strip():
            from PIL import Image
            with Image.open(p) as page_image:
                page_size = [page_image.width, page_image.height]
            for line in lines:
                line.update(_page=i + 1, _page_size=page_size)
            texts.append(t)
        else:
            warnings.append(f"صفحة {i + 1}: لم يُستخرج نص؛ راجع وضوح الصفحة.")
        all_lines.append(lines)
    return "\n".join(texts), warnings, all_lines


# محرك ثاني خفيف للنصوص اللاتينية/الأرقام (مستخدم في مناطق الختم)
_DIG = None
DIG_DICT = MODELS / "ppocrv6_dict.txt"


def digits_engine():
    """نموذج الأرقام/اللاتيني المرفق (بلا أي تنزيل) لقراءة مناطق الختم."""
    global _DIG
    with OCR_LOCK:
        if _DIG is not None:
            return _DIG
        from rapidocr import RapidOCR
        params = {"Global.use_cls": False, "Global.text_score": 0.30,
                  "Global.max_side_len": 1400}
        if DIG_DICT.exists():
            params["Rec.rec_keys_path"] = str(DIG_DICT)
        _DIG = RapidOCR(params=params)   # Rec = النموذج المرفق PP-OCRv6_small
        return _DIG


def crop_digits_px(prepped: Path, x0, y0, x1, y1, workdir: Path, tag: str, scale=DEFAULT_SCALE):
    """قص منطقة (بكسلات) من الصفحة + تكبير + قراءتها بالنموذج الرقمي.

    القص مؤقت في الذاكرة فقط (ملف temp يُمسح فورًا) — لا يُحفظ أي قص
    في مجلد الورقة. الناتج المحفوظ لكل ورقة هو: الصورة الأصلية + Word.
    """
    from PIL import Image
    import tempfile
    im = Image.open(prepped)
    W, H = im.size
    x0, y0 = max(0, int(x0)), max(0, int(y0))
    x1, y1 = min(W, int(x1)), min(H, int(y1))
    if x1 - x0 < 20 or y1 - y0 < 12:
        return []
    crop = im.crop((x0, y0, x1, y1))
    k = max(1, min(scale, 1400 // max(1, crop.width)))
    if k > 1:
        crop = crop.resize((crop.width * k, crop.height * k))
    fd, tmp = tempfile.mkstemp(prefix=f"digits_{tag}_", suffix=".png")
    os.close(fd)
    try:
        crop.save(tmp)
        try:
            engine = digits_engine()
        except Exception:
            return []
        try:
            res = _run_engine(engine, tmp)
        except Exception:
            return []
        return [str(x).strip() for x in (getattr(res, "txts", None) or []) if x]
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


def crop_digits(prepped: Path, box, workdir: Path, tag: str, scale=3):
    """قص منطقة نسبية (0..1) — للتوافق مع الأدوات القديمة."""
    from PIL import Image
    im = Image.open(prepped)
    W, H = im.size
    return crop_digits_px(prepped, box[0] * W, box[1] * H, box[2] * W, box[3] * H,
                          workdir, tag, scale)


def crop_date_band_px(prepped: Path, x0, y0, x1, y1, workdir: Path, tag: str,
                      scale=3, binarize=False):
    """قص شريط التاريخ + تكبير + قراءة بالمحرك العربي (بيقرأ ٠١٢٣ صح).

    نموذج الأرقام اللاتيني (digits_engine) مش بيقرأ الأرقام العربية، لكن
    لما الشريط يتحول لأبيض/أسود نضيف (Otsu) بتبقى قراءة السنة أوضح —
    خصوصاً لما الـdetector يقفل على التاريخ وبيرجع جزء منه بس.

    القص مؤقت (ملف temp يُمسح فورًا) — لا يُحفظ أي قص في مجلد الورقة.
    """
    from PIL import Image
    import tempfile
    im = Image.open(prepped).convert("L")
    page_w, page_h = im.size
    x0, y0 = max(0, int(x0)), max(0, int(y0))
    x1, y1 = min(page_w, int(x1)), min(page_h, int(y1))
    if x1 - x0 < 20 or y1 - y0 < 10:
        return []
    crop = im.crop((x0, y0, x1, y1))
    k = max(1, min(scale, 1600 // max(1, crop.width)))
    if k > 1:
        crop = crop.resize((crop.width * k, crop.height * k), Image.LANCZOS)
    if binarize:
        try:
            import cv2
            import numpy as np
            arr = np.array(crop)
            _thr, otsu = cv2.threshold(arr, 0, 255,
                                       cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            crop = Image.fromarray(otsu)
        except Exception:
            pass
    fd, tmp = tempfile.mkstemp(prefix=f"dateband_{tag}_", suffix=".png")
    os.close(fd)
    try:
        try:
            crop.save(tmp)
        except Exception:
            return []
        try:
            engine = ocr_engine()
        except Exception:
            return []
        try:
            res = _run_engine(engine, tmp)
        except Exception:
            return []
        return [str(x).strip() for x in (getattr(res, "txts", None) or []) if x]
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


def digits_near(prepped: Path, anchor, workdir: Path, tag: str,
                left=0.55, right_pad=0.02, up=0.03, down=0.12):
    """الأرقام حوالين سطر الختم (عادة على الشمال منه). المواضع بالبكسل."""
    if anchor is None:
        return []
    from PIL import Image
    W, H = Image.open(prepped).size
    x0 = max(0, anchor["x0"] - W * left)
    x1 = min(W, anchor.get("x0", W) + W * right_pad)
    cy = (anchor["y0"] + anchor["y1"]) / 2
    y0 = max(0, cy - H * up)
    y1 = min(H, cy + H * down)
    return crop_digits_px(prepped, x0, y0, x1, y1, workdir, tag)


# ------------------------------------------------------------------ استخراج الحقول
def _valid(y, m, d):
    try:
        if 2000 <= y <= 2100:
            return date(y, m, d)
    except ValueError:
        pass
    return None


def _compact_date(run):
    """أرقام ملزقة من غير فواصل (بيحصل في القراءة العربية): ٢٠٢٦٩١٥ -> 2026-09-15."""
    n = len(run)
    if n < 6 or n > 8:
        return None
    # الأولوية للتقسيم الطبيعي: سنة 4 + شهر 2 + يوم 2 (مثال 20261005 -> 2026-10-05).
    # التقسيمات الأقصر (شهر/يوم برقم واحد) تيجي بعدها عشان متاكلش أرقام الشهر.
    order = []
    if n == 8:
        order = [(4, 2, 2), (4, 1, 3), (4, 3, 1), (4, 2, 1), (4, 1, 2),
                 (2, 2, 4), (2, 1, 5), (1, 2, 5), (1, 1, 6)]
    else:
        order = [(a, b, n - a - b) for a in (4, 2, 1) for b in (2, 1)]
    for a, b, c in order:
            if not 1 <= c <= 4:
                continue
            p1, p2, p3 = int(run[:a]), int(run[a:a + b]), int(run[a + b:])
            tries = []
            if a == 4:
                tries += [_valid(p1, p2, p3)]
            if c == 4:
                tries += [_valid(p3, p2, p1), _valid(p3, p1, p2)]
            for d in tries:
                if d:
                    return d
    # 6 أرقام بسنة كاملة وشهر/يوم ملزوقين (مثال: 202696 = 2026/9/6 أو 2026/06/9)
    if n == 6:
        y, rest = int(run[:4]), run[4:]
        for mm, dd in ((int(rest[0]), int(rest[1])),):
            for m, day in ((mm, dd), (dd, mm)):
                if 1 <= m <= 12 and 1 <= day <= 31:
                    got = _valid(y, m, day)
                    if got:
                        return got
    return None


def _short_year_date(run):
    """سنة من رقمين + شهر + يوم (٥ أرقام): ٢٦٨١١ -> 2026-08-11.

    بيحصل لما الـrecognizer يضيّع أول رقمين من «٢٠٢٦» فيقرأ «٢٦»، فتبقى
    السنة ناقصة. بنكمّل القرن بشكل متسق (٠٠-٤٩ -> 2000s، ٥٠-٩٩ -> 1900s)
    وبنتأكد إن الشهر واليوم مناسبين قبل القبول.
    """
    run = str(run)
    if len(run) != 5 or not run.isdigit():
        return None
    yy, mm, dd = int(run[:2]), int(run[2:4]), int(run[4:])
    if not 1 <= mm <= 12 or not 1 <= dd <= 31:
        return None
    return _valid(2000 + yy if yy < 50 else 1900 + yy, mm, dd)


def _date_from_fragment(text):
    """تاريخ من قراءة ناقصة: ٦-٨ أرقام أولاً، وبعدها ٥ أرقام (سنة من رقمين)."""
    digits = re.sub(r"\s+", "", str(text).translate(AR_DIGITS))
    # A recognizer may retain the year separator but merge month and day.
    # Preserve the complete year; never infer it from today's date.
    partial = re.fullmatch(r"((?:20|21)\d{2})[/٬،,](\d{2,4})", digits)
    if partial:
        got = _compact_date("".join(partial.groups()))
        if got:
            return got
    split_day = re.search(r"(?<!\d)((?:20|21)\d{2})(\d{1,2})[/٬،,](\d{1,2})(?!\d)", digits)
    if split_day:
        got = _valid(*(int(value) for value in split_day.groups()))
        if got:
            return got
    runs = sorted(re.findall(r"\d{5,8}", digits), key=len, reverse=True)
    for run in runs:
        got = _compact_date(run) or _short_year_date(run)
        if got:
            return got
    return find_date(text)


def _month_date_at(s):
    """تاريخ مكتوب فيه اسم الشهر بالحروف: «٣٠ سبتمبر ٢٠٢٦» -> (الموضع، التاريخ)."""
    s = s.translate(AR_DIGITS).translate(_AR_LIGHT)
    best = None
    for m in MONTH_DATE_RE.finditer(s):
        d = _valid(int(m.group(3)), _MONTH_NUM[m.group(2)], int(m.group(1)))
        if d and (best is None or m.start() < best[0]):
            best = (m.start(), d)
    return best


def find_month_date(s):
    m = _month_date_at(s)
    return m[1] if m else None


# «سبتمبر ٢٠٢٦» / «٢٠٢٦ سبتمبر» — شهر وسنة بس من غير يوم (بيحصل في ترويسة بعض الجهات)
MONTH_YEAR_RE = re.compile(
    r"(" + "|".join(m for m, _n in AR_MONTHS) + r")\D{0,12}?(\d{4})"
    r"|(\d{4})\D{0,12}?(" + "|".join(m for m, _n in AR_MONTHS) + r")")


def find_month_year(s):
    """«سبتمبر ٢٠٢٦» -> (الموضع، التاريخ) — بياخد أول يوم في الشهر."""
    s = s.translate(AR_DIGITS).translate(_AR_LIGHT)
    best = None
    for m in MONTH_YEAR_RE.finditer(s):
        if m.group(2):                      # شهر ثم سنة
            name, year = m.group(1), int(m.group(2))
        else:                               # سنة ثم شهر
            year, name = int(m.group(3)), m.group(4)
        d = _valid(year, _MONTH_NUM[name], 1)
        if d and (best is None or m.start() < best[0]):
            best = (m.start(), d)
    return best


def _fresh(d):
    """التاريخ قريب من النهاردة؟ (بيمنع تسجيل تواريخ أحكام/قضايا قديمة غلط)."""
    if not d:
        return False
    days = (date.today() - d).days
    return -30 <= days <= DATE_WINDOW_DAYS


def _valid_stamp_date(d):
    """A located incoming stamp may be historical; reject implausible future years."""
    return bool(d and d.year >= 2000 and (d - date.today()).days <= 30)


def find_date(s):
    s = s.translate(AR_DIGITS)
    pats = [(r'(\d{4})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{1,2})', "ymd"),
            (r'(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{4})', "dmy")]
    best = None
    for pat, kind in pats:
        for m in re.finditer(pat, s):
            a, b, c = (int(x) for x in m.groups())
            d = _valid(a, b, c) if kind == "ymd" else _valid(c, b, a)
            if d and (best is None or m.start() < best[0]):
                best = (m.start(), d)
    mon = _month_date_at(s)         # تاريخ بالحروف (سبتمبر/أكتوبر...)
    if mon and (best is None or mon[0] < best[0]):
        best = mon
    if best:
        return best[1]
    for run in re.findall(r'\d{6,8}', s):   # تاريخ ملزق من غير فواصل
        d = _compact_date(run)
        if d:
            return d
    return None


def _positional_candidates(lines_info):
    """أسطر OCR اللي عندها مواضع: [(idx, {text,x0,x1,y0,y1})]."""
    band = []
    for i, l in enumerate(lines_info):
        if not isinstance(l, dict):
            continue
        tx = str(l.get("text", "")).strip()
        if not tx:
            continue
        d = dict(l)
        d["text"] = tx
        band.append((i, d))
    return {"_band": band}


def _arabic_ratio(s):
    if not s:
        return 0.0
    letters = [c for c in s if c.isalpha()]
    if not letters:
        return 0.0
    return sum(1 for c in letters if 0x0600 <= ord(c) <= 0x06FF) / len(letters)


# قاموس الجهات: الاسم العربي الصحيح + أخطاء OCR الشائعة + المقابل الإنجليزي.
# بيتستخدم تصحيحاً فقط — لو الاسم مش موجود بيتساب زي ما هو.
ENTITY_FIXES = (
    # (الاسم الصحيح، المقابل الإنجليزي، كلمات عربية مميزة)
    ("شركة ميناء القاهرة الجوي", "cairo airport",
     ("ميناء", "القاهره", "القاهرة", "الجوي")),
    ("الشركة المصرية القابضة للمطارات", "egyptian holding",
     ("القابضه", "القابضة", "المطارات", "المصري")),
    ("رئيس قطاع الشؤون القانونية", "head of legal",
     ("رئيس", "الشئون", "الشؤون", "القانونيه")),
    ("قطاع الشؤون القانونية", "legal affairs",
     ("الشئون", "الشؤون", "القانونيه", "القانونية")),
    ("قطاع التنمية الإدارية والموارد البشرية", "admin development",
     ("التنميه", "التنمية", "الاداريه", "الموارد", "البشريه")),
    ("قطاع العمليات", "operations", ("العمليات", "الطيران")),
    ("قطاع الهندسة والتجهيزات", "engineering", ("الهندسه", "التجهيزات")),
    ("إدارة الموارد البشرية", "human resources", ("الموارد", "البشريه")),
    ("الإدارة العامة للمشتريات والمخازن", "procurement",
     ("المشتريات", "المخازن", "التوريدات")),
    ("الإدارة العامة للمالية", "finance", ("الماليه", "المالية", "الحسابات")),
)

_ENTITY_AR = [e[0] for e in ENTITY_FIXES]

# أخطاء OCR الشائعة في العربي (تُطبق قبل المطابقة عشان ترفع الدقة).
# مثال: «القاھرة» -> «القاهرة»، «الشئون» -> «الشؤون».
_OCR_CHAR_FIXES = (
    ("ھ", "ه"), ("ہ", "ه"), ("ۃ", "ة"), ("گ", "ك"), ("ڤ", "ف"),
    ("پ", "ب"), ("چ", "ج"),
)
_OCR_WORD_FIXES = (
    ("القاھرة", "القاهرة"), ("القاهره", "القاهرة"), ("الجوى", "الجوي"),
    ("الشئون", "الشؤون"), ("القانونيه", "القانونية"), ("الاداريه", "الإدارية"),
    ("التنميه", "التنمية"), ("البشريه", "البشرية"), ("الماليه", "المالية"),
    ("القابضه", "القابضة"), ("المصرى", "المصري"), ("الهندسه", "الهندسة"),
    ("مديريه", "مديرية"), ("ادارة", "إدارة"), ("الاستاذ", "الأستاذ"),
    ("التنغيذي", "التنفيذي"), ("الصياحي", "الصباحي"), ("الصبياحي", "الصباحي"),
    ("بدلأمن", "بدلاً من"),
)


def _fix_ocr_spelling(s):
    s = s or ""
    for bad, good in _OCR_CHAR_FIXES:
        s = s.replace(bad, good)
    for bad, good in _OCR_WORD_FIXES:
        s = s.replace(bad, good)
    return s


def recipient_job_title(value):
    value = _fix_ocr_spelling(str(value or '')).strip()
    match = re.search(r'\b(?:رئيس|نائب|مدير|المدير|مساعد|وكيل|أمين|امين|محافظ|وزير|الوزير|عميد|العميد|سكرتير|مسؤول|مسئول|إدارة|ادارة|قطاع)\b', value)
    return value[match.start():].strip(' /،,:') if match else ''


def fix_entity(s):
    """يصحّح اسم الجهة لو كان قريب جدًا من واحد معروف."""
    s = _fix_ocr_spelling((s or "").strip())
    if not s or len(s) < 4:
        return s
    # 1) مقابل إنجليزي معروف -> عربي
    low = s.lower()
    for ar, en, _kw in ENTITY_FIXES:
        if len(en) >= 6 and en in low:
            return ar
    if _arabic_ratio(s) < 0.4:
        return s          # إنجليزي مش معروف -> سيبه زي ما هو
    # 2) قريب جدًا (تشابه الحروف) أو فيه كلمات مميزة -> الاسم الصحيح
    n = _norm(s)
    best, best_score = None, 0.0
    for ar, _en, kw in ENTITY_FIXES:
        na = _norm(ar)
        r = difflib.SequenceMatcher(None, n, na).ratio()
        hits = sum(1 for k in kw if _norm(k) in n)
        if hits:
            r = min(1.0, r + 0.06 * hits)      # كلمات مريحة بتقوّي المطابقة
        if r > best_score:
            best, best_score = ar, r
    return best if best_score >= 0.80 else s


_DECISION_ARTICLE_RE = re.compile(
    r"(?:المادة|مادة|مسادة)\s*(?:[أاإ]ول[ىي]|ث(?:اني|الث|الئ|لئ|ئ)[ةه]|رابع[ةه]|"
    r"خامس[ةه]|سادس[ةه]|[0-9٠-٩]+)\s*[:：/\-–]?")
_DECISION_ACTION_RE = re.compile(
    r"(?<![\u0621-\u064a])(?:إنهاء\s+تكليف|انهاء\s+تكليف|إنهاء\s+ندب|انهاء\s+ندب|"
    r"محو\s+الجزاء|إلغاء\s+الجزاء|الغاء\s+الجزاء|مجازاة|"
    r"تكليف|يكلف|انتداب|ندب|يندب|ترقية|يرقى|تعيين|تسوية|نقل|منح|صرف)(?![\u0621-\u064a])")


def _is_decision_document(lines):
    """Require a decision heading; a letter quoting a decision is still a letter."""
    heading = any(re.match(r"^(?:تابع\s+)?قرار(?:\s|$)", str(line).strip())
                  for line in lines[:25])
    articles = any(_DECISION_ARTICLE_RE.search(str(line)) for line in lines)
    greeting = any(re.search(r"تحية|تحيه|السلام عليكم", str(line))
                   for line in lines[:25])
    return heading and articles and not greeting


def _decision_sender(lines):
    """Read the issuing role from a decision's heading, without personal names."""
    if not _is_decision_document(lines):
        return ""
    heading = []
    for line in lines[:30]:
        if re.search(r"بعد\s*[اإأ]ل?اطلاع|قانون|وعلى", line):
            break
        heading.append(re.sub(r"^(?:تابع\s+)?قرار\s*", "", line).strip())
    roles = ("مكتب رئيس مجلس الإدارة", "رئيس مجلس الإدارة")
    candidates = []
    for line in heading:
        normalized = re.sub(r"\s+", "", _norm(line))
        for role in roles:
            wanted = re.sub(r"\s+", "", _norm(role))
            score = difflib.SequenceMatcher(None, normalized, wanted).ratio()
            if score >= 0.80:
                candidates.append((score + (0.03 if role.startswith("مكتب") else 0), role))
    if not candidates:
        return ""
    role = max(candidates)[1]
    # This is an issuing title in the heading, not a name inferred from a signatory.
    if (not role.startswith("مكتب")
            and any(re.search(r"العضو\s+المنتدب", line) for line in heading)):
        role += " والعضو المنتدب"
        if any(re.search(r"التنفيذ[ىي]", line) for line in heading):
            role += " التنفيذي"
    return role


def _decision_subject(lines, limit=320):
    """Summarize the operative article, not the logo or legal preamble.

    All phrases come from OCR. In particular, a missing employee name or job
    title is not supplied from another decision or from a reviewed sample.
    """
    if not _is_decision_document(lines):
        return ""
    joined = " ".join(str(line).strip() for line in lines)
    markers = list(_DECISION_ARTICLE_RE.finditer(joined))
    first = next((marker for marker in markers
                  if re.search(r"[أاإ]ول[ىي]|(?<!\d)[1١](?!\d)", marker.group(0))), None)
    if first is None:
        # Continuation pages can contain only a promotion table and article two.
        # Its group heading and promoted job title are the available subject.
        group = next((re.search(r"المجموعة\s+النوعية.+", line).group(0)
                      for line in lines
                      if re.search(r"المجموعة\s+النوعية.+", line)), "")
        promotion = next((re.search(r"المرق[وؤ]ن\s+لوظائف\s+(.+)", line)
                          for line in lines
                          if re.search(r"المرق[وؤ]ن\s+لوظائف\s+(.+)", line)), None)
        if group and promotion:
            return ("ترقية العاملين في " + _fix_ocr_spelling(group)
                    + " إلى وظائف " + _fix_ocr_spelling(promotion.group(1)))[:limit]
        return ""
    pieces = []
    for index, marker in enumerate(markers):
        if marker.start() < first.start():
            continue
        end = markers[index + 1].start() if index + 1 < len(markers) else len(joined)
        article = joined[marker.end():end]
        article = re.split(r"التوقيع|رقم\s+الوارد|"
                           r"E-mail|Tel[.:]", article, maxsplit=1, flags=re.I)[0]
        action = _DECISION_ACTION_RE.search(article)
        if not action:
            continue
        phrase = article[action.start():]
        phrase = re.split(r"وذلك|كما\s+هو\s+وارد|رقم\s*العامل|"
                          r"بالقرار\s+رقم|بالإضافة\s+إلى", phrase, maxsplit=1)[0]
        phrase = _fix_ocr_spelling(phrase)
        phrase = re.sub(r"^يكلف\b", "تكليف", phrase)
        phrase = re.sub(r"^يندب\b", "ندب", phrase)
        # Correct only close readings of common occupational/organization
        # vocabulary. Personal names and document numbers are not in this list.
        vocabulary = ("المحاسب", "المشروعات", "التراخيص", "التمويل", "المحاسبة",
                      "المطارات", "الشركة", "التعاقدات", "التخصصية")
        def tidy_word(match):
            word = match.group(0)
            if len(word) < 6 or not word.startswith(("ال", "بال", "لل")):
                return word
            prefix = "ب" if word.startswith("بال") else "ل" if word.startswith("لل") else ""
            probe = word[len(prefix):]
            if prefix == "ل":
                probe = "ا" + probe
            closest = difflib.get_close_matches(_norm(probe), [_norm(w) for w in vocabulary],
                                               n=1, cutoff=0.86)
            if not closest:
                return word
            fixed = next(w for w in vocabulary if _norm(w) == closest[0])
            return prefix + (fixed[1:] if prefix == "ل" else fixed)
        phrase = re.sub(r"[\u0621-\u064a]+", tidy_word, phrase)
        phrase = re.sub(r"\bمدير\s+غام\b", "مدير عام", phrase)
        phrase = re.sub(r"\bكب[يى]ه?ر\b", "كبير", phrase)
        if re.match(r"(?:[إا]نهاء\s+(?:تكليف|ندب)|تكليف|ندب)\b", phrase):
            # The employee's current grade can occupy most of the summary.
            # Keep their name, the action, the new function and the organization.
            phrase = re.sub(r"\s+كبير\s+.+?(?=من\s+القيام|للقيام|لقيام|للعمل)", " ", phrase)
            phrase = re.sub(r"\s+محاسب\s+أول\s+.+?(?=\s+بشركة)", " ", phrase)
        # Restore common word boundaries lost by recognition, only in the summary.
        phrase = re.sub(r"(?<!\s)(?=السيد|بوظيفة|بالدرجة|بالمجموعة|بخصم)", " ", phrase)
        phrase = re.sub(r"(مجازاة|السيد)(?=[\u0621-\u064a])", r"\1 ", phrase)
        phrase = re.sub(r"(بوظيفة)(?=[\u0621-\u064a])", r"\1 ", phrase)
        phrase = re.sub(r"\s+", " ", phrase).strip(" ،,;؛./-")
        if phrase:
            pieces.append(phrase)
        if len(pieces) >= 2:
            break
    if not pieces:
        return ""
    subject = "؛ ".join(pieces)
    if len(subject) > limit:
        subject = subject[:limit + 1].rsplit(" ", 1)[0].rstrip(" ،,;؛")
        # A duration is meaningful and may occur after a long occupational title.
        duration = re.search(r"لمدة\s+(?:عام(?:ين)?|سنة|سنتين|\d+\s+(?:أشهر|شهور|سنوات))", pieces[0])
        if duration and duration.group(0) not in subject:
            subject = subject[:limit - len(duration.group(0)) - 1].rsplit(" ", 1)[0]
            subject += " " + duration.group(0)
    return subject


def _pick_sender(band_lines, pos_info):
    """الجهة المرسلة: عربي أولاً (بالتطبيع)، يمين فوق، وأطول سطر."""
    band = pos_info.get("_band", [])
    org = (r'(شركة|هيئة|وزارة|مديرية|مديريه|إدارة|ادارة|الجهاز|مصلحة|محافظة|'
           r'مجلس|جامعة|نقابة|بنك|قطاع|Company|Authority|Ministry|Department)')
    skip = r'^(السيد|السادة|الأستاذ|الاستاذ|السيدة|تحية|Mr|Dear)'
    cands = []
    for i, l in band:
        tx = l["text"]
        if not tx or re.match(skip, tx) or not re.search(org, tx):
            continue
        ar = _arabic_ratio(tx)
        x1 = float(l.get("x1", 0) or 0)
        y0 = float(l.get("y0", 10 ** 9) or 10 ** 9)
        right_bias = 2.0 if x1 > 0 else 0.0
        top_bias = 1.0 if y0 > 0 else 0.0
        score = ar * 3.0 + right_bias + top_bias + min(len(tx), 60) / 60.0
        cands.append((score, i, tx[:120]))
    if not cands:
        return ""
    cands.sort(key=lambda x: -x[0])
    return cands[0][2]


def _same_row(line, others, tol=0.6):
    """أسطر على نفس الشريط الأفقي (للجهة العربية جنب الإنجليزية)."""
    try:
        ya, yb = float(line.get("y0", -1)), float(line.get("y1", -1))
    except (TypeError, ValueError):
        return []
    out = []
    for i, l in others:
        try:
            y0, y1 = float(l.get("y0", -2)), float(l.get("y1", -2))
        except (TypeError, ValueError):
            continue
        h = max(yb - ya, y1 - y0, 1.0)
        if abs((ya + yb) / 2 - (y0 + y1) / 2) <= h * tol:
            out.append((i, l))
    return out


def _same_col(line, others, tol=0.5):
    """أسطر تحت/فوق بعضها بنفس العرض (أعمدة الختم)."""
    try:
        xa, xb = float(line.get("x0", -1)), float(line.get("x1", -1))
    except (TypeError, ValueError):
        return []
    out = []
    for i, l in others:
        try:
            x0, x1 = float(l.get("x0", -2)), float(l.get("x1", -2))
        except (TypeError, ValueError):
            continue
        w = max(xb - xa, x1 - x0, 1.0)
        if abs((xa + xb) / 2 - (x0 + x1) / 2) <= w * tol:
            out.append((i, l))
    return out


def stamp_rows(pos_info):
    """أسطر الختم: اللي فيها 'وارد' أو 'صادر' أو مشتقاتها."""
    rows = []
    for i, l in pos_info.get("_band", []):
        n = _norm(l["text"])
        if not n:
            continue
        if ("وارد" in n or "صادر" in n or "المرفقات" in n or "تنميه" in n):
            rows.append((i, l))
    return rows


def stamp_value(prepped, anchor, workdir, tag):
    """أقرب أرقام حوالين سطر الختم (يمين/شمال/تحت) + تواريخ النموذج الرقمي."""
    from PIL import Image
    W, H = Image.open(prepped).size
    nums, dates = [], []
    zones = [
        ("left", max(0, anchor["x0"] - W * 0.55), max(0, anchor["y0"] - H * 0.03),
         min(W, anchor.get("x0", W) + W * 0.03), min(H, anchor["y1"] + H * 0.14)),
        ("self", anchor["x0"], anchor["y0"], anchor["x1"], anchor["y1"]),
        ("below", anchor["x0"] - W * 0.1, anchor["y1"],
         anchor["x1"] + W * 0.1, anchor["y1"] + H * 0.10),
    ]
    for zn, x0, y0, x1, y1 in zones:
        for s in crop_digits_px(prepped, x0, y0, x1, y1, workdir, f"{tag}-{zn}"):
            dd = find_date(s)
            if dd:
                dates.append(dd.isoformat())
            for tok in re.findall(r"\d[\d/\- ]{0,12}", s.translate(AR_DIGITS)):
                v = tok.strip(" /-")
                if v and not find_date(v):
                    nums.append(v)
    out = {"numbers": [], "dates": []}
    for v in nums:
        if v not in out["numbers"]:
            out["numbers"].append(v)
    for v in dates:
        if v not in out["dates"]:
            out["dates"].append(v)
    return out


def _stamp_anchors(pos):
    """أسطر الختم السفلي (الوارد) والعلوي الشمالي (الصادر) من المواضع."""
    rows = stamp_rows(pos)
    bottom, top = [], []
    for i, l in rows:
        n = _norm(l["text"])
        y = float(l.get("y0", 0) or 0)
        if y > 0:
            (bottom if "وارد" in n or y > 1800 or "المرفقات" in n or "تنميه" in n
             else top).append((i, l))
        else:
            (bottom if "وارد" in n else top).append((i, l))
    return bottom, top


def _same_col_numbers(anchor, pos, prefer_year=False):
    """أقرب أرقام بنفس العمود حوالين سطر الختم (القيمة شمال العنوان عادة)."""
    same = _same_col(anchor, pos["_band"])
    out = []
    for _i, l in same:
        s = str(l["text"]).translate(AR_DIGITS)
        for tok in re.findall(r"\d[\d/\- ]{0,14}", s):
            v = tok.strip(" /-")
            if not v or find_date(v):
                continue
            out.append(v)
    return out


def _anchor_number(anchor_rows, lines, pos):
    """رقم الختم: الأقرب عدديًا بنفس العمود، وإلا أول مرشح رقمي."""
    for _i, l in anchor_rows:
        for v in _same_col_numbers(l, pos):
            return v, "high"
    for _i, l in anchor_rows:
        s = str(l["text"]).translate(AR_DIGITS)
        for tok in re.findall(r"\d[\d/\- ]{0,14}", s):
            v = tok.strip(" /-")
            if v and not find_date(v):
                return v, "medium"
    return "", "low"


def _row_date(same):
    texts = " ".join(str(l["text"]) for _i, l in same)
    dd = find_date(texts)
    return dd.isoformat() if dd else ""


def _anchor_date(lines):
    """أول تاريخ جنب سطر فيه «التاريخ» في أول سادس الورقة (بأخطاء OCR الشائعة)."""
    labels = ("تاريخ", "تاربخ", "تاريح", "ثاريخ")
    for index, l in enumerate(lines[:9]):
        if any('وارد' in _norm(str(v)) for v in lines[max(0,index-2):index+1]):
            if not re.search(r'تاريخ (?:الخطاب|الكتاب|التحرير)', str(l)):
                continue
        if not any(k in _norm(str(l)) for k in labels):
            continue
        dd = find_date(str(l))
        if dd:
            return dd.isoformat()
    return ""


def _page_dims(lines_info):
    W = max([float(l.get("x1", 0) or 0) for l in lines_info] or [1])
    H = max([float(l.get("y1", 0) or 0) for l in lines_info] or [1])
    return max(W, 1.0), max(H, 1.0)


def _page_regions(lines_info, W, H):
    """توزيع حسب الموضع: الرأس / ختم الصادر (يسار فوق) / ختم الوارد (يمين تحت).

    ملاحظة مهمة: الرأس لازم يتحدّد بسطر «التحية» (تحية طيبة…)، لأن لو اعتمدنا
    على y1 < 30% من الصفحة بس، أسماء علم التوقيع في أواخر الورقة بتدخل
    وتغلب اسم الجهة في الترويسة. لو مفيش سطر تحية واضح، بنرجع للسلوك القديم.
    """
    header, top_left, bottom_right = [], [], []
    try:
        first_body = min(float(l.get("y0", 0) or 0) for l in lines_info
                         if re.search(r"تحية|التحية|وبعد|السلام عليكم|بسم الله",
                                      str(l.get("text", ""))))
    except ValueError:
        first_body = None
    header_limit = first_body if first_body else H * 0.30
    for l in lines_info:
        try:
            x0, x1 = float(l.get("x0", 0) or 0), float(l.get("x1", 0) or 0)
            y0, y1 = float(l.get("y0", 0) or 0), float(l.get("y1", 0) or 0)
        except (TypeError, ValueError):
            continue
        if y1 and y1 < header_limit:
            header.append(l)
            if x1 < W * 0.52:
                top_left.append(l)
        if y0 and y0 > H * 0.72 and x0 > W * 0.45:
            bottom_right.append(l)
    return {"header": header, "top_left": top_left, "bottom_right": bottom_right}


def region_digits(ctx, subset, W, H, tag):
    """قص المنطقة المشتركة لعدد من الأسطر + قراءتها بالنموذج الرقمي -> نصوص."""
    if not ctx or not ctx.get("prepped") or not subset:
        return []
    try:
        x0 = min(float(l.get("x0", 10 ** 9) or 10 ** 9) for l in subset)
        y0 = min(float(l.get("y0", 10 ** 9) or 10 ** 9) for l in subset)
        x1 = max(float(l.get("x1", 0) or 0) for l in subset)
        y1 = max(float(l.get("y1", 0) or 0) for l in subset)
    except (TypeError, ValueError):
        return []
    pad_x, pad_y = W * 0.03, H * 0.012
    workdir = Path(ctx.get("workdir") or TMP)
    return crop_digits_px(Path(ctx["prepped"]), x0 - pad_x, y0 - pad_y,
                          x1 + pad_x, y1 + pad_y, workdir, tag)


def _digits_parse(texts):
    """(أرقام، تواريخ) من نصوص النموذج الرقمي — الأسنان قد تكون ملزقة."""
    nums, dates = [], []
    for s in texts or []:
        s = str(s).translate(AR_DIGITS)
        d = find_date(s)
        if d and d.isoformat() not in dates:
            dates.append(d.isoformat())
        # سطر فيه «وارد/رقم/صادر» -> أرقامه أولاً
        labeled = bool(re.search(r"وارد|رقم|صادر|التاريخ", s))
        # الأرقام أحيانًا بتيجي مقطّعة (١ | ٢ | ٣) أو مقروءة حرف ورا حرف،
        # فبنجمع التلاتات المتجاورة في نفس السطر
        chunks = re.findall(r"\d{1,6}", s)
        for idx, tok in enumerate(chunks):
            v = tok.strip()
            if not v:
                continue
            if len(v) >= 6 and re.fullmatch(r"(19|20)\d{2}\d{0,4}", v):
                dd = find_date(v)
                if dd and dd.isoformat() not in dates:
                    dates.append(dd.isoformat())
                continue
            if re.fullmatch(r"(19|20|21)\d{2}", v):
                if labeled:
                    nums.append(v)   # سنة في سطر مرقّم: ممكن تكون سنة الإشارة
                continue
            nums.append(v)
            # تجميع رقم مقطّع: لو التلاتة الحالية قصيرة وجاية قريبة في نفس السطر
            merged = v
            for nxt in chunks[idx + 1:idx + 4]:
                if len(merged) >= 3 or len(nxt) > 3 or not nxt.isdigit():
                    break
                merged += nxt
            if len(merged) >= 3 and len(merged) <= 5 and merged not in nums:
                nums.append(merged)
    # ترتيب: الأرقام المرقّمة أولًا ثم الأقصر ثم الأطول
    return nums, dates


def _pick_header_sender(header, W, H):
    """الجهة المرسلة من الرأس، مع تفضيل النص العربي في يمين الصفحة."""
    org = (r'(شركة|هيئة|وزارة|مديرية|مديريه|إدارة|ادارة|الجهاز|مصلحة|محافظة|'
           r'مجلس|جامعة|نقابة|بنك|قطاع|Company|Authority|Ministry|Department)')
    skip = r'^(السيد|السادة|الأستاذ|الاستاذ|السيدة|تحية|Mr|Dear)'
    # شريط الشركة الأم في أعلى الترويسة مجرد لوجو، مش جهة المرسل.
    # نستبعده ما دام في ترشيح عربي أسفله (ترويسة الشركة الحقيقية).
    parent = re.compile(r'القابضة للمطارات|الملاحة الجوية|Egyptian Holding|'
                        r'الشركةة المصرية القابضة|الشركة المصرية القابضة',
                        re.IGNORECASE)
    cands = []
    for l in header:
        tx = str(l.get("text", "")).strip()
        if (not tx or len(tx) > 85 or re.match(skip, tx) or not re.search(org, tx)
                or re.search(r"إيماء|بشأن|بخصوص|في إطار|نحيط|يرجى|مرفق|وحيث|لذا", tx)):
            continue
        try:
            x1, y0 = float(l.get("x1", 0) or 0), float(l.get("y0", 0) or 0)
        except (TypeError, ValueError):
            x1, y0 = 0, 0
        right = x1 / max(W, 1)
        top = 1 - (y0 / max(H, 1))
        ar = _arabic_ratio(tx)
        score = ar * 4.0 + right * 3.0 + top * 3.0 + min(len(tx), 50) / 50.0
        if re.search(r"قطاع|إدارة|ادارة|مكتب|لشئون|لشؤون", tx) and not re.search(r"شركة|القابضة", tx):
            score += 4.0
        if ar < 0.4:
            score -= 6.0   # الإنجليزي آخر حل
        cands.append((score, tx[:120], ar, bool(parent.search(tx))))
    if not cands:
        return ""
    # لو في ترشيح غير شريط الشركة الأم -> نفضّله (ترويسة الشركة هي المرسل فعلًا)
    non_parent = [c for c in cands if not c[3]]
    pool = non_parent or cands
    pool.sort(key=lambda x: -x[0])
    chosen, chosen_ar = pool[0][1], pool[0][2]
    # لو اللي اتختر إنجليزي وفي عربي في نفس المنطقة -> نفضّل العربي
    if chosen_ar < 0.4:
        for _s, tx, ar, _par in pool:
            if ar >= 0.4:
                chosen = tx
                break
    # تصحيح أخطاء OCR في أسماء الجهات (مثال: «إنقاهرة» <- «القاهرة»)
    chosen = fix_entity(chosen)
    return chosen


def _plausible_stamp(num, crop_texts):
    """هل قراءة ختم الوارد مقنعة؟ خط اليد بيدي أرقام ملزقة وغريبة أحياناً."""
    if not num:
        return False
    n = str(num).strip()
    if not n.isdigit():
        return False
    if len(n) > 5:              # رقم وارد عادي = 1..5 أرقام
        return False
    joined = " ".join(crop_texts or [])
    # لو النص المقروء فيه نفس الرقم بالظبط -> القراءة متسقة
    if n in joined.replace(" ", ""):
        return True
    # الأرقام المطبوعة جوه المنطقة (تليفون/إيميل) بتلخبط القراءة -> نرفض
    if re.search(r"(?:\d{6,})", joined):
        return False
    return len(n) <= 4


def _handwritten_23_candidates(number):
    """بدائل قليلة لالتباس ٢/٣ في رقم الوارد المكتوب يدويًا.

    لا نغيّر الرقم تلقائيًا: رقم ٢ قد يكون صحيحًا. نضع أولًا البديل الذي
    يبدّل كل مواضع ٢ إلى ٣؛ لذلك 4252 ينتج 4353 مباشرة، ثم بدائل الموضع الواحد.
    """
    value = str(number or "").strip()
    if not value.isdigit() or "2" not in value:
        return []
    candidates = [value.replace("2", "3")]
    for index, digit in enumerate(value):
        if digit == "2":
            candidates.append(value[:index] + "3" + value[index + 1:])
    return list(dict.fromkeys(candidate for candidate in candidates
                              if candidate != value))


def _best_number(nums, date_hint=""):
    """أرجح رقم: مش سنة تاريخ، مش تليفون، مش طويل أوي."""
    year = (date_hint or "")[:4]
    good = []
    for v in nums or []:
        d = str(v).strip()
        if not d or not d.isdigit():
            continue
        if len(d) > 6:            # أرقام ملزقة أو تليفون
            continue
        if len(d) == 4 and d == year:
            continue              # سنة تاريخ الجواب
        if len(d) == 3:           # عداد صفحات/مرفقات غالباً
            continue
        good.append(d)
    return good[0] if good else ""


def _label_kind(text, phrases):
    n = _norm(str(text or ""))
    return next((kind for kind, variants in phrases.items()
                 if any(_norm(v) in n for v in variants)), "")


def _line_distance(a, b, W, H):
    """Distance between OCR boxes, scaled to the page rather than fixed stamp coordinates."""
    if a is b:
        return 0.0
    if a.get("_page", 1) != b.get("_page", 1):
        return None
    try:
        ax0, ax1 = float(a["x0"]), float(a["x1"])
        ay0, ay1 = float(a["y0"]), float(a["y1"])
        bx0, bx1 = float(b["x0"]), float(b["x1"])
        by0, by1 = float(b["y0"]), float(b["y1"])
    except (KeyError, TypeError, ValueError):
        return None
    ah, bh = max(ay1 - ay0, 1), max(by1 - by0, 1)
    acy, bcy = (ay0 + ay1) / 2, (by0 + by1) / 2
    y_gap = max(0, abs(acy - bcy) - (ah + bh) / 2)
    x_gap = max(0, bx0 - ax1, ax0 - bx1)
    y_limit = max((ah + bh) * 2.5, H * 0.075)
    x_limit = max((ax1 - ax0 + bx1 - bx0) * 1.5, W * 0.24)
    if y_gap > y_limit or x_gap > x_limit:
        return None
    return y_gap / max(y_limit, 1) + x_gap / max(x_limit, 1)


def _numbers_after_label(line, label_variants):
    text = str(line.get("text", "")).translate(AR_DIGITS)
    norm_text = _norm(text)
    starts = [norm_text.find(_norm(v)) for v in label_variants if _norm(v) in norm_text]
    compact_text = re.sub(r"\s+", "", norm_text)
    compact_labels = [re.sub(r"\s+", "", _norm(v)) for v in label_variants]
    compact_hit = any(label and label in compact_text for label in compact_labels)
    if starts:
        start = min(starts)
        tail = text[start:]
    elif compact_hit:
        tail = text
    else:
        tail = text
    values = []
    for value in re.findall(r"(?<!\d)\d{1,6}(?!\d)", tail):
        if len(value) == 4 and re.fullmatch(r"(?:19|20|21)\d{2}", value):
            continue
        if value not in values:
            values.append(value)
    return values


def _contextual_number(lines_info, anchors, W, H, label_variants):
    options = []
    for anchor in anchors:
        same_line = _numbers_after_label(anchor, label_variants)
        options.extend((0.0, value, "high") for value in same_line)
        for line in lines_info:
            if line is anchor:
                continue
            distance = _line_distance(anchor, line, W, H)
            if distance is None:
                continue
            for value in _numbers_after_label(line, ()):
                options.append((distance + 0.15, value, "medium"))
    options.sort(key=lambda item: item[0])
    if not options:
        return "", [], "low"
    best_score = options[0][0]
    best_values = list(dict.fromkeys(value for score, value, _level in options
                                     if score <= best_score + 0.08))
    if len(best_values) != 1:
        return best_values[0], best_values, "low"
    level = options[0][2]
    return best_values[0], best_values, level


def _incoming_marker_anchor(lines_info, W, H):
    """Recover an incoming number from its marker, number row and adjacent date.

    These three independent layout cues prevent a damaged printed label from
    sending number extraction to an unrelated outgoing stamp.
    """
    candidates = []
    for marker in lines_info:
        compact = re.sub(r'\s+', '', _norm(marker.get('text', '')))
        if len(compact) > 22 or not re.search(r'وارد|والد|و[١٢12]لد', compact):
            continue
        if re.search(r'صادر|عامل|هاتف|مرفق', compact):
            continue
        for line in lines_info:
            if _line_distance(marker, line, W, H) is None:
                continue
            raw = str(line.get('text', '')).translate(AR_DIGITS)
            if re.search(r'@|mail|tel|fax|عامل', raw, re.I):
                continue
            height = float(line['y1']) - float(line['y0'])
            marker_height = float(marker['y1']) - float(marker['y0'])
            dy = abs((float(line['y0'])+float(line['y1'])-float(marker['y0'])-float(marker['y1']))/2)
            if dy > max(H*.02, (height+marker_height)/2):
                continue
            values = []
            for match in re.finditer(r'(?<!\d)\d{1,6}(?!\d)', raw):
                value = match.group()
                if (match.start() and raw[match.start()-1].isalpha()) or (match.end()<len(raw) and raw[match.end()].isalpha()):
                    continue
                if re.fullmatch(r'(?:19|20|21)\d{2}',value):
                    continue
                values.append(value)
            if len(set(values)) != 1:
                continue
            date_below = any(
                other.get('_page',1)==line.get('_page',1)
                and 0 < float(other['y0'])-float(line['y1']) < H*.12
                and _line_distance(line,other,W,H) is not None
                and (find_date(str(other.get('text',''))) or
                     re.search(r'(?:20|21)\d{4,8}',str(other.get('text','')).translate(AR_DIGITS)))
                for other in lines_info)
            if date_below:
                candidates.append({**line, '_incoming_marker_number':values[0]})
    numbers = {line['_incoming_marker_number'] for line in candidates}
    if len(numbers) == 1:
        return min(candidates,key=lambda line:float(line['y1'])-float(line['y0']))
    return None


def _read_incoming_number(prepped, anchor, lines_info, W, H):
    """Read the complete handwritten value, including strokes lost by detection."""
    if prepped is None or anchor is None:
        return "", [], []
    from PIL import Image
    import numpy as np
    candidates = []
    ax = (float(anchor['x0']) + float(anchor['x1'])) / 2
    ay = (float(anchor['y0']) + float(anchor['y1'])) / 2
    for line in lines_info:
        if line.get('_page', 1) != anchor.get('_page', 1):
            continue
        raw = str(line.get('text', '')).translate(AR_DIGITS).strip()
        digits = re.findall(r'\d', raw)
        if not 2 <= len(digits) <= 6 or re.search(r'[/@]|mail|tel|fax|عامل', raw, re.I):
            continue
        if re.fullmatch(r'(?:19|20|21)\d{2}', raw):
            continue
        if len(digits) / max(len(raw.replace(' ', '')), 1) < .55:
            continue
        cy = (float(line['y0']) + float(line['y1'])) / 2
        cx = (float(line['x0']) + float(line['x1'])) / 2
        if not -.012 * H <= cy - ay <= .045 * H or abs(cx - ax) > .10 * W:
            continue
        candidates.append((abs(cy - ay) / H + abs(cx - ax) / W, line))
    if not candidates:
        return "", [], []
    line = min(candidates, key=lambda item: item[0])[1]
    with Image.open(prepped) as source:
        image = source.convert('L')
    votes, evidence = {}, []
    for pad_x in (.003, .006, .008):
        for pad_y in (0, .003):
            box = (max(0, int(line['x0'] - W * pad_x)),
                   max(0, int(line['y0'] - H * pad_y)),
                   min(image.width, int(line['x1'] + W * pad_x)),
                   min(image.height, int(line['y1'] + H * pad_y)))
            result = _run_engine(ocr_engine(), np.asarray(image.crop(box)),
                                 use_det=False, use_cls=False)
            for reading in _res_lines(result):
                raw = re.sub(r'\s+', '', reading['text'].translate(AR_DIGITS))
                if not re.fullmatch(r'\d{2,6}', raw) or reading.get('ocr_score', 0) < .65:
                    continue
                votes[raw] = votes.get(raw, 0) + 1
                evidence.append({**line, 'text': reading['text'],
                                 'ocr_score': reading.get('ocr_score'),
                                 '_number_crop_confirmed': True})
    ranked = sorted(votes, key=votes.get, reverse=True)
    if ranked and votes[ranked[0]] >= 3 and (len(ranked) == 1
            or votes[ranked[0]] >= 2 * votes[ranked[1]]):
        value = ranked[0]
        return value, ranked, [item for item in evidence
                               if item['text'].translate(AR_DIGITS).strip() == value]
    return "", ranked, []


def _contextual_date(lines_info, anchors, W, H, *, explicit=(), require_anchor=False,
                     incoming=False):
    options = []
    explicit_norm = tuple(_norm(v) for v in explicit)
    for line in lines_info:
        text = str(line.get("text", ""))
        normalized = _norm(text)
        if anchors and not any(a.get("_page", 1) == line.get("_page", 1) for a in anchors):
            continue
        found = find_date(text)
        if not found or not (_valid_stamp_date(found) if incoming else _fresh(found)):
            continue
        explicitly_labeled = any(term in normalized for term in explicit_norm)
        if explicitly_labeled:
            options.append((0.0, found.isoformat(), "high"))
            continue
        compact = re.sub(r"\s+", "", normalized)
        if not any(re.sub(r"\s+", "", term) in compact
                   for term in ("التاريخ", "تاريخ", "تاربخ", "تاريح", "الثاري")):
            continue
        distances = [_line_distance(anchor, line, W, H) for anchor in anchors]
        distances = [d for d in distances if d is not None]
        if distances:
            options.append((min(distances) + 0.05, found.isoformat(), "medium"))
        elif not require_anchor:
            try:
                y0 = float(line.get("y0", H))
            except (TypeError, ValueError):
                continue
            if y0 <= H * 0.32:
                options.append((0.4 + y0 / max(H, 1), found.isoformat(), "medium"))
    options.sort(key=lambda item: item[0])
    if not options:
        return "", "low"
    best = options[0][0]
    dates = list(dict.fromkeys(value for score, value, _level in options
                               if score <= best + 0.08))
    if len(dates) != 1:
        return dates[0], "low"
    return dates[0], options[0][2]


def _field_evidence(fields, lines_info):
    """يربط كل اقتراح بأقرب سطر OCR، موضعه، ودرجة التعرف إن وفرها المحرك."""
    labels = {
        "number": ("رقم الوارد", "رقم وارد", "الوارد رقم", "رقم الورد"),
        "wared_date": ("تاريخ الوارد", "التاريخ", "الثاري"),
        "export_number": ("رقم الصادر", "رقم الاشارة", "رقم الإشارة"),
        "export_date": ("تاريخ الصادر", "تاريخ الخطاب", "تاريخ الكتاب",
                        "الثاري", "تاريح"),
        "attachments": ("عدد المرفقات", "المرفقات"),
        "sender": ("شركة", "قطاع", "إدارة", "هيئة", "وزارة"),
        "recipient": ("السيد", "السادة", "رئيس", "مدير", "إلى"),
        "jeha": ("رئيس قطاع", "قطاع", "إدارة", "الادارة"),
        "subject": ("الموضوع", "subject"),
    }
    numeric = {"number", "export_number", "attachments"}
    dates = {"wared_date", "export_date"}
    evidence = {}
    confidence = fields.get("_confidence", {})
    for key in ("number", "wared_date", "export_number",
                "export_date", "attachments", "sender", "recipient", "jeha",
                "subject"):
        value = str(fields.get(key) or "").strip()
        value_norm = _norm(value)
        matches = []
        if value:
            for line in lines_info or []:
                phrase = str(line.get("text", "")).strip()
                normalized = _norm(phrase)
                match_strength = 0
                if key in dates:
                    found = _date_from_fragment(phrase)
                    match_strength = 3 if found and found.isoformat() == value else 0
                elif key in numeric:
                    normalized_digits = phrase.translate(AR_DIGITS)
                    if re.search(rf"(?<!\d){re.escape(value)}(?!\d)", normalized_digits):
                        match_strength = 3
                    else:
                        # القيمة المنظفة (مثل 1590 من ١٥٩٠٠ بعد حذف الصفر الزائد)
                        # جزء من سطر OCR — دليل صالح لكن بدرجة عنوان (1) لا تحقق كامل.
                        compact_line = re.sub(r"\s+", "", normalized_digits)
                        compact_val = re.sub(r"\s+", "", value)
                        if compact_val and compact_val in compact_line:
                            match_strength = 1
                elif value_norm and value_norm in normalized:
                    match_strength = 3
                elif key == "subject":
                    words = [word for word in value_norm.split() if len(word) > 2]
                    overlap = sum(1 for word in words if word in normalized)
                    if overlap >= (1 if len(words) == 1 else 2):
                        match_strength = overlap
                if match_strength:
                    matches.append((match_strength, line))
        if not matches:
            markers = tuple(_norm(label) for label in labels.get(key, ()))
            for line in lines_info or []:
                phrase = str(line.get("text", "")).strip()
                normalized = _norm(phrase)
                if any(marker and marker in normalized for marker in markers):
                    matches.append((1, line))
        matches.sort(key=lambda item: (-item[0], item[1].get("_page", 1),
                                       item[1].get("y0", 0)))
        sources = []
        verified = any(strength >= 3 for strength, _line in matches)
        for strength, line in matches[:2]:
            source = {"phrase": str(line.get("text", "")).strip()}
            source["matches_value"] = strength >= 3
            if "_page" in line:
                source["page"] = line["_page"]
            if all(k in line for k in ("x0", "y0", "x1", "y1")):
                source["position"] = {
                    "bbox": [line["x0"], line["y0"], line["x1"], line["y1"]],
                    "page_size": line.get("_page_size"),
                }
            if isinstance(line.get("ocr_score"), (int, float)):
                source["ocr_score"] = line["ocr_score"]
            sources.append(source)
        evidence[key] = {
            "value": value,
            "confidence": confidence.get(key, "low") if value else "low",
            # المقصود: هل ظهر المقترح نفسه في سطر OCR؟ وجود عنوان الحقل وحده
            # ليس دليلاً كافيًا لملء قيمة تلقائيًا.
            "verified": verified,
            "sources": sources,
        }
    return evidence


# أرقام بالحروف بتظهر في خانات الختم المكتوبة بخط اليد، و«واصر» خطأ OCR
# متكرر لخط اليد «صفر» في 값 خانة عدد المرفقات.
AR_NUMBER_WORDS = (("صفر", "0"), ("واصر", "0"),
                   ("واحد", "1"), ("اثنان", "2"), ("اثنين", "2"),
                   ("ثلاث", "3"), ("اربع", "4"), ("خمس", "5"),
                   ("ست", "6"), ("سبع", "7"), ("ثمان", "8"),
                   ("تسع", "9"), ("عشره", "10"))
# عناوين الخانات المجاورة: لو ظهرت جوا سطر القيمة يبقى السطر ده مش قيمة.
_ATTACHMENT_BLOCKERS = ("وارد", "صادر", "مرفقات", "تاريخ", "ملف",
                        "توقيع", "تحيه", "ref")


def _attachment_cell_value(label, lines_info, W, H):
    """قيمة خانة «عدد المرفقات» اللي في نفس عمود العنوان وتحته (جدول الختم).

    خانات الختم بترتب «العنوان» فوق «القيمة»، والقيمة بتكون احياناً
    بالحروف («صفر») مش ارقام. بنرجعها رقم نصي زي باقي الحقول، وبنشترط
    ان القراءة واحدة غير مبهمة عشان متقدرش من سطر واحد.
    """
    lx0, lx1 = float(label.get("x0", 0)), float(label.get("x1", 0))
    ly1 = float(label.get("y1", 0))
    found = set()
    for line in lines_info:
        if line is label:
            continue
        text = str(line.get("text", "")).strip()
        if not text:
            continue
        normalized = _norm(text)
        if any(word in normalized for word in _ATTACHMENT_BLOCKERS):
            continue
        mx0, mx1 = float(line.get("x0", 0)), float(line.get("x1", 0))
        my0 = float(line.get("y0", 0))
        if mx0 >= lx1 + W * 0.05 or mx1 <= lx0 - W * 0.05:
            continue                    # عمود تاني
        if my0 < ly1 - H * 0.01 or my0 > ly1 + H * 0.06:
            continue                    # مش في الصف اللي تحتي مباشرة
        digits = "".join(c for c in text.translate(AR_DIGITS) if c.isdigit())
        if digits and len(digits) <= 2 and not re.search(r"\d{3,}",
                                                         text.translate(AR_DIGITS)):
            found.add(digits)
            continue
        for word, value in AR_NUMBER_WORDS:
            if word in normalized:
                found.add(value)
                break
    return next(iter(found)) if len(found) == 1 else ""


# عروض شريط تاريخ الوارد: (اتساع يسار، فوق، تحت) كنسبة من H.
# عرض مختلف بيغيّر آخر رقم في القراءة، فبنصوّت على النتيجة بدل ما ناخد أول قراءة.
_DATE_BAND_SHAPES = ((0.14, -0.005, 0.075), (0.16, -0.005, 0.075),
                     (0.18, -0.005, 0.075), (0.20, -0.005, 0.075),
                     (0.22, -0.005, 0.075))


def _date_band_boxes(num_line, lines_info, W, H):
    """مناطق مرشّقة لتاريخ الوارد: عروض الشريط تحت «رقم الوارد» + سطر التاريخ."""
    nx0, nx1 = float(num_line.get("x0", 0)), float(num_line.get("x1", W))
    ny1 = float(num_line.get("y1", H))
    boxes = [(nx0 - W * left, ny1 + H * top, nx1 + W * 0.12, ny1 + H * bottom)
             for left, top, bottom in _DATE_BAND_SHAPES]
    best = None
    for line in lines_info:
        if line is num_line:
            continue
        if re.search(r"@|e[.-]?mail|\btel\b|\bfax\b", str(line.get("text", "")), re.I):
            continue
        ly0, ly1 = float(line.get("y0", 0)), float(line.get("y1", 0))
        lx0, lx1 = float(line.get("x0", 0)), float(line.get("x1", 0))
        if ly0 < ny1 - H * 0.01 or ly0 > ny1 + H * 0.08:
            continue
        if lx0 > nx1 + W * 0.25 or lx1 < nx0 - W * 0.25:
            continue
        count = len(re.findall(r"\d", str(line.get("text", "")).translate(AR_DIGITS)))
        if count >= 2 and (best is None or count > best[0]):
            best = (count, (lx0 - W * 0.02, ly0 - H * 0.012,
                            lx1 + W * 0.03, ly1 + H * 0.02))
    if best:
        boxes.append(best[1])
    return boxes


def _date_without_slashes(crop):
    """Remove two isolated diagonal separators, keeping handwritten digits intact."""
    import cv2
    import numpy as np
    pixels = np.asarray(crop).copy()
    mask = (pixels < 200).astype("uint8") * 255
    _, components, stats, _ = cv2.connectedComponentsWithStats(mask)
    separators = []
    for index, (x, y, width, height, area) in enumerate(stats[1:], 1):
        if (height < 15 or not 0.2 < width / height < 0.75
                or height > pixels.shape[0] * 0.7 or area / (width * height) > 0.3):
            continue
        ys, xs = np.where(components == index)
        if abs(float(np.corrcoef(xs, ys)[0, 1])) > 0.75:
            separators.append((int(x), int(y), int(width), int(height)))
    if len(separators) != 2:
        return None
    for x, y, width, height in separators:
        pixels[max(0, y - 2):y + height + 2, max(0, x - 2):x + width + 2] = 255
    return pixels


def _ink_confirms_day11(crop):
    """Two date separators followed by two tall strokes distinguish 11 from 10."""
    import cv2
    import numpy as np
    pixels = np.asarray(crop.convert('L'))
    _, components, stats, _ = cv2.connectedComponentsWithStats((pixels < 200).astype('uint8'))
    slashes = []
    for i, (x,y,w,h,area) in enumerate(stats[1:], 1):
        if h < 15 or not .2 < w/h < .75 or area/(w*h) > .35 or h > pixels.shape[0]*.8:
            continue
        ys,xs = np.where(components == i)
        if float(np.corrcoef(xs,ys)[0,1]) < -.75:
            slashes.append((x,y,w,h))
    groups = []
    for slash in sorted(slashes):
        if groups and slash[0]-groups[-1][-1][0] < max(s[3] for s in slashes)*.9:
            groups[-1].append(slash)
        else:
            groups.append([slash])
    if len(groups) != 2:
        return False
    edge = max(x+w for x,y,w,h in slashes)
    strokes = [(x,y,w,h) for x,y,w,h,area in stats[1:]
               if x > edge and h > max(s[3] for s in slashes)*1.5
               and h < pixels.shape[0]*.9 and .12 < w/h < .5]
    return len(strokes) == 2 and abs(strokes[0][3]-strokes[1][3]) < max(s[3] for s in strokes)*.25


def _table_stamp_date(prepped, num_line, W, H):
    """Locate a ruled stamp containing its incoming label; never use a page region alone."""
    import cv2
    import numpy as np
    from PIL import Image
    im = Image.open(prepped).convert("L")
    ink = (np.asarray(im) < 210).astype("uint8") * 255
    horizontal = cv2.morphologyEx(ink, cv2.MORPH_OPEN,
                                  np.ones((1, max(30, int(W // 20))), np.uint8))
    vertical = cv2.morphologyEx(ink, cv2.MORPH_OPEN,
                                np.ones((max(30, int(H // 40)), 1), np.uint8))
    rules = cv2.morphologyEx(horizontal | vertical, cv2.MORPH_CLOSE,
                             np.ones((7, 7), np.uint8))
    contours, _ = cv2.findContours(rules, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    cx = (float(num_line['x0']) + float(num_line['x1'])) / 2
    cy = (float(num_line['y0']) + float(num_line['y1'])) / 2
    engine = ocr_engine()
    votes, readings = {}, []
    rectangles = []
    for contour in contours:
        x, y, width, height = cv2.boundingRect(contour)
        if not (.12 * W < width < .6 * W and .035 * H < height < .3 * H
                and x <= cx <= x + width and y <= cy <= y + height):
            continue
        if any(abs(x-r[0]) < 15 and abs(y-r[1]) < 15 for r in rectangles):
            continue
        rectangles.append((x, y, width, height))
        for top in (.72, .75, .8):
            for right in (.77, .85, .95):
                crop = im.crop((x + 8, int(y + height * top),
                                int(x + width * right), int(y + height * 1.12)))
                result = _run_engine(engine, np.asarray(crop), use_det=False, use_cls=False)
                for phrase in getattr(result, 'txts', None) or []:
                    raw = str(phrase).translate(AR_DIGITS)
                    if not re.fullmatch(r'(?:20|21)\d{5,6}', raw.strip()):
                        continue
                    got = _date_from_fragment(raw)
                    if got and _valid_stamp_date(got):
                        value = got.isoformat()
                        votes[value] = votes.get(value, 0) + 1
                        readings.append({**num_line, 'text': raw,
                                         '_date_crop_confirmed': True})
    if len(votes) == 1 and next(iter(votes.values())) >= 3:
        return next(iter(votes)), readings
    return '', []


def _read_stamp_date_line(prepped, num_line, lines_info, W, H, *, allow_extended=False):
    """Read a located date value directly when detection splits the date label."""
    from PIL import Image
    import numpy as np
    labels = [line for line in lines_info if "تاريخ" in _norm(line.get("text", ""))
              and _line_distance(num_line, line, W, H) is not None]
    pairs = []
    for label in labels:
        if re.search(r"(?:20|21)\d{2}", str(label.get("text", "")).translate(AR_DIGITS)):
            pairs.append((0, label, label))
        for line in lines_info:
            tx = str(line.get("text", "")).translate(AR_DIGITS)
            if line is label or not re.search(r"\d{5,8}", tx) or re.search(r"@|mail|tel|fax", tx, re.I):
                continue
            if line.get("_page", 1) != label.get("_page", 1):
                continue
            cy = (float(line["y0"]) + float(line["y1"])) / 2
            label_cy = (float(label["y0"]) + float(label["y1"])) / 2
            if abs(cy - label_cy) <= H * 0.025 and _line_distance(label, line, W, H) is not None:
                pairs.append((abs(cy - label_cy), label, line))
    if not pairs:
        # A merged handwritten date may lose both its year and printed label
        # in detection. Re-read that actual line directly, below the stamp
        # number, rather than accepting dates from a broad crop.
        for line in lines_info:
            tx = str(line.get("text", "")).translate(AR_DIGITS)
            if line.get("_page", 1) != num_line.get("_page", 1):
                continue
            if len(re.findall(r"\d", tx)) < 3 or re.search(r"@|mail|tel|fax", tx, re.I):
                continue
            dy = float(line["y0"]) - float(num_line["y1"])
            if 0 < dy < H * 0.10 and _line_distance(num_line, line, W, H) is not None:
                digits_count = len(re.findall(r"\d", tx))
                rank = (-2 if re.search(r"(?:20|21)\d{2}", tx) else
                        -1 - digits_count / 100 if digits_count >= 5 else dy)
                pairs.append((rank, line, line))
    if not pairs:
        return _table_stamp_date(prepped, num_line, W, H) if allow_extended else ("", [])
    _, label, line = min(pairs, key=lambda item: item[0])
    im = Image.open(prepped).convert("L")
    engine = ocr_engine()
    original_digits = re.sub(r'\s+', '', str(line.get('text','')).translate(AR_DIGITS))
    if re.fullmatch(r'(?:20|21)\d{7,8}', original_digits):
        # A complete date cannot have nine or ten digits. Read the whole line
        # again only after two actual diagonal separators have been isolated.
        ink_votes, ink_readings = {}, []
        for pad_y in (.006,.008,.010):
            crop = im.crop((max(0,int(line['x0']-W*.005)), max(0,int(line['y0']-H*pad_y)),
                min(im.width,int(line['x1']+W*.005)), min(im.height,int(line['y1']+H*pad_y))))
            clean = _date_without_slashes(crop)
            if clean is None:
                continue
            result = _run_engine(engine, clean, use_det=False, use_cls=False)
            for phrase in getattr(result,'txts',None) or []:
                raw = str(phrase).translate(AR_DIGITS).strip()
                reversed_chunks = re.fullmatch(r'(\d{2,4})\s+((?:20|21)\d{2})',raw)
                if reversed_chunks:
                    raw = reversed_chunks[2] + reversed_chunks[1]
                if not re.fullmatch(r'(?:20|21)\d{5,6}',raw):
                    continue
                got = _date_from_fragment(raw)
                if got and _valid_stamp_date(got):
                    value = got.isoformat(); ink_votes[value] = ink_votes.get(value,0)+1
                    ink_readings.append({**line,'text':raw,'_date_crop_confirmed':True,'_date_separators_confirmed':True})
        if len(ink_votes)==1 and next(iter(ink_votes.values()))>=3:
            return next(iter(ink_votes)),ink_readings
    votes, readings, separated_votes = {}, [], {}
    # Asymmetric padding keeps the full handwriting without the lower ruling.
    tight_votes, tight_readings = {}, []
    for left_pad in ((.005, .0075, .01) if allow_extended else ()):
        for top_pad in (.002, .003, .004, .005):
            for right_pad in (.025, .03, .04):
                # Round outward to small pixel grids: fractional OCR boxes
                # must not shave an endpoint from a handwritten stroke.
                box = (max(0, int((float(line['x0']) - W * left_pad) // 10) * 10),
                       max(0, int((float(line['y0']) - H * top_pad) // 5) * 5),
                       int((float(line['x1']) - W * right_pad) // 10) * 10,
                       min(im.height, int(math.ceil(
                           (float(line['y1']) + H * .001) / 5)) * 5))
                if box[2] <= box[0]:
                    continue
                result = _run_engine(engine, np.asarray(im.crop(box)), use_det=False, use_cls=False)
                for phrase in getattr(result, 'txts', None) or []:
                    raw = str(phrase).translate(AR_DIGITS)
                    match = re.search(r'(?<!\d)((?:20|21)\d{5,6})(?!\d)', raw)
                    got = _date_from_fragment(match[1]) if match else None
                    if got and _valid_stamp_date(got):
                        value = got.isoformat()
                        tight_votes[value] = tight_votes.get(value, 0) + 1
                        tight_readings.append({**line, 'text': raw, '_date_crop_confirmed': True})
    # Printed table borders and date separators can become spurious digits.
    # Only use this independent reading when two actual slash strokes are found.
    cleaned_votes, cleaned_readings = {}, []
    for right_pad in (0.03, 0.035, 0.04):
        box = (max(0, int(float(line["x0"]) - W * 0.04)),
               max(0, int(float(line["y0"]) - H * 0.006)),
               min(im.width, int(float(line["x1"]) - W * right_pad)),
               min(im.height, int(float(line["y1"]) + H * 0.006)))
        if box[2] <= box[0] or box[3] <= box[1]:
            continue
        cleaned = _date_without_slashes(im.crop(box))
        if cleaned is None:
            continue
        for angle in (-1, 0, 1):
            image = Image.fromarray(cleaned).rotate(angle, expand=True, fillcolor=255)
            result = _run_engine(engine, np.asarray(image), use_det=False, use_cls=False)
            for phrase in getattr(result, "txts", None) or []:
                raw = str(phrase).translate(AR_DIGITS).strip()
                if not re.fullmatch(r"(?:20|21)\d{4,6}", raw):
                    continue
                got = _date_from_fragment(raw)
                if got and _valid_stamp_date(got):
                    value = got.isoformat()
                    cleaned_votes[value] = cleaned_votes.get(value, 0) + 1
                    cleaned_readings.append({**line, "text": raw, "_date_crop_confirmed": True})
    if len(cleaned_votes) == 1 and next(iter(cleaned_votes.values())) >= 3:
        value = next(iter(cleaned_votes))
        parsed = _d(value)
        if parsed.day == 10:
            full = im.crop((max(0,int(line['x0']-W*.04)), max(0,int(line['y0']-H*.006)),
                            min(im.width,int(line['x1'])), min(im.height,int(line['y1']+H*.006))))
            if _ink_confirms_day11(full):
                value = parsed.replace(day=11).isoformat()
                cleaned_readings.append({**line,'text':value,'_date_crop_confirmed':True,'_day_ink_confirmed':True})
        return value, cleaned_readings
    crop_options = [(pad_y, 0.015, -0.015) for pad_y in (0.001, 0.002, 0.003, 0.004, 0.006)]
    if label is line:
        crop_options += [(pad_y, left_pad, right_pad)
                        for pad_y in (0.001, 0.002, 0.003, 0.004, 0.006)
                        for left_pad in (0.04, 0.05, 0.06)
                        for right_pad in (0.025, 0.035, 0.04)]
        crop_options += [(pad_y, 0.005, right_pad)
                         for pad_y in (0.001, 0.002, 0.003)
                         for right_pad in (0.03, 0.04)]
    for pad_y, left_pad, right_pad in crop_options:
        left = max(0, float(line["x0"]) - W * left_pad)
        right = min(im.width, float(line["x1"]) - W * right_pad)
        if float(label["x0"]) > float(line["x0"]):
            right = min(right, float(label["x0"]))
        box = (int(left), max(0, int(float(line["y0"]) - H * pad_y)),
               int(right), min(im.height, int(float(line["y1"]) + H * pad_y)))
        if box[2] - box[0] < 20 or box[3] - box[1] < 10:
            continue
        phrases = []
        for angle in ((-1, 0, 1) if left_pad == 0.005 else (0,)):
            try:
                crop = im.crop(box)
                if angle:
                    crop = crop.rotate(angle, expand=True, fillcolor=255)
                result = _run_engine(engine, np.asarray(crop), use_det=False, use_cls=False)
                phrases.extend(getattr(result, "txts", None) or [])
            except Exception:
                continue
        for tx in phrases:
            # Six compact digits can be a clipped day 11 mistaken for 1.
            # Accept a complete separated date or at least seven digits here.
            raw_digits = str(tx).translate(AR_DIGITS)
            separated = re.search(r"(?:20|21)\d{2}\d{1,2}[/٬،,]\d{1,2}", raw_digits)
            if len(re.findall(r"\d", raw_digits)) < 7 and not separated and not re.search(
                    r"\d{4}\s*[/.-]\s*\d{1,2}\s*[/.-]\s*\d{1,2}", raw_digits):
                continue
            got = _date_from_fragment(tx)
            if got and _valid_stamp_date(got):
                value = got.isoformat()
                votes[value] = votes.get(value, 0) + 1
                if separated:
                    separated_votes[value] = separated_votes.get(value, 0) + 1
                readings.append({**line, "text": str(tx), "_date_crop_confirmed": True})
    if len(separated_votes) == 1 and next(iter(separated_votes.values())) >= 2:
        return next(iter(separated_votes)), [line for line in readings
               if _date_from_fragment(line["text"]).isoformat() == next(iter(separated_votes))]
    if len(votes) == 1 and next(iter(votes.values())) >= 2:
        return next(iter(votes)), readings
    if not allow_extended:
        return "", []
    # Preserve the whole handwritten line. A tight right crop can shave off the
    # day; a little vertical breathing room makes slashes less likely to become 1.
    full_votes, full_readings = {}, []
    for left_pad in (0, .005):
        for pad_y in (.006, .008, .010):
            for right_pad in (0, .005):
                crop = im.crop((max(0,int(line['x0']-W*left_pad)),
                    max(0,int(line['y0']-H*pad_y)),
                    min(im.width,int(line['x1']+W*right_pad)),
                    min(im.height,int(line['y1']+H*pad_y))))
                result = _run_engine(engine,np.asarray(crop),use_det=False,use_cls=False)
                for phrase in getattr(result,'txts',None) or []:
                    raw = str(phrase).translate(AR_DIGITS).strip()
                    if not re.fullmatch(r'(?:20|21)\d{5,6}',raw):
                        continue
                    got = _date_from_fragment(raw)
                    if got and _valid_stamp_date(got):
                        value = got.isoformat()
                        full_votes[value] = full_votes.get(value,0)+1
                        full_readings.append({**line,'text':str(phrase),'_date_crop_confirmed':True})
    ranked = sorted(full_votes.items(),key=lambda item:item[1],reverse=True)
    if ranked and ranked[0][1] >= 4 and (len(ranked)==1 or ranked[0][1] >= 2*ranked[1][1]):
        return ranked[0][0], [reading for reading in full_readings
            if _date_from_fragment(reading['text']).isoformat()==ranked[0][0]]
    table_date, table_readings = _table_stamp_date(prepped, num_line, W, H)
    if table_date:
        return table_date, table_readings
    if len(tight_votes) == 1 and next(iter(tight_votes.values())) >= 3:
        return next(iter(tight_votes)), tight_readings
    return "", []


def _stamp_date_line(prepped, num_line, lines_info, W, H, *, allow_extended=False):
    value, evidence = _read_stamp_date_line(prepped,num_line,lines_info,W,H,
                                           allow_extended=allow_extended)
    if value and _d(value).day == 10:
        from PIL import Image
        with Image.open(prepped) as source:
            image=source.convert('L')
        for line in lines_info:
            if line.get('_page',1) != num_line.get('_page',1):
                continue
            dy=float(line['y0'])-float(num_line['y1'])
            if not (0 < dy < H*.1 and _line_distance(num_line,line,W,H) is not None):
                continue
            if len(re.findall(r'\d',line.get('text','').translate(AR_DIGITS))) < 5:
                continue
            crop=image.crop((max(0,int(line['x0']-W*.04)),max(0,int(line['y0']-H*.006)),
                             min(image.width,int(line['x1']+W*.01)),min(image.height,int(line['y1']+H*.006))))
            if _ink_confirms_day11(crop):
                value=_d(value).replace(day=11).isoformat()
                evidence.append({**line,'text':value,'_date_crop_confirmed':True,'_day_ink_confirmed':True})
                break
    return value,evidence


def _read_located_stamp_date(prepped, geometry):
    """Read the bottom date cell of a text-confirmed incoming table."""
    if not prepped or not geometry.get('confirmed'):
        return '', []
    from PIL import Image, ImageOps
    import numpy as np
    with Image.open(prepped) as source:
        image = source.convert('L')
    x0, y0, x1, y1 = geometry['box']
    width, height = x1 - x0, y1 - y0
    cell = geometry.get('date_box')
    # A located table supplies bounds. Nothing is inferred from the page corner
    # or from the decision's independently printed date.
    votes, evidence = {}, []
    angle = math.degrees(math.atan(geometry.get('slope', 0)))
    for top in ((0,) if cell else (.70, .73, .76)):
        for right in (.76, .80, .84):
            bx0 = max(0, round((cell[0] if cell else x0) - width * .025))
            by0 = max(0, round(cell[1] if cell else y0 + height * top))
            bx1 = min(image.width, round(x0 + width * right))
            by1 = min(image.height, round((cell[3] if cell else y1) + height * .015))
            if bx1 <= bx0 or by1 <= by0:
                continue
            crop = image.crop((bx0, by0, bx1, by1))
            if abs(angle) > .4:
                crop = crop.rotate(angle, expand=True, fillcolor=255)
            for variant in (crop, ImageOps.autocontrast(crop, cutoff=1)):
                result = _run_engine(ocr_engine(), np.asarray(variant),
                                     use_det=False, use_cls=False)
                for reading in _res_lines(result):
                    raw = re.sub(r'\s+', '', reading['text'].translate(AR_DIGITS))
                    if not re.fullmatch(r'(?:20|21)\d{5,6}', raw):
                        continue
                    found = _date_from_fragment(raw)
                    if not _valid_stamp_date(found):
                        continue
                    value = found.isoformat()
                    votes[value] = votes.get(value, 0) + 1
                    evidence.append({'text': reading['text'], 'x0': bx0, 'x1': bx1,
                        'y0': by0, 'y1': by1, 'ocr_score': reading.get('ocr_score'),
                        '_page': geometry.get('_page', 1),
                        '_page_size': [image.width, image.height],
                        '_date_crop_confirmed': True, '_stamp_cell_read': True})
    ranked = sorted(votes, key=votes.get, reverse=True)
    if ranked and votes[ranked[0]] >= 3 and (len(ranked) == 1
            or votes[ranked[0]] >= 2 * votes[ranked[1]]):
        value = ranked[0]
        return value, [line for line in evidence
                       if _date_from_fragment(line['text']).isoformat() == value]
    return '', []


def _wared_date_from_bands(prepped, num_line, lines_info, W, H, workdir):
    """قراءة أخيرة لتاريخ الوارد من قص شريط الختم (نص عادي + Otsu).

    الـdetector أحياناً بيقفل على التاريخ وبيرجّع جزء منه بس، وفيه فرق
    دقيقة في عرض القص بتغيّر آخر رقم. فبنقص بعدة عروض ونتفق: لو غالب
    القراءات على تاريخ واحد بنقبله بمستوى ثقة منخفض (الحقل بيتعرض
    للمراجعة)، ولو القراءات متضاربة فبنتركه فاضي — أحسن ما نسجّل تاريخ
    غلط في السجل الرسمي.
    """
    if prepped is None or num_line is None:
        return ""
    value, readings = _stamp_date_line(prepped, num_line, lines_info, W, H,
                                     allow_extended=True)
    if value:
        lines_info.extend(readings)
        return value
    votes = {}
    for index, box in enumerate(_date_band_boxes(num_line, lines_info, W, H)):
        for variant, scale in (("", 3), ("otsu", 3), ("otsu", 4)):
            tag = f"wared_{index}_{variant or 'raw'}_{scale}"
            for text in crop_date_band_px(prepped, *box, workdir, tag,
                                          scale=scale, binarize=bool(variant)):
                numeric = str(text).translate(AR_DIGITS)
                if len(re.findall(r"\d", numeric)) < 7 and not re.search(
                        r"\d{4}\s*[/.-]\s*\d{1,2}\s*[/.-]\s*\d{1,2}", numeric):
                    continue
                got = _date_from_fragment(text)
                if got and _valid_stamp_date(got):
                    key = got.isoformat()
                    votes[key] = votes.get(key, 0) + 1
    if len(votes) != 1:
        return ""
    value, hits = max(votes.items(), key=lambda item: item[1])
    if hits < 2 or hits * 2 < sum(votes.values()):
        return ""

    return value


def extract_fields(text, jehat=None, lines_info=None, ctx=None, orig_path=None,
                   qr_paths=None):
    t = text.translate(AR_DIGITS)
    lines = [l["text"] if isinstance(l, dict) else str(l) for l in (lines_info or text.splitlines())]
    lines = [str(l).strip() for l in lines]
    t = "\n".join(lines).translate(AR_DIGITS)
    f = {"number": "", "sender": "", "recipient": "", "subject": "", "jeha": "",
         "export_number": "", "export_date": "", "body_date": "", "serial": "",
         "wared_date": "", "attachments": "",
         "_confidence": {"number": "low", "wared_date": "low",
                         "export_number": "low",
                         "export_date": "low", "sender": "low", "recipient": "low",
                         "subject": "low", "attachments": "low",
                         "jeha": "low"}}
    pos = _positional_candidates(lines_info or [])

    # ---- المناطق حسب الموضع (رأس / ختم صادر يسار فوق / ختم وارد يمين تحت) ----
    W, H = _page_dims(lines_info or [])
    regs = _page_regions(lines_info or [], W, H)
    ctx = ctx or {}
    prepped = ctx.get("prepped")
    workdir = Path(ctx.get("workdir") or TMP)
    lines_info = lines_info or []

    def find_line(keys, pool):
        for l in pool:
            tx = re.sub(r"\s+", "", _norm(l.get("text", "")))
            if tx and any(re.sub(r"\s+", "", _norm(k)) in tx for k in keys):
                return l
        return None

    # ===== ختم الوارد (أسفل يمين): «رقم الوارد» والرقم فوقه بخط اليد =====
    # مرادفات كتير لأخطاء OCR المتكررة في «رقم الوارد» (وقم القواود، قم القواود…)
    num_labels = ("رقم الوارد", "رقم وارد", "الوارد رقم", "وقم القواود",
                  "قم القواود", "رقم الواود", "رقم الورد", "رقم واود",
                  "وقم الوارد", "رقم القواود", "قم الورد", "رقم الواصد",
                  "زشم الوارد", "زشّم الوارد", "رقم الواره")
    num_line = find_line(num_labels, lines_info)
    if num_line is None:
        from difflib import SequenceMatcher
        candidates = []
        for line in lines_info:
            label = re.sub(r"\s+", "", _norm(line.get("text", "")))
            if not 5 <= len(label) <= 20 or re.search(r"صا.?ر|عامل|هاتف|\d", label):
                continue
            similarity = max(SequenceMatcher(None, label, target).ratio()
                             for target in ("رقمالوارد", "رقموارد"))
            # The printed stamp can lose dots: رقم becomes رتم or رفم,
            # and الوارد becomes اليارد. Require the complete label shape
            # and a nearby incoming marker rather than the exact first word.
            nearby_incoming = any(
                re.search(r"وارد|والد|و[١٢٣]لد", str(other.get("text", "")))
                and _line_distance(line, other, W, H) is not None
                for other in lines_info)
            if similarity >= 0.75 or (similarity >= 0.60 and nearby_incoming):
                candidates.append((similarity, line))
        if candidates:
            num_line = max(candidates, key=lambda item: item[0])[1]
    corroborated_anchor = _incoming_marker_anchor(lines_info, W, H)
    if corroborated_anchor is not None:
        if num_line is not None and _line_distance(num_line, corroborated_anchor, W, H) is not None:
            # Preserve the printed label's geometry when it already locates
            # this stamp; the date reader was calibrated against that label.
            num_line = {**num_line, '_incoming_marker_number':corroborated_anchor['_incoming_marker_number']}
        else:
            num_line = corroborated_anchor
    incoming_geometry = {}
    prepared_for_stamp = ctx.get('prepared_pages') or ([prepped] if prepped else [])
    for page_index, page_path in enumerate(prepared_for_stamp, 1):
        if num_line is not None and num_line.get('_page', 1) != page_index:
            continue
        page_lines = [line for line in lines_info if line.get('_page', 1) == page_index]
        try:
            from postal_stamp import recover_incoming_stamp
            recovered = recover_incoming_stamp(page_path, page_lines, ocr_engine(),
                _run_engine, _res_lines, _norm, AR_DIGITS, anchor=num_line)
        except Exception:
            recovered = {}
        if not recovered.get('confirmed'):
            continue
        incoming_geometry = recovered
        incoming_geometry['_page'] = page_index
        recovered_lines = recovered.get('lines', [])
        lines_info.extend(recovered_lines)
        if num_line is None:
            num_line = find_line(num_labels, recovered_lines)
            if num_line is None:
                from postal_stamp import _incoming
                num_line = next((line for line in recovered_lines
                                 if _incoming(line.get('text', ''), _norm)), None)
            if num_line is None and recovered.get('number_box'):
                x0, y0, x1, y1 = recovered['number_box']
                values = [line for line in recovered_lines
                          if x0 <= (line['x0'] + line['x1']) / 2 <= x1
                          and y0 <= (line['y0'] + line['y1']) / 2 <= y1
                          and re.fullmatch(r'\d{2,6}',
                              line.get('text', '').translate(AR_DIGITS).strip())]
                if values:
                    num_line = max(values, key=lambda line: line.get('ocr_score', 0))
        break
    wared_lines = [line for line in lines_info
                   if num_line is None or line.get("_page", 1) == num_line.get("_page", 1)]
    if incoming_geometry and num_line is not None:
        x0, y0, x1, y1 = incoming_geometry['box']
        wared_lines = [line for line in wared_lines
                       if x0 - W * .03 <= (line['x0'] + line['x1']) / 2 <= x1 + W * .03
                       and y0 - H * .01 <= (line['y0'] + line['y1']) / 2 <= y1 + H * .025]
    stamp_prepped = prepped
    prepared_pages = ctx.get("prepared_pages", [])
    if num_line is not None and prepared_pages:
        stamp_prepped = prepared_pages[int(num_line.get("_page", 1)) - 1]
    stamp_W, stamp_H = _page_dims(wared_lines)
    dig_wared = []
    if num_line is not None and stamp_prepped is not None:
        # رقم ختم الوارد الداخلي يظهر أسفل عنوان الختم في النماذج الفعلية.
        dig_wared = crop_digits_px(
            stamp_prepped,
            float(num_line["x0"]) - stamp_W * 0.03,
            float(num_line["y0"]) - stamp_H * 0.01,
            float(num_line["x1"]) + stamp_W * 0.04,
            float(num_line["y1"]) + stamp_H * 0.18,
            workdir, "wared_num", scale=5)

    # ===== ختم الصادر (أعلى شمال): «رقم الصادر» أو «مادر رقم» =====
    sader_labels = ("رقم الصادر", "صادر رقم", "مادر رقم", "ميادر رقم", "ميادر", "مادر",
                    "رقم الإشارة", "رقم الاشارة", "الاشارة رقم", "الإشارة رقم")
    sader_line = find_line(sader_labels, lines_info)
    dig_sader = []
    if sader_line is not None and prepped is not None:
        dig_sader = crop_digits_px(
            prepped,
            float(sader_line["x0"]) - W * 0.03,
            float(sader_line["y0"]) - H * 0.035,
            float(sader_line["x1"]) + W * 0.06,
            float(sader_line["y1"]) + H * 0.030,
            workdir, "sader_num", scale=5)

    f["wared_box_text"] = " | ".join(dig_wared)
    f["sader_box_text"] = " | ".join(dig_sader)

    stamp_primary, stamp_candidates, stamp_confidence = _contextual_number(
        lines_info, [num_line] if num_line else [], W, H, num_labels)
    if num_line is not None and num_line.get('_incoming_marker_number'):
        stamp_primary = num_line['_incoming_marker_number']
        stamp_candidates, stamp_confidence = [stamp_primary], 'medium'
    cell_number, cell_candidates, cell_evidence = _read_incoming_number(
        stamp_prepped, num_line, wared_lines, stamp_W, stamp_H)
    if cell_number:
        stamp_primary, stamp_confidence = cell_number, 'medium'
        stamp_candidates = list(dict.fromkeys([cell_number] + cell_candidates))
        wared_lines.extend(cell_evidence)
        lines_info.extend(line for line in cell_evidence if line not in lines_info)
    if not stamp_primary and num_line and dig_wared:
        crop_nums, _crop_dates = _digits_parse(dig_wared)
        crop_nums = [n for n in crop_nums if n.isdigit() and len(n) <= 5]
        crop_nums = list(dict.fromkeys(crop_nums))
        if len(crop_nums) == 1:
            stamp_primary, stamp_confidence = crop_nums[0], "medium"
            stamp_candidates = crop_nums
    if stamp_primary:
        dig_wared = [stamp_primary] + dig_wared
        f["wared_box_text"] = " | ".join(dict.fromkeys(dig_wared))

    # ===== عدد المرفقات: نحدد الرقم بالنسبة لعنوانه، لا بإحداثيات ثابتة =====
    # «الرفقات» مرادف مطلوب: محرك OCR بيسقط الميم من «المرفقات» أحياناً.
    attachment_keys = ("عدد المرفقات", "المرفقات", "الرفقات", "عدد الرفقات")
    attachment_line = find_line(attachment_keys, regs["bottom_right"]) \
        or find_line(attachment_keys, lines_info)
    if attachment_line is not None:
        crop_values = []
        if prepped is not None:
            crop_values = crop_digits_px(
                prepped,
                max(0, float(attachment_line["x0"]) - W * 0.20),
                max(0, float(attachment_line["y0"]) - H * 0.05),
                min(W, float(attachment_line["x1"]) + W * 0.35),
                min(H, float(attachment_line["y1"]) + H * 0.15),
                workdir, "attachments", scale=5)
        crop_candidates = set()
        for value in crop_values:
            value = str(value).translate(AR_DIGITS)
            for match in re.finditer(r"(?<!\d)(\d{1,2})(?!\d)", value):
                crop_candidates.add(match.group(1))

        label_center_y = (float(attachment_line.get("y0", 0)) + float(attachment_line.get("y1", 0))) / 2
        label_center_x = (float(attachment_line.get("x0", 0)) + float(attachment_line.get("x1", 0))) / 2
        nearby = []
        for line in lines_info:
            value = str(line.get("text", "")).translate(AR_DIGITS)
            center_y = (float(line.get("y0", 0)) + float(line.get("y1", 0))) / 2
            center_x = (float(line.get("x0", 0)) + float(line.get("x1", 0))) / 2
            if abs(center_y - label_center_y) > H * 0.06:
                continue
            if abs(center_x - label_center_x) > W * 0.35:
                continue
            for match in re.finditer(r"(?<!\d)(\d{1,2})(?!\d)", value):
                nearby.append((abs(center_x - label_center_x), match.group(1)))
        line_candidates = {value for _, value in nearby}
        if crop_candidates and line_candidates:
            agreed = crop_candidates & line_candidates
            if len(agreed) == 1:
                f["attachments"] = next(iter(agreed))
        elif len(crop_candidates) == 1 and all(len(value) <= 2 for value in crop_candidates):
            f["attachments"] = next(iter(crop_candidates))
        elif len(line_candidates) == 1 and all(len(value) <= 2 for value in line_candidates):
            f["attachments"] = next(iter(line_candidates))
        if f["attachments"]:
            f["_confidence"]["attachments"] = "medium"
        if crop_candidates or line_candidates:
            f["attachments_candidates"] = sorted(crop_candidates | line_candidates)
        # القيمة ممكن تكون في خانة تحتي في نفس العمود ومكتوبة بالحروف
        # («صفر») — القص الرقمي مالوش لازمة فيها.
        if not f["attachments"]:
            cell_value = _attachment_cell_value(attachment_line, lines_info, W, H)
            if cell_value:
                f["attachments"] = cell_value
                f["_confidence"]["attachments"] = "medium"
        # خط اليد في الخانة أسفل العنوان مباشرة مش بيتمسك من الـdetector،
        # فالقص الواسع بيجيب ضوضاء. حل: قصّين ضيّقين أسفل الخانة والرقم
        # المعتمد هو اللي يظهر في الاتنين (توافق) — يحمي من أرقام الشريط
        # والتاريخ المجاورة.
        if not f["attachments"] and prepped is not None:
            lx0, lx1 = float(attachment_line["x0"]), float(attachment_line["x1"])
            ly1 = float(attachment_line["y1"])
            votes = []
            for tag, (bx0, by0, bx1, by1) in (
                    ("att-n", (lx0, ly1, lx1, ly1 + H * 0.06)),
                    ("att-m", (lx0 - W * 0.05, ly1, lx1 + W * 0.05,
                               ly1 + H * 0.07))):
                nums = set()
                for value in crop_digits_px(prepped, bx0, by0, bx1, by1,
                                            workdir, tag, scale=5):
                    value = str(value).translate(AR_DIGITS)
                    for match in re.finditer(r"(?<!\d)(\d{1,2})(?!\d)", value):
                        nums.add(match.group(1))
                if nums:
                    votes.append(nums)
            agreed = set.intersection(*votes) if len(votes) == 2 else set()
            if len(agreed) == 1:
                f["attachments"] = next(iter(agreed))
                f["_confidence"]["attachments"] = "medium"

    # ===== التاريخ: من نص الختم العربي مباشرة (أدق من قص الأرقام) =====
    # نموذج الأرقام اللاتيني مش بيقرأ الأرقام العربية (٠١٢٣)، لكن نموذج
    # العربية العربي بيقراها تمام — فبناخد التاريخ من سطر «التاريخ» في النص.
    # أخطاء OCR شائعة في كلمة «التاريخ»
    DATE_LABELS = ("تاريخ", "تاربخ", "تاريح", "ثاريخ", "تاربسج", "الثاري")

    def date_in(pool):
        """تاريخ ملزق بعنوان «التاريخ»: نفس السطر أو سطر جنبه (فوق/تحت)."""
        for l in pool:
            tx = str(l.get("text", ""))
            n = _norm(tx)
            if not any(k in n for k in DATE_LABELS):
                continue
            d = find_date(tx)
            if d:
                return d.isoformat()
            # العنوان في سطر والقيمة في سطر تاني جنبه (شكل الختم الفعلي)
            ly0, ly1 = float(l.get("y0", 0)), float(l.get("y1", 0))
            lx0, lx1 = float(l.get("x0", 0)), float(l.get("x1", 0))
            near = []
            for o in lines_info:
                if o is l:
                    continue
                oy0, oy1 = float(o.get("y0", 0)), float(o.get("y1", 0))
                ox0, ox1 = float(o.get("x0", 0)), float(o.get("x1", 0))
                if oy0 < ly0 - H * 0.035 or oy1 > ly1 + H * 0.035:
                    continue                    # مش جنب سطر العنوان
                if ox1 < lx0 - W * 0.35 or ox0 > lx1 + W * 0.35:
                    continue                    # بعيد أفقيًا
                d = find_date(str(o.get("text", "")))
                if d:
                    near.append((abs(oy0 - ly1), d))
            if near:
                return sorted(near)[0][1].isoformat()
        return ""

    def month_date_bottom():
        """تاريخ مكتوب بالحروف (٣٠ سبتمبر ٢٠٢٦) في منطقة ختم أسفل الورقة."""
        for o in lines_info:
            if float(o.get("y0", 0)) < H * 0.78:
                continue
            d = find_month_date(str(o.get("text", "")))
            if d and _fresh(d):
                return d.isoformat()
        return ""

    wared_date_labels = ("تاريخ الوارد", "تاريخ وارد")
    direct_date, direct_readings = _read_located_stamp_date(stamp_prepped, incoming_geometry)
    if not direct_date:
        direct_date, direct_readings = _stamp_date_line(
            stamp_prepped, num_line, wared_lines, stamp_W, stamp_H
        ) if stamp_prepped is not None and num_line is not None else ("", [])
    if direct_date:
        wared_lines.extend(direct_readings)
        lines_info.extend(line for line in direct_readings if line not in lines_info)
    d_wared, wared_date_conf = _contextual_date(
        wared_lines, [num_line] if num_line else [], stamp_W, stamp_H,
        explicit=wared_date_labels, require_anchor=True, incoming=True)
    if direct_date:
        d_wared, wared_date_conf = direct_date, "medium"
    if not d_wared:
        d_wared, wared_date_conf = _contextual_date(
            wared_lines, [num_line] if num_line else [], stamp_W, stamp_H, require_anchor=True, incoming=True)
    # آخر محاولة: قص شريط الختم وقراءته بالمحرك العربي (نص عادي + Otsu).
    # الـdetector أحياناً بيقفل على التاريخ وبيرجّع جزء منه بس، فبنقص
    # الشريط تاني وبنقبل النتيجة لو طلعت تاريخ كامل قريب من النهاردة.
    if not d_wared:
        band_date = _wared_date_from_bands(stamp_prepped, num_line, wared_lines,
                                           stamp_W, stamp_H, workdir)
        if band_date:
            d_wared = band_date
            wared_date_conf = "medium" if any(
                line.get("_date_crop_confirmed") for line in wared_lines) else "low"
            for line in wared_lines:
                if line not in lines_info:
                    lines_info.append(line)
    # Keep incoming stamp lines out of the independently printed issue date.
    def independent_date_line(line):
        label = _norm(str(line.get('text', '')))
        if re.search(r'تاريخ (?:الخطاب|الكتاب|التحرير|الصادر)', label):
            return True
        if 'وارد' in label or line.get('_date_crop_confirmed'):
            return False
        if num_line is not None and _line_distance(num_line, line, W, H) is not None:
            return False
        return True
    letter_lines = [line for line in lines_info if independent_date_line(line)]
    d_sader, sader_date_conf = _contextual_date(
        letter_lines, [sader_line] if sader_line else [], W, H,
        explicit=("تاريخ الخطاب", "تاريخ الكتاب", "تاريخ الصادر"),
        require_anchor=True)
    if not d_sader:
        d_sader, sader_date_conf = _contextual_date(
            letter_lines, [sader_line] if sader_line else [], W, H,
            require_anchor=bool(sader_line))
    nums_w, dates_w = _digits_parse(dig_wared)
    nums_s, dates_s = _digits_parse(dig_sader)
    if not d_sader and len(dates_s) == 1:
        d_sader, sader_date_conf = dates_s[0], "medium"
    if d_wared:
        f["wared_date"] = d_wared
        f["_confidence"]["wared_date"] = wared_date_conf
    if d_sader:
        f["export_date"] = d_sader
        f["_confidence"]["export_date"] = sader_date_conf

    f["number"] = stamp_primary
    if not f["number"]:
        # fallback للنص العادي (من غير lines_info/مواضع): «رقم الوارد 4471 / 2026».
        # نأخذ أول رقم بعد العنوان مباشرة، ونتجاهل سنة الخطاب لوحدها.
        m = re.search(r"(?:رقم\s*الوارد|رقم\s*وارد|الوارد\s*رقم)\s*[:：\-]?\s*(\d{1,6})", t)
        if m:
            cand = m.group(1)
            if not (len(cand) == 4 and cand == (f.get("wared_date", "") or "")[:4]):
                f["number"] = cand
                f["_number_conf"] = "text"
                f["_confidence"]["number"] = "medium"
    if f["number"]:
        if f.get("_number_conf") != "text":
            f["_number_conf"] = "stamp"
            f["_confidence"]["number"] = stamp_confidence
        # رقم الختم بخط اليد قد يخلط ٣ مع ٢. تظهر البدائل بجانب الحقل
        # كمجرد اختيار بنقرة، ولا نمس أرقام الخطاب المطبوعة أو QR.
        handwriting_alternatives = _handwritten_23_candidates(f["number"])
        candidates = list(dict.fromkeys(
            list(stamp_candidates or []) + handwriting_alternatives))
        if len(candidates) > 1:
            f["number_candidates"] = candidates

    export_number, export_candidates, export_confidence = _contextual_number(
        lines_info, [sader_line] if sader_line else [], W, H, sader_labels)
    if not export_number and sader_line:
        export_number, export_candidates, export_confidence = _contextual_number(
            [{"text": "رقم الصادر " + sader_line.get("text", ""),
              **{k: sader_line[k] for k in ("x0", "x1", "y0", "y1") if k in sader_line}}],
            [{"text": "رقم الصادر " + sader_line.get("text", ""),
              **{k: sader_line[k] for k in ("x0", "x1", "y0", "y1") if k in sader_line}}],
            W, H, sader_labels)
    if export_number:
        # إذا انتهى رقم الصادر بصفر زائد بسبب قراءة الـ OCR للنقطة أو الشرطة في الختم (مثل ١٥٩٠٠ بدلاً من ١٥٩٠)
        if len(export_number) == 5 and export_number.endswith("00"):
            export_number = export_number[:-1]
        f["export_number"] = export_number
        f["_confidence"]["export_number"] = export_confidence
        cands = f.setdefault("export_candidates", [])
        if export_number not in cands:
            cands.insert(0, export_number)

    # خط اليد في الختم مش بيتقرأ reliably -> لو القراءة مش مقنعة نسيبها فاضية
    # (أحسن من رقم غلط؛ الحقل بيتلوّن أصفر وانت تكتبه من صورة الختم المعروضة)
    if (f["number"] and f.get("_number_conf") == "stamp"
            and num_line is not None
            and not _plausible_stamp(f["number"], dig_wared)):
        f["_rejected_number"] = f["number"]
        f["_confidence"]["number"] = "low"
    else:
        # لا يوجد رفض حالي: أي قراءة موثوقة (نص صريح) أو QR لاحق
        # يجب ألا يترك تحذير «تخمين منخفض» قديمًا بجانب القيمة.
        f["_rejected_number"] = ""

    # الجهة المرسلة: من الرأس، عربي، أقصى يمين، ونفضّل «قطاع/إدارة» ثم اسم الشركة.
    header = regs["header"]
    decision_document = _is_decision_document(lines)
    address_index = None if decision_document else next((i for i, line in enumerate(header)
                          if re.match(r"^(?:السيد|السادة|السيدة|الأستاذ|الاستاذ|معالي|سعادة)",
                                      str(line.get("text", "")).strip())), None)
    sender_header = header[:address_index] if address_index is not None else header
    if address_index is not None:
        address = []
        address_start = next((i for i, line in enumerate(lines_info)
                              if line is header[address_index]), None)
        address_lines = (lines_info[address_start:] if address_start is not None
                         else header[address_index:])
        for line in address_lines:
            if line.get("_page", 0) != header[address_index].get("_page", 0):
                break
            phrase = str(line.get("text", "")).strip()
            if re.search(r"تحية|وبعد|السلام عليكم", phrase):
                break
            if (len(phrase) < 7 or _arabic_ratio(phrase) < 0.6
                    or re.search(r"\b(?:صادر|وارد)\b|سبتمبر|أكتوبر|اكتوبر|تاريخ|تحية", phrase)
                    or find_date(phrase)):
                continue
            address.append(phrase)
            if len(address) >= 2:
                break
        f["recipient"] = " ".join(address)[:240]
        f["_confidence"]["recipient"] = "medium" if address else "low"
    f["sender"] = _decision_sender(lines) or _pick_header_sender(sender_header, W, H)
    if f["sender"]:
        for i, line in enumerate(sender_header[:-1]):
            if _norm(fix_entity(str(line.get("text", "")))) == _norm(f["sender"]):
                for following in sender_header[i + 1:i + 4]:
                    continuation = str(following.get("text", "")).strip()
                    if re.match(r"^لش[ؤئو]ون|^لشئون", continuation):
                        f["sender"] += " " + continuation
                        break
                break
    if not f["sender"]:
        # fallback للنص العادي: «الجهة : شركة ...» أو أول سطر فيه شركة/هيئة/وزارة.
        m = re.search(r"(?:الجهة|جهة|المرسل)\s*[:：\-]?\s*([^\n]{3,80})", text)
        if m:
            cand = m.group(1).strip()
            if any(k in cand for k in ("شركة", "هيئة", "وزارة", "قطاع", "إدارة", "مصلحة", "محافظة", "حي ")):
                f["sender"] = fix_entity(cand[:80])
    if f["sender"]:
        f["_confidence"]["sender"] = "medium"

    # ---- QR / باركود: رقم مضمون 100% (لو الورقة عليها QR) ----
    # بندور على الصورة المجهّزة الأول وبعدها على الأصلية/كل صفحات الـ PDF،
    # لأن التنظيف (تحويل للرمادي + تكبير) أحيانًا بيبطّئ قراءة الـ QR.
    if (orig_path or qr_paths) and not f["number"]:
        qr = find_qr([orig_path] + list(qr_paths or []))
        if qr and qr.get("number"):
            if qr.get("_qr_kind") == "serial":
                # QR فيه الرقم التسلسلي -> يروح خانة التسلسل مش رقم الوارد
                f["serial"] = qr["_qr_raw"]
                f["_qr_serial"] = qr["_qr_raw"]
            else:
                f["number"] = qr["number"]
                f["_qr_raw"] = qr["_qr_raw"]
                f["_confidence"]["number"] = "high"
                cands = f.setdefault("number_candidates", [])
                if f["number"] in cands:
                    cands.remove(f["number"])
                cands.insert(0, f["number"])
                f["_number_conf"] = "qr"
            f["_rejected_number"] = ""

    # الموضوع: نبدأ من متن الخطاب بعد التحية، ولا نسمح باسم الشركة أو المرسل
    # أن يتحول تلقائيًا إلى موضوع عند غياب عنوان «الموضوع».
    for i, l in enumerate(lines):
        m = re.search(r'(?:الموضوع|Subject)\s*[:：\-]?\s*(.*)', l)
        if m:
            subj = m.group(1).strip()
            # بعض الماسحات تفصل العنوان عن قيمته، أو تقطع الموضوع الطويل
            # إلى سطرين. نضم السطر التالي فقط حين تبدو القيمة ناقصة، ونقف
            # أمام التحية/التوقيع/عنوان حقل جديد حتى لا يدخل متن الخطاب كله.
            continuation_stops = ("السيد", "السادة", "الأستاذ", "الاستاذ",
                                  "تحية", "السلام", "وتفضلوا", "مع خالص",
                                  "التوقيع", "رقم الوارد", "رقم الصادر",
                                  "التاريخ", "المرفقات")
            needs_more = (not subj or len(subj) < 32 or
                          subj.endswith(("بشأن", "بخصوص", "عن", "في", "و")))
            if needs_more:
                for nxt in lines[i + 1:i + 3]:
                    nxt = nxt.strip()
                    if (not nxt or any(word in nxt for word in continuation_stops)
                            or re.search(r"^(?:الموضوع|Subject)\s*[:：\-]?", nxt)):
                        break
                    subj = (subj + " " + nxt).strip()
                    if len(subj) >= 45:
                        break
            f["subject"] = subj[:200]
            f["_confidence"]["subject"] = "medium"
            break

    if not f["subject"] and decision_document:
        f["subject"] = _decision_subject(lines)
        if f["subject"]:
            f["_confidence"]["subject"] = "medium"

    if not f["subject"] and not decision_document:
        greeting = next((i for i, l in enumerate(lines)
                         if re.search(r"تحية|وبعد|السلام عليكم", l)), 1)
        stop_words = ("وتفضلوا", "مع خالص", "التوقيع",
                      "رقم الوارد", "وارد من", "عدد المرفقات")
        body = []
        for l in lines[greeting + 1:]:
            if any(word in l for word in stop_words):
                break
            # Names at the start of body lines describe the request. They
            # must not be discarded as if they were the header addressee.
            if (len(l) >= 18
                    and not re.search(r"^(شركة|هيئة|وزارة|CAIRO AIRPORT|Egyptian)", l)):
                body.append(l)
        if body:
            combined = _fix_ocr_spelling(" ".join(body[:5]))
            has_shifts = bool(re.search(r"(?:العمل|للعمل|يعمل)\s+بنظام\s+النوبات", combined))
            focused = re.search(r"(?:بشأن|بخصوص)\s+(.+?)(?=\s+(?:مرفق|يرجى|وحيث)|$)", combined)
            if not focused:
                focused = re.search(r"(?:بإيقاف|إيقاف)\s+صرف\s+النفقة", combined)
            if not focused:
                focused = re.search(r"(?:تغيبوا?|تغيب)\s+(.+?)(?=\s+(?:وذلك|وبيانهم|برجاء|وتفضلوا)|$)", combined)
            request = re.search(r"(?:الموافقة\s+على|يرجى\s+التكرم\s+بـ?)\s*(.+?)(?=\s+(?:رقم\s+(?:العامل|عامل)|حيث|وحيث|وذلك)|$)", combined)
            subj = (focused.group(0).strip() if focused else
                    request.group(1).strip() if request else " ".join(body[:3]))
            action = re.search(r"(?:تكليف|ندب|انتداب|صرف\s+(?:البدلات|بدل|النفقة)|رفع\s+بدل)\b.+", combined)
            if action and not focused:
                subj = re.split(r"\s+(?:رقم\s+(?:العامل|عامل)|(?:كود|كوذ|كوذوذ)\s+مال[ىي]|حيث|وحيث|وذلك)\b",
                                action.group(0), maxsplit=1)[0]
            if "السيد المذكور" in subj:
                named = re.search(r"(?:من|المقدم\s+من)\s+السيد\s+(.+?)(?=\s+(?:بالإدارة|بالادارة|رقم|كود|بوظيفة))", combined)
                if named:
                    subj = subj.replace("السيد المذكور", "السيد " + named.group(1).strip())
            if re.search(r"تكليف|ندب|انتداب", subj):
                role = re.search(r"(?:للقيام\s+بأعمال|للعمل\s+بوظيفة|للقيام\s+بمهام)\s+(.+?)(?=\s*(?:[،,]\s*)?(?:وذلك|للاستفادة|لصالح)|$)", combined)
                if role and role.group(1).strip(" ،,.") not in subj:
                    subj = subj.rstrip(" ،,.") + " للقيام بأعمال " + role.group(1).strip(" ،,.")
                elif "بنظام النوبات" not in subj and has_shifts:
                    subj += " للعمل بنظام النوبات"
            if (re.search(r"نظام\s+عمل\s+النوبات\s+(?:إلى|الى)\s+نظام\s+(?:عمل\s+)?الصباحي", combined)
                    and "الصباحي" not in subj):
                subj += " لتغيير نظام العمل من النوبات إلى الصباحي"
            elif action and "نظام النوبات" not in subj and has_shifts:
                subj += " للعمل بنظام النوبات"
            subj = re.sub(r"\s+", " ", subj).strip()
            subj = re.sub(r"\s+إ?عتبار[ااً]*\s*من\s*(?:تاريخ\s*)?$", "", subj)
            f["subject"] = (subj if len(subj) <= 240 else subj[:241].rsplit(" ", 1)[0])
            # حتى لو لم تقرأ كلمة «الموضوع»، النص المستخرج من متن الخطاب
            # يظل اقتراحًا صالحًا للمستخدم. لا نتركه low لأن الواجهة تخفي
            # المنخفض؛ يعرض كمستوى متوسط وقابل للمراجعة.
            f["_confidence"]["subject"] = "medium"

    if f["subject"] and any(
            re.search(r'(?:الموضوع|Subject)\s*[:：\-]?', line)
            for line in lines):
        f["_confidence"]["subject"] = "high"

    # اقتراح الجهة المحوّل لها من قائمة القوائم (اسم قريب من مذكور في المتن)
    if jehat:
        norm_j = {_norm(j): j for j in jehat if j and _norm(j)}
        keys = list(norm_j)
        for l in lines[3:]:
            n = _norm(l)
            if len(n) < 3:
                continue
            hit = difflib.get_close_matches(n, keys, n=1, cutoff=0.72)
            if not hit:
                hit = [k for k in keys if len(k) > 4 and k in n][:1]
            if hit and norm_j[hit[0]] != f["sender"]:
                f["jeha"] = norm_j[hit[0]]
                break
    f['recipient'] = recipient_job_title(f.get('recipient'))
    f["_evidence"] = _field_evidence(f, lines_info)
    return f


# ------------------------------------------------------------------ الإكسل
# الأعمدة بتتعرّف أوتوماتيك من عناوين الصف الأول -> الأداة تشتغل على أي شكل ملف
COLMAP = {
    "num": ("م",),
    "serial": ("الرقم التسلسلي", "التسلسلي", "رقم القيد"),
    "received": ("تاريخ اليوم", "تاريخ الاستلام", "تاريخ الورود"),
    "received_time": ("وقت الاستلام",),
    "wared_date": ("تاريخ الوارد",),
    "_request_id": ("معرف عملية الحفظ",),
    "number": ("رقم وارد", "رقم الوارد", "الوارد", "رقم الكتاب", "رقم الخطاب"),
    "sender": ("الجهة المرسل", "الجهة المرسلة", "جهة المرسل", "المرسل"),
    "recipient": ("الجهة الموجه إليها", "الجهة الموجه اليها", "الموجه إليه", "الموجه اليه"),
    "subject": ("الموضوع", "موضوع"),
    "tasheer": ("التاشيره", "التأشيرة", "الاشاره", "الإشارة"),
    "tasheer_date": ("تاريخ التاشيره", "تاريخ التأشير", "تاريخ الاشاره"),
    "action": ("الاجراء", "الإجراء"),
    "action_date": ("تاريخ الاجراء", "تاريخ الإجراء"),
    "jeha": ("الجهة المحول لها", "الجهة المحول اليه", "جهة التحويل"),
    "dist_date": ("تاريخ التوزيع", "تاريخ الاحاله"),
    "status": ("حالة القرار", "الحالة"),
    "dec_number": ("رقم القرار",),
    "dec_date": ("تاريخ القرار",),
    "notes": ("ملاحظات", "ملاحظه"),
}
HEADER_ROW = int(os.environ.get("POSTAL_HEADER_ROW", "1"))
ROW_ROLES = ("serial", "number", "sender", "recipient", "subject", "received",
             "tasheer", "action", "dec_number", "notes", "jeha")


def _hdr(v):
    s = _norm(str(v or ""))
    return re.sub(r"[\s\-_/.:،()\u0640]+", "", s)


def load_layout(ws):
    """يرجّع {الدور: رقم العمود} حسب عناوين الصف الأول (مطابقة تامة الأول)."""
    heads = {c: _hdr(ws.cell(HEADER_ROW, c).value) for c in range(1, ws.max_column + 1)}
    heads = {c: h for c, h in heads.items() if h}
    layout = {}
    for depth in (1, 2):
        for role, aliases in COLMAP.items():
            if role in layout:
                continue
            for a in (_hdr(x) for x in aliases):
                if not a:
                    continue
                for c, h in heads.items():
                    if c in layout.values():
                        continue
                    ok = (h == a) if depth == 1 else (len(a) >= 4 and (a in h or h in a))
                    if ok:
                        layout[role] = c
                        break
                if role in layout:
                    break
    return layout


def target_sheet(wb):
    if SHEET and SHEET in wb.sheetnames:
        return wb[SHEET]
    for name in DATA_SHEETS:
        if name in wb.sheetnames:
            return wb[name]
    return wb.worksheets[0]


# الأعمدة الناقصة في الملف الرسمي: بتتحط في أول عمود فاضي على الشمال.
# (الدور، العنوان اللي يتكتب في الصف الأول)
EXTRA_HEADERS = (
    ("number", "رقم الوارد"),
    ("wared_date", "تاريخ الوارد"),
    ("subject", "نبذة عن الموضوع"),
    ("sender", "الجهة المرسلة"),
    ("recipient", "الجهة الموجه إليها"),
    ("received", "تاريخ الاستلام"),
    ("received_time", "وقت الاستلام"),
    ("_request_id", "معرف عملية الحفظ"),
    ("serial", "الرقم التسلسلي"),
    ("tasheer", "التأشيرة"),
    ("tasheer_date", "تاريخ التأشيرة"),
    ("action", "الإجراء"),
    ("action_date", "تاريخ الإجراء"),
)


RETIRED_HEADERS = ("تاريخ تحرير الخطاب", "تاريخ الخطاب", "تاريخ الجواب", "تاريخ الكتاب",
                   "جهة خارجيه", "الجهة الخارجية", "خارجيه")


def remove_retired_columns(ws):
    """Remove record fields explicitly retired by the user."""
    retired = {_hdr(title) for title in RETIRED_HEADERS}
    columns = [col for col in range(1, ws.max_column + 1)
               if _hdr(ws.cell(HEADER_ROW, col).value) in retired]
    if not columns:
        return 0
    from openpyxl.utils import get_column_letter
    dimensions = {col: copy.copy(ws.column_dimensions[get_column_letter(col)])
                  for col in range(1, ws.max_column + 1)}
    for col in reversed(columns):
        ws.delete_cols(col)
    ws.column_dimensions.clear()
    for old_col, dimension in dimensions.items():
        if old_col in columns:
            continue
        col = old_col - sum(removed < old_col for removed in columns)
        name = get_column_letter(col)
        dimension.index = name; dimension.min = col; dimension.max = col
        ws.column_dimensions[name] = dimension
    return len(columns)


def ensure_record_columns(ws):
    """Add the agreed record columns without renaming existing data."""
    from openpyxl.utils import get_column_letter
    if remove_retired_columns(ws):
        format_workbook_visual(ws.parent)
    layout = load_layout(ws)
    for role, title in EXTRA_HEADERS:
        if role in layout:
            continue
        col = ws.max_column + 1
        ws.cell(HEADER_ROW, col).value = title
        ws.column_dimensions[get_column_letter(col)].width = 20
        layout[role] = col
    ws.column_dimensions[get_column_letter(layout["_request_id"])].hidden = True
    return layout


REGISTRY_COLUMNS = (
    ('received','تاريخ الاستلام'), ('number','رقم الوارد'), ('wared_date','تاريخ الوارد'),
    ('sender','الجهة المرسلة'), ('recipient','الجهة الموجه إليها'),
    ('subject','نبذة عن الموضوع'),
    ('tasheer','التأشيرة'), ('tasheer_date','تاريخ التأشيرة'),
    ('action','الإجراء'), ('action_date','تاريخ الإجراء'),
    ('serial','الرقم التسلسلي'), ('received_time','وقت الاستلام'), ('_request_id','معرف عملية الحفظ'))


def rebuild_registry(wb):
    """Explicit user-requested redesign; callers must back up before applying."""
    old = target_sheet(wb)
    layout = load_layout(old)
    records = []
    for row in range(HEADER_ROW+1, old.max_row+1):
        if not _row_used(old, layout, row):
            continue
        record = {role: old.cell(row,col).value for role,col in layout.items()}
        record['recipient'] = recipient_job_title(record.get('recipient'))
        records.append(record)
    title, index = old.title, wb.worksheets.index(old)
    wb.remove(old)
    ws = wb.create_sheet(title, index)
    ws.append([label for role,label in REGISTRY_COLUMNS])
    for record in records:
        ws.append([record.get(role) for role,label in REGISTRY_COLUMNS])
    ensure_record_columns(ws)
    format_workbook_visual(wb)
    return len(records)

# جهات شائعة في Cairo Airport — عدّلها بحرية من ورقة «القوائم» في الإكسل.
DEFAULT_JEHAT = (
    "قطاع الشؤون القانونية",
    "رئيس قطاع الشؤون القانونية",
    "الإدارة القانونية",
    "قطاع التنمية الإدارية والموارد البشرية",
    "إدارة الموارد البشرية",
    "قطاع العمليات",
    "قطاع الهندسة والتجهيزات",
    "الإدارة العامة للمشتريات والمخازن",
    "الإدارة العامة للمالية",
    "الإدارة العامة للتجارة والتوزيع",
)


def ensure_workbook(apply=True):
    """يضبط الملف الرسمي: عناوين الأعمدة الفاضية + ورقة القوائم.

    بيرجّع (عمل إيه, رسالة). لو apply=False بيقول بس.
    """
    with LOCK:
        wb = _open_wb()
        try:
            ws = target_sheet(wb)
            done = []
            removed = remove_retired_columns(ws)
            if removed:
                format_workbook_visual(wb)
                done.append("حذفت الأعمدة الملغاة من السجل")
            layout = load_layout(ws)
            # 1) عناوين الأعمدة الناقصة
            missing = [(role, title) for role, title in EXTRA_HEADERS
                       if role not in layout]
            if missing:
                col = ws.max_column + 1
                for role, title in missing:
                    ws.cell(HEADER_ROW, col).value = title
                    try:                      # نفس عرض وتنسيق عمود الجوار
                        from openpyxl.utils import get_column_letter
                        src = ws.cell(HEADER_ROW, col - 1)
                        c = ws.cell(HEADER_ROW, col)
                        c.font = copy.copy(src.font)
                        c.fill = copy.copy(src.fill)
                        c.border = copy.copy(src.border)
                        c.alignment = copy.copy(src.alignment)
                        ws.column_dimensions[get_column_letter(col)].width = 16
                    except Exception:
                        pass
                    col += 1
                done.append(f"{len(missing)} عمود جديد: " + "، ".join(t for _, t in missing))
            ensure_record_columns(ws)
            # 2) ورقة القوائم (قائمة الجهات المحوّل لها)
            if LISTS not in wb.sheetnames:
                lw = wb.create_sheet(LISTS)
                lw.cell(1, 1).value = "الجهة المحول لها"
                try:
                    lw.column_dimensions["A"].width = 38
                except Exception:
                    pass
                for i, name in enumerate(DEFAULT_JEHAT, start=2):
                    if name:
                        lw.cell(i, 1).value = name
                done.append("أنشأت ورقة «القوائم»")
            elif not any(wb[LISTS].cell(r, 1).value for r in range(2, 60)):
                for i, name in enumerate(DEFAULT_JEHAT, start=2):
                    if name:
                        wb[LISTS].cell(i, 1).value = name
                done.append("ملأت ورقة «القوائم»")
            if done and apply:
                _atomic_save(wb)
            return done, ws.title
        finally:
            try:
                wb.close()
            except Exception:
                pass


def _open_wb():
    from openpyxl import load_workbook
    return load_workbook(XLSX)


def _row_used(ws, layout, r):
    for role in ROW_ROLES:
        c = layout.get(role)
        if c and ws.cell(r, c).value not in (None, ""):
            return True
    return False


def used_count(ws, layout):
    return sum(1 for r in range(HEADER_ROW + 1, ws.max_row + 1)
               if _row_used(ws, layout, r))


def first_empty_row(ws, layout):
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        if not _row_used(ws, layout, r):
            return r
    return ws.max_row + 1


def serial_prefix():
    return f"وارد-{date.today().year}-"


def suggest_serial(ws, layout):
    """الرقم اللي يكتب على الجواب الورقي (وبيتحفظ لو فيه عمود تسلسلي)."""
    prefix = serial_prefix()
    mx = 0
    c = layout.get("serial")
    if c:
        for r in range(HEADER_ROW + 1, ws.max_row + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith(prefix):
                try:
                    mx = max(mx, int(v[len(prefix):]))
                except ValueError:
                    pass
    # الملف الرسمي فيه صفوف عاملين عمود «م» بس -> نحسب أعلى رقم فيها بس المتسجل فعليًا،
    # عشان الترقيم الجاهز (858 صف فاضي) ميففزش الرقم التسلسلي لآلاف.
    ncol = layout.get("num")
    if ncol:
        for r in range(HEADER_ROW + 1, ws.max_row + 1):
            if not _row_used(ws, layout, r):
                continue
            v = ws.cell(r, ncol).value
            if isinstance(v, (int, float)):
                try:
                    mx = max(mx, int(v))
                except (ValueError, OverflowError):
                    pass
    mx = max(mx, used_count(ws, layout))
    return f"{prefix}{mx + 1:04d}"


def _atomic_save(wb):
    """حفظ آمن: ملف مؤقت -> نسخة احتياطية من القديم -> استبدال مباشر.

    بيجرّب كذا مرة لو الملف اتمسك لحظيًا (قارئ تاني / فحص أمني / إكسل):
    من غير كده أي تصادم صغير كان بيطلع للمستخدم «اقفل ملف الإكسل» وهو مقفول.
    """
    BACKUPS.mkdir(exist_ok=True)
    if FREEZE:
        try:
            target_sheet(wb).freeze_panes = FREEZE   # openpyxl بيلغي التثبيت أحيانًا
        except Exception:
            pass
    tmp = XLSX.with_suffix(".new.xlsx")
    attempts = 5
    for attempt in range(attempts):
        try:
            wb.save(tmp)
            if XLSX.exists():
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                shutil.copy2(XLSX, BACKUPS / f"{XLSX.stem}-{stamp}.xlsx")
                old = sorted(BACKUPS.glob(f"{XLSX.stem}-*.xlsx"),
                             key=lambda p: p.stat().st_mtime, reverse=True)
                for drop in old[20:]:
                    try:
                        drop.unlink()
                    except OSError:
                        pass
            tmp.replace(XLSX)
            return
        except PermissionError:
            # الملف مشغول مؤقتًا — ننظف المحاولة ونعيد بعد لحظة
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
            if attempt == attempts - 1:
                raise
            time.sleep(0.15 * (attempt + 1))



def _style_registry_row(ws, layout, row):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    subject = str(ws.cell(row,layout['subject']).value or '') if layout.get('subject') else ''
    ws.row_dimensions[row].height = max(60, min(144, 20*math.ceil(len(subject)/58)))
    manual = {layout.get(k) for k in ('tasheer','tasheer_date','action','action_date')}
    dates = {layout.get(k) for k in ('received','wared_date','tasheer_date','action_date')}
    centered = dates | {layout.get(k) for k in ('number','serial','received_time')}
    edge=Side(style='thin',color='D7E2EB')
    for col in range(1,ws.max_column+1):
        cell=ws.cell(row,col)
        cell.font=Font(name='Arial',size=12,color='172B4D',bold=col==layout.get('number'))
        cell.alignment=Alignment(horizontal='center' if col in centered else 'right',vertical='center',wrap_text=True,readingOrder=2)
        cell.border=Border(left=edge,right=edge,bottom=edge)
        color='FFF5DD' if col in manual else ('EDF4FA' if row%2==0 else 'FFFFFF')
        cell.fill=PatternFill('solid',fgColor=color)
        if col in dates: cell.number_format='dd/mm/yyyy'


def format_workbook_visual(wb):
    """Readable registry with distinct manual follow-up columns."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    ws=target_sheet(wb); layout=load_layout(ws)
    widths={'received':17,'number':14,'wared_date':18,'sender':36,
        'recipient':45,'subject':70,'tasheer':36,'tasheer_date':18,
        'action':36,'action_date':18,'serial':23,'received_time':16}
    manual={layout.get(k) for k in ('tasheer','tasheer_date','action','action_date')}
    ws.sheet_view.showGridLines=False; ws.sheet_view.rightToLeft=True
    ws.sheet_properties.tabColor='24587A'; ws.freeze_panes=FREEZE or 'D2'
    ws.row_dimensions[1].height=42
    for col in range(1,ws.max_column+1):
        cell=ws.cell(1,col)
        role=next((k for k,v in layout.items() if v==col),'')
        ws.column_dimensions[get_column_letter(col)].width=widths.get(role,20)
        cell.font=Font(name='Arial',size=13,bold=True,color='FFFFFF')
        cell.fill=PatternFill('solid',fgColor='936313' if col in manual else '24587A')
        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True,readingOrder=2)
    for row in range(2,max(ws.max_row,3)+1): _style_registry_row(ws,layout,row)
    if layout.get('_request_id'): ws.column_dimensions[get_column_letter(layout['_request_id'])].hidden=True
    ws.auto_filter.ref=f'A1:L{max(ws.max_row,3)}'
    ws.print_title_rows='1:1'; ws.print_area=f'A1:L{max(ws.max_row,3)}'
    ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True


def format_official_workbook():
    """تطبيق التنسيق على السجل الرسمي مع نسخة احتياطية آمنة."""
    with LOCK:
        wb = _open_wb()
        try:
            format_workbook_visual(wb)
            _atomic_save(wb)
            return XLSX.name
        finally:
            try:
                wb.close()
            except Exception:
                pass


def _shift_formula(f, src_row, dst_row):
    """=IF(K2="",...) -> =IF(K50="",...) بنقل رقم الصف بس."""
    def rep(m):
        n = int(m.group(2))
        return m.group(1) + str(n - src_row + dst_row) if n == src_row else m.group(0)
    return re.sub(r"([A-Z]{1,3}\$?)(\d{1,7})", rep, f)


def _clone_row_style(ws, src_row, dst_row):
    """نسخ التنسيق والصيغ من صف موجود عشان الصف الجديد يطلع بنفس المنظر."""
    if src_row < 2 or dst_row == src_row:
        return
    for c in range(1, ws.max_column + 1):
        s = ws.cell(src_row, c)
        d = ws.cell(dst_row, c)
        d.font = copy.copy(s.font)
        d.border = copy.copy(s.border)
        d.fill = copy.copy(s.fill)
        d.number_format = s.number_format
        d.alignment = copy.copy(s.alignment)
        if isinstance(s.value, str) and s.value.startswith("="):
            d.value = _shift_formula(s.value, src_row, dst_row)


def read_meta():
    # تحت LOCK: أي قارئ متزامن مع الحفظ كان بيسبب PermissionError
    # على tmp.replace (القارئ بيفتح الملف من غير مشاركة الحذف على ويندوز).
    with LOCK:
        wb = _open_wb()
        try:
            ws = target_sheet(wb)
            layout = load_layout(ws)
            jehat = []
            if LISTS in wb.sheetnames:
                lw = wb[LISTS]
                for r in range(2, 60):
                    v = lw.cell(r, 1).value
                    if v:
                        jehat.append(str(v))
            heads = {role: str(ws.cell(HEADER_ROW, c).value or "").strip()
                     for role, c in layout.items()}
            dec_sheet = next((n for n in wb.sheetnames if n in DEC_SHEETS), "")
            out_sheet = next((n for n in wb.sheetnames if n in OUT_SHEETS), "")
            return {"serial": suggest_serial(ws, layout), "jehat": jehat,
                    "today": date.today().isoformat(), "sheet": ws.title,
                    "roles": sorted(layout), "heads": heads,
                    "rows": used_count(ws, layout), "file": XLSX.name,
                    "dec_sheet": dec_sheet, "out_sheet": out_sheet,
                    "next_row": first_empty_row(ws, layout)}
        finally:
            wb.close()



def _d(s):
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


def _norm(s):
    """تطبيع عربي بسيط للبحث: تشكيل + همزات + ة/ه."""
    s = (s or "").translate(AR_DIGITS)
    s = re.sub(r"[\u064b-\u0652]", "", s)
    for a, b in (("أ", "ا"), ("إ", "ا"), ("آ", "ا"), ("ة", "ه"), ("ى", "ي")):
        s = s.replace(a, b)
    return s.lower().strip()


def _safe_name(s):
    return re.sub(r'[\\/:*?"<>|\s]+', "-", str(s or "")).strip() or "scan"


_DOC_FIELDS = (
    ("serial", "الرقم التسلسلي"),
    ("received", "تاريخ اليوم"),
    ("received_time", "وقت الاستلام"),
    ("number", "رقم الوارد"),
    ("wared_date", "تاريخ الوارد"),
    ("export_number", "رقم الصادر بالخطاب"),
    ("export_date", "تاريخ الصادر بالخطاب"),
    ("sender", "الجهة المرسل منها"),
    ("recipient", "الجهة الموجه إليها"),
    ("subject", "الموضوع"),
    ("jeha", "الجهة المحول لها"),
    ("tasheer", "التأشيرة"),
    ("tasheer_date", "تاريخ التأشيرة"),
    ("action", "الإجراء"),
    ("action_date", "تاريخ الإجراء"),
    ("status", "حالة القرار"),
    ("dec_number", "رقم القرار"),
    ("dec_date", "تاريخ القرار"),
    ("notes", "ملاحظات"),
)


# درجات الثقة بالعربي للعرض في تقرير Word
_CONF_AR = {"high": "عالية", "medium": "متوسطة", "low": "منخفضة",
            "reviewed": "راجعه الموظف"}


def _write_record_docx(path: Path, serial: str, fields: dict, ocr_text: str):
    """Readable Arabic record sheet followed by the complete OCR text on a new page."""
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor, Cm

    doc = Document()
    doc.core_properties.title = f"بيانات الخطاب الوارد {fields.get('number') or serial}"
    doc.core_properties.subject = "بيانات التسجيل والنص الكامل للخطاب"
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(1.8)
    section.left_margin = section.right_margin = Cm(2)
    for name, size in [('Normal', 14), ('Title', 21), ('Heading 1', 17)]:
        style = doc.styles[name]
        style.font.name, style.font.size = 'Arial', Pt(size)
        style.font.color.rgb = RGBColor.from_string('172B4D' if name != 'Title' else '000000')
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = 1.15
        ppr = style.element.get_or_add_pPr()
        for border in list(ppr.findall(qn('w:pBdr'))): ppr.remove(border)
        rpr = style.element.get_or_add_rPr()
        fonts = rpr.rFonts
        fonts.set(qn('w:cs'), 'Arial')
        cs = OxmlElement('w:szCs'); cs.set(qn('w:val'), str(size*2)); rpr.append(cs)

    def rtl(p, size=None):
        p._p.get_or_add_pPr().append(OxmlElement('w:bidi'))
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.widow_control = True
        for run in p.runs:
            run.font.name = 'Arial'
            if size: run.font.size = Pt(size)
            rpr = run._element.get_or_add_rPr()
            fonts = rpr.rFonts
            if fonts is None:
                fonts = OxmlElement('w:rFonts'); rpr.insert(0,fonts)
            for key in ('ascii','hAnsi','eastAsia','cs'): fonts.set(qn('w:'+key),'Arial')
            direction = OxmlElement('w:rtl')
            if not re.search(r'[\u0600-\u06ff]',run.text): direction.set(qn('w:val'),'0')
            rpr.append(direction)
            if size:
                cs = OxmlElement('w:szCs'); cs.set(qn('w:val'),str(size*2)); rpr.append(cs)

    title = doc.add_paragraph('بيانات الخطاب الوارد', style='Title')
    rtl(title); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph('رقم الوارد  ' + str(fields.get('number') or 'غير محدد'))
    rtl(subtitle,14); subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(14)
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width, table.columns[1].width = Cm(4.3), Cm(12.7)
    pr = table._tbl.tblPr
    pr.insert_element_before(OxmlElement('w:bidiVisual'), 'w:tblW','w:jc','w:tblLayout')
    margins = OxmlElement('w:tblCellMar')
    for side, amount in [('top',100),('bottom',100),('left',140),('right',140)]:
        el=OxmlElement('w:'+side); el.set(qn('w:w'),str(amount)); el.set(qn('w:type'),'dxa'); margins.append(el)
    pr.append(margins)
    borders=OxmlElement('w:tblBorders')
    for side in ('top','bottom','left','right','insideH','insideV'):
        el=OxmlElement('w:'+side); el.set(qn('w:val'),'single'); el.set(qn('w:sz'),'4'); el.set(qn('w:color'),'D8E2EC'); borders.append(el)
    pr.append(borders)

    def fill_row(row, values, header=False, alternate=False):
        row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
        for i,(cell,value) in enumerate(zip(row.cells,values)):
            cell.width = Cm(4.3 if i == 0 else 12.7)
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p=cell.paragraphs[0]; run=p.add_run(str(value)); run.bold=header or i==0
            rtl(p,13); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.1
            color='24587A' if header else ('EDF4FA' if alternate else 'FFFFFF')
            shade=OxmlElement('w:shd'); shade.set(qn('w:fill'),color); cell._tc.get_or_add_tcPr().append(shade)
            if header: run.font.color.rgb=RGBColor(255,255,255)

    fill_row(table.rows[0],('البيان','التفاصيل'),header=True)
    table.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
    report_fields=[('received','تاريخ الاستلام'),('received_time','وقت الاستلام'),
        ('number','رقم الوارد'),('wared_date','تاريخ ختم الوارد'),
       ('sender','الجهة المرسلة'),
        ('recipient','الجهة الموجه إليها'),
        ('subject','نبذة الموضوع'),('serial','الرقم التسلسلي')]
    for key,label in report_fields:
        value=fields.get(key) or (serial if key=='serial' else '')
        if not value: continue
        if key in ('received','wared_date'):
            parsed=_d(str(value))
            if parsed: value=parsed.strftime('%d/%m/%Y')
        fill_row(table.add_row(),(label,value),alternate=len(table.rows)%2==0)

    heading=doc.add_paragraph('النص الكامل للخطاب',style='Heading 1')
    rtl(heading); heading.paragraph_format.page_break_before=True
    for line in str(ocr_text or '').splitlines() or ['لم يُستخرج نص من الصورة.']:
        paragraph=doc.add_paragraph(line)
        rtl(paragraph,14)
        paragraph.paragraph_format.space_after=Pt(0)
        paragraph.paragraph_format.line_spacing=Pt(18)
    footer=section.footer.paragraphs[0]
    footer.add_run('صفحة ')
    field=OxmlElement('w:fldSimple'); field.set(qn('w:instr'),'PAGE'); footer._p.append(field)
    rtl(footer,10); footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def _prepare_row(ws, layout, row):
    """تنسيق الصف الجديد + ترقيم «م» + نسخ الصيغ من الصف اللي فوقه."""
    if row > ws.max_row:
        _clone_row_style(ws, row - 1, row)
    for col in range(1, ws.max_column + 1):
        if ws.cell(row, col).value not in (None, ""):
            continue
        prev = ws.cell(row - 1, col).value
        if isinstance(prev, str) and prev.startswith("="):
            ws.cell(row, col).value = _shift_formula(prev, row - 1, row)
        elif isinstance(prev, (int, float)) and layout.get("num") == col:
            ws.cell(row, col).value = int(prev) + 1


def _record_stem(fields, serial, row):
    """اسم مجلد الورقة: «رقم الوارد-تاريخ اليوم» (مثال 3966-2026-10-05).

    لو اتسجلت ورقتين بنفس الرقم في نفس اليوم بنضيف الرقم التسلسلي
    عشان من نطمس أول ورقة، ومن غير ما نمنع حفظ التانية.
    """
    today = date.today().isoformat()
    number = str(fields.get("number") or "").strip()
    stem = _safe_name(f"{number}-{today}") if number else _safe_name(f"{serial}-{today}")
    if not (SCANS / stem).exists():
        return stem
    alt = _safe_name(f"{stem}-{_safe_name(serial)}")
    if not (SCANS / alt).exists():
        return alt
    return _safe_name(f"{alt}-s{row}")


def save_row(fields, scan_path: Path = None, ocr_text: str = "", *,
             archive_source=True, request_id=""):
    """يحفظ صف السجل، ويحفظ معه مجلد الورقة (صورة المسح + تقرير Word).

    الإكسل هو المرجع الأساسي: كل الحقول بتتكتب في صف واحد. ومجلد الورقة في
    scans بيتسمى «رقم الوارد-تاريخ اليوم» (مثال 3966-2026-10-05) وبيحوي
    ملفين فقط: الصورة الأصلية وتقرير Word بنفس الاسم.
    """
    with LOCK:
        wb = _open_wb()
        staged, archived = [], []
        try:
            ws = target_sheet(wb)
            layout = ensure_record_columns(ws)
            if request_id:
                for saved_row in range(HEADER_ROW + 1, ws.max_row + 1):
                    if ws.cell(saved_row, layout["_request_id"]).value == request_id:
                        return str(ws.cell(saved_row, layout["serial"]).value), saved_row
            next_serial = suggest_serial(ws, layout)
            fields = dict(fields)
            now = datetime.now()
            fields["received"] = now.date().isoformat()
            fields["received_time"] = now.strftime("%H:%M:%S")
            fields["_request_id"] = request_id
            if not layout.get("number") and not layout.get("subject"):
                raise ValueError(
                    f"مش لاقي أعمدة مناسبة في ورقة «{ws.title}» — راجع عناوين الصف الأول.")
            row = first_empty_row(ws, layout)
            _prepare_row(ws, layout, row)
            conv = {"received": _d, "wared_date": _d, "tasheer_date": _d,
                    "action_date": _d, "dist_date": _d, "dec_date": _d}
            written = {}
            for role in COLMAP:
                col = layout.get(role)
                if not col:
                    continue
                val = fields.get(role)
                val = val.strip() if isinstance(val, str) else val
                if val in (None, ""):
                    continue
                if role in conv:
                    val = conv[role](val)
                    if not val:
                        continue
                ws.cell(row, col).value = val
                written[role] = val
            if layout.get("received") and "received" not in written:
                ws.cell(row, layout["received"]).value = date.today()
            serial = (fields.get("serial") or "").strip()
            if serial and layout.get("serial"):
                ws.cell(row, layout["serial"]).value = serial
            if not serial:
                serial = next_serial
            # Always persist the generated serial, and avoid a stale UI suggestion.
            if any(ws.cell(r, layout["serial"]).value == serial
                   for r in range(HEADER_ROW + 1, ws.max_row + 1) if r != row):
                serial = next_serial
            ws.cell(row, layout["serial"]).value = serial
            fields["serial"] = serial
            if scan_path is not None or ocr_text:
                if archive_source and scan_path is not None and not scan_path.exists():
                    raise ValueError(f"ملف المسح الأصلي غير موجود: {scan_path.name}")
                SCANS.mkdir(exist_ok=True)
                # كل ورقة مسح في مجلد مستقل: الصورة + تقرير Word بنفس الاسم،
                # والإكسل يفضل المرجع الأساسي للبحث والعدّ.
                stem = _record_stem(fields, serial, row)
                record_dir = SCANS / stem
                token = uuid.uuid4().hex
                staged_dir = SCANS / f".{token}.tmp"
                staged_dir.mkdir()
                staged.append(staged_dir)
                docx = staged_dir / f"{_safe_name(fields.get('number') or serial)}.docx"
                if archive_source and scan_path is not None:
                    original = staged_dir / f"{stem}{scan_path.suffix.lower()}"
                    shutil.copy2(scan_path, original)
                    if scan_path.suffix.lower() in ('.png', '.jpg', '.jpeg'):
                        from PIL import Image
                        try:
                            with Image.open(scan_path) as scanned:
                                view = _trim_scanner_blank(scanned)
                                if view.size != scanned.size:
                                    view.save(original, dpi=scanned.info.get('dpi', (300, 300)))
                        except OSError:
                            pass  # Preserve an opaque original instead of altering it.
                _write_record_docx(docx, serial, fields, ocr_text)
                # مجلد الورقة = ملفان فقط: الصورة + Word. لا قصاصات ولا
                # processed ولا analysis.json — الإكسل هو المرجع الأساسي.
                # Windows may briefly lock newly copied scans during indexing
                # or antivirus inspection. Keep the commit atomic and retry;
                # never overwrite a pre-existing archive directory.
                for attempt in range(6):
                    try:
                        if record_dir.exists():
                            raise FileExistsError(f"Archive already exists: {record_dir.name}")
                        staged_dir.replace(record_dir)
                        break
                    except PermissionError:
                        if attempt == 5:
                            raise
                        time.sleep(.15 * (attempt + 1))
                staged.remove(staged_dir)
                archived.append(record_dir)
            _style_registry_row(ws, layout, row)
            ws.auto_filter.ref = f'A1:L{max(ws.max_row, row)}'
            ws.print_area = f'A1:L{max(ws.max_row, row)}'
            _atomic_save(wb)
            return serial, row
        except Exception:
            for path in staged + archived:
                try:
                    if path.is_dir():
                        shutil.rmtree(path)
                    else:
                        path.unlink()
                except OSError:
                    pass
            raise
        finally:
            try:
                wb.close()
            except Exception:
                pass


def _cell_str(ws, r, c):
    if not c:
        return ""
    v = ws.cell(r, c).value
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return "" if v is None else str(v)


def _row_info(ws, layout, r):
    info = {"row": r}
    for role in COLMAP:
        c = layout.get(role)
        if c:
            info[role] = _cell_str(ws, r, c)
    for role in ("serial", "received", "number", "wared_date", "sender",
                 "subject", "jeha", "status", "dec_number", "dec_date"):
        info.setdefault(role, "")
    return info


def find_rows(q="", only_open=False, limit=40):
    """بحث في السجل بالرقم التسلسلي / رقم الوارد / الجهة / الموضوع.

    يدعم البحث متعدد الكلمات: كل كلمة (3+ حروف) يجب أن توجد في الصف.
    مثال: «محمد الجيد» يجد «محمد أشرف عبد الجيد» حتى مع تخطي كلمة.
    """
    q = _norm(q)
    words = [w for w in q.split() if len(w) >= 2]
    out = []
    with LOCK:                      # قراءة تحت القفل: متتآزرة مع الحفظ = فشل الاستبدال
        wb = _open_wb()
        try:
            ws = target_sheet(wb)
            layout = load_layout(ws)
            for r in range(HEADER_ROW + 1, ws.max_row + 1):
                if not _row_used(ws, layout, r):
                    continue
                info = _row_info(ws, layout, r)
                if words:
                    hay = _norm(" | ".join(info.get(k, "") for k in
                                           ("serial", "number", "sender", "subject",
                                            "dec_number", "tasheer", "action", "jeha")))
                    if not all(w in hay for w in words):
                        continue
                if only_open and (info.get("dec_number") or info.get("dec_date")
                                  or info.get("action") or info.get("action_date")):
                    continue
                out.append(info)
                if len(out) >= limit:
                    break
        finally:
            wb.close()
    return out


DEC_ROLES = ("tasheer", "tasheer_date", "action", "action_date",
             "status", "dec_number", "dec_date")


def save_decision(row, values=None, **kw):
    """تسجيل القرار على صف موجود — يكتب بس الأعمدة الموجودة في الملف."""
    if values is None:
        values = kw
    row = int(row)
    with LOCK:
        wb = _open_wb()
        try:
            ws = target_sheet(wb)
            layout = load_layout(ws)
            if not _row_used(ws, layout, row):
                raise ValueError(f"الصف {row} فاضي — اختار صف صح من البحث.")
            written = {}
            for role in DEC_ROLES:
                col = layout.get(role)
                if not col:
                    continue
                v = values.get(role)
                if v in (None, ""):
                    continue
                if role.endswith("_date"):
                    v = _d(v)
                    if not v:
                        continue
                else:
                    v = str(v).strip()
                ws.cell(row, col).value = v
                written[role] = v
            # منع: لو الحالة فاضية ورقم القرار موجود -> «صدر»
            if layout.get("status") and not written.get("status") \
                    and (written.get("dec_number") or written.get("action")):
                ws.cell(row, layout["status"]).value = values.get("status") or "صدر"
            _atomic_save(wb)
            return _row_info(ws, layout, row)
        finally:
            try:
                wb.close()
            except Exception:
                pass


def append_out_register(values, subject=""):
    """صف جديد في ورقة «الصادر»: (التاريخ=اليوم، رقم الصادر، التاريخ، الموضوع)."""
    with LOCK:
        wb = _open_wb()
        try:
            ws = None
            for name in OUT_SHEETS:
                if name in wb.sheetnames:
                    ws = wb[name]
                    break
            if ws is None:
                return None
            layout = load_layout(ws)
            row = first_empty_row(ws, layout)
            _prepare_row(ws, layout, row)
            date_cols = [c for c in range(1, ws.max_column + 1)
                         if _hdr(ws.cell(HEADER_ROW, c).value) in ("تاريخ", "تاريخالصادر")]
            if date_cols:
                ws.cell(row, date_cols[0]).value = _d(values.get("received")) or date.today()
            num_col = layout.get("number") or layout.get("export_number")
            if num_col:
                ws.cell(row, num_col).value = (values.get("export_number")
                                               or values.get("number") or "").strip() or None
            if layout.get("subject"):
                ws.cell(row, layout["subject"]).value = (subject or "").strip() or None
            _atomic_save(wb)
            return row
        finally:
            try:
                wb.close()
            except Exception:
                pass


def append_decision_register(values, subject=""):
    """صف جديد في ورقة «القرارات»: (التاريخ=اليوم، رقم القرار، التاريخ=تاريخ القرار، الموضوع)."""
    with LOCK:
        wb = _open_wb()
        try:
            ws = None
            for name in DEC_SHEETS:
                if name in wb.sheetnames:
                    ws = wb[name]
                    break
            if ws is None:
                return None
            layout = load_layout(ws)
            row = first_empty_row(ws, layout)
            _prepare_row(ws, layout, row)
            date_cols = [c for c in range(1, ws.max_column + 1)
                         if _hdr(ws.cell(HEADER_ROW, c).value) in ("تاريخ", "تاريخالقرار")]
            if date_cols:
                ws.cell(row, date_cols[0]).value = _d(values.get("received")) or date.today()
            if len(date_cols) > 1 and values.get("dec_date"):
                ws.cell(row, date_cols[1]).value = _d(values.get("dec_date"))
            if layout.get("dec_number"):
                ws.cell(row, layout["dec_number"]).value = (values.get("dec_number") or "").strip() or None
            if layout.get("subject"):
                ws.cell(row, layout["subject"]).value = (subject or "").strip() or None
            _atomic_save(wb)
            return row
        finally:
            try:
                wb.close()
            except Exception:
                pass


# ------------------------------------------------------------------ المسح من السكانر (WIA)
def wia_list():
    """أجهزة المسح المتاحة عبر WIA."""
    import subprocess
    try:
        r = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", str(BASE / "scan_wia.ps1"), "-Action", "list"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=60)
        out = (r.stdout or "").strip().splitlines()
        payload = json.loads(out[-1]) if out else {}
        if isinstance(payload, dict) and payload.get("ok"):
            return payload
        return {"ok": False, "devices": payload.get("devices", []),
                "error": payload.get("error", "WIA غير متاح")}
    except Exception as e:
        return {"ok": False, "devices": [], "error": str(e)}


def wia_scan(device=-1, dpi=300, color="gray", source="auto"):
    """مسح من السكانر لملف PNG ثم رجوع مساره."""
    import subprocess
    out_file = TMP / f"scan_{uuid.uuid4().hex[:8]}.png"
    args = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(BASE / "scan_wia.ps1"), "-Action", "scan",
            "-Out", str(out_file), "-Dpi", str(int(dpi)),
            "-Color", str(color), "-Source", str(source)]
    try:
        dev = int(device)
    except (TypeError, ValueError):
        dev = -1
    if dev >= 0:
        args += ["-Device", str(dev)]
    r = subprocess.run(args, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=300)
    out = (r.stdout or "").strip().splitlines()
    payload = None
    if out:
        for line in reversed(out):
            try:
                payload = json.loads(line)
                break
            except Exception:
                continue
    if payload is None:
        raise ValueError("المسح مرجّع من غير نتيجة: " + ((r.stderr or "") or "").strip()[:300])
    if not payload.get("ok"):
        err = str(payload.get("error") or payload.get("msg") or "فشل المسح")
        if err == "NO_SCANNER":
            raise ValueError("مفيش سكانر متوصّل — شوف تعريف السكانر الأول.")
        raise ValueError(err)
    if not out_file.exists():
        raise ValueError("الصورة الممسوحة مش موجودة على الجهاز.")
    from PIL import Image
    with Image.open(out_file) as scanned_image:
        if scanned_image.format != "PNG":
            encoded = out_file.with_name(out_file.stem + "_encoded.png")
            scanned_image.save(encoded, format="PNG")
        else:
            encoded = None
    # On Windows the scanner driver or an image decoder can still hold the
    # transfer file open. Read the new PNG directly instead of replacing it.
    # Keep the original transfer intact so the scan can always be recovered.
    return encoded if encoded is not None else out_file


# ------------------------------------------------------------------ الويب
# ------------------------------------------------------------------ الواجهة (CSS مشترك)
# ستايل واحد للصفحتين بهوية شركة ميناء القاهرة الجوي (كحلي/فيروزي/ذهبي).
_UI_CSS = r"""
*{box-sizing:border-box}
:root{
 --navy:#071b33;--navy-2:#0b3c5b;--blue:#0b4f71;--blue-2:#086789;--teal:#0a5a68;
 --cyan:#2bd3c6;--gold:#f4bd62;--paper:#f4f7f6;--ink:#122c3d;--muted:#5d7a86;
 --line:#dbe5e4;--line-2:#c7d6d5;--white:#fff;--ok:#0e7a4b;--err:#b3282d;--info:#1c4f8f;
 --r-sm:10px;--r:14px;--r-lg:20px;--r-pill:999px;
 --sh-1:0 1px 2px #0b254013,0 8px 22px #0b254017;--sh-2:0 16px 38px #12324c20;
 --focus:0 0 0 3px #2bd3c659;
}
html{scroll-behavior:smooth}
body{font-family:"Segoe UI",Tahoma,"Noto Naskh Arabic",sans-serif;margin:0;color:var(--ink);
 background:var(--paper);line-height:1.7;-webkit-font-smoothing:antialiased}
body:before{content:"";position:fixed;inset:0;z-index:-2;pointer-events:none;
 background:radial-gradient(1100px 460px at 92% -12%,#d9eceb 0,transparent 62%),
            linear-gradient(135deg,#eaf2f0 0,#f8faf6 54%,#e8f0f6 100%)}
:focus-visible{outline:0;box-shadow:var(--focus);border-radius:8px}
::selection{background:#2bd3c644}
h1,h2,h3{margin:0;line-height:1.4}

/* ===== الهيدر + ديكور الرادار + الطيارات ===== */
header{position:relative;isolation:isolate;overflow:hidden;color:#fff;
 padding:24px clamp(18px,4vw,52px);display:flex;align-items:center;gap:18px;flex-wrap:wrap;
 justify-content:space-between;
 background:linear-gradient(118deg,var(--navy) 0,var(--navy-2) 60%,var(--teal) 100%);
 box-shadow:0 10px 30px #071b332e}
.header-deco{position:absolute;inset:0;z-index:0;overflow:hidden;pointer-events:none;
 background:radial-gradient(circle at 74% 130%,#2bd3c61f,transparent 56%)}
.radar-scope{position:absolute;right:6%;top:50%;width:320px;height:320px;transform:translateY(-50%);
 border-radius:50%;overflow:hidden;opacity:.6}
.radar-scope i{position:absolute;inset:0;display:block;border-radius:50%}
.radar-scope .ring{border:1px solid #47d6d64d;box-shadow:0 0 0 34px #47d6d514,0 0 0 76px #47d6d50d}
.radar-scope .sweep{background:conic-gradient(from 0deg,#64efe000 0 316deg,#64efe03d 348deg,#64efe000 360deg);
 animation:radar-turn 7s linear infinite}
.radar-scope .dot{inset:auto;width:7px;height:7px;left:31%;top:57%;background:var(--cyan);
 box-shadow:0 0 12px var(--cyan);animation:pulse 3.8s ease-in-out infinite}
.radar-scope .dot.b{left:61%;top:33%;animation-delay:1.5s}
@keyframes radar-turn{to{transform:rotate(360deg)}}
@keyframes pulse{0%,100%{opacity:.25;transform:scale(.85)}50%{opacity:1;transform:scale(1.12)}}
/* الطيارات: تطير بمقدمتها للأمام (rotate 180 للاتجاه LTR الصحيح) */
.flight{position:absolute;z-index:0;color:var(--gold);font-size:19px;line-height:1;opacity:0;
 text-shadow:0 0 10px #f4bd6266;will-change:transform;pointer-events:none}
.flight.a{top:52%;right:-6%;font-size:26px;animation:fly 16s linear infinite}
.flight.b{top:72%;right:-13%;font-size:19px;animation:fly 21s linear infinite -7s}
.flight.c{top:86%;right:-9%;font-size:15px;animation:fly 26s linear infinite -14s}
@keyframes fly{
 0%{transform:translate3d(0,0,0) rotate(180deg) scale(.9);opacity:0}
 8%{opacity:.7}55%{opacity:.55}88%{opacity:.4}
 100%{transform:translate3d(-120vw,-8px,0) rotate(180deg) scale(1.05);opacity:0}}
/* ===== الهيدر الكبير: ضعف الحجم + تناسق كامل ===== */
header{min-height:230px;padding:clamp(26px,4.5vw,48px) clamp(22px,5vw,60px)}
.brand{display:flex;align-items:center;gap:20px;position:relative;z-index:2;min-width:0}
.brand img{width:clamp(120px,15vw,180px);height:clamp(120px,15vw,180px);object-fit:contain;
 background:#fff;border-radius:28px;padding:10px;border:4px solid #ffffffcc;
 box-shadow:0 16px 34px #00152466;flex:none}
/* مقياس أحجام موحّد: العنوان والسطر الفرعي كل في سطر مستقل (display:block) */
.brand>div{min-width:0}
.brand strong{display:block;font-size:clamp(22px,3vw,34px);font-weight:800;line-height:1.4}
.brand small{display:block;font-size:clamp(13px,1.6vw,17px);font-weight:600;
 margin-top:5px;line-height:1.6}
.navlinks{position:relative;z-index:2;display:flex;gap:12px;flex-wrap:wrap;align-items:center}
.navlinks a,.btn-ghost{color:#fff;border:1px solid #ffffff3d;background:#ffffff14;
 border-radius:var(--r-pill);text-decoration:none;cursor:pointer;
 transition:background .2s,border-color .2s,transform .2s;
 padding:13px 24px;font-size:15px;font-weight:700}
.navlinks a:hover,.btn-ghost:hover{background:#ffffff26;border-color:#ffffff8c;transform:translateY(-1px)}
"""


_UI_CSS += r"""
.card{background:#ffffffd6;border:1px solid var(--white);backdrop-filter:blur(12px);
 border-radius:var(--r-lg);padding:20px;box-shadow:var(--sh-1)}
.view-title,.sec-title{display:flex;align-items:center;gap:9px;font-weight:700;font-size:15px;margin:0 0 10px}
.view-title:before,.sec-title:before{content:"";width:4px;height:17px;border-radius:3px;
 background:linear-gradient(180deg,var(--cyan),var(--teal))}
.view-title span,.sec-title span{color:var(--muted);font-weight:400;font-size:12px;margin-inline-start:auto}
label{display:block;margin:12px 0 4px;font-size:13px;font-weight:600;color:#3d5b68}
input,select,textarea{width:100%;padding:10px 11px;font-size:15px;font-family:inherit;color:var(--ink);
 background:#fff;border:1px solid var(--line-2);border-radius:var(--r-sm);
 transition:border-color .2s,box-shadow .2s}
input:focus,select:focus,textarea:focus{border-color:var(--cyan);box-shadow:var(--focus);outline:0}
input[type=checkbox]{width:auto;margin:0}
textarea{resize:vertical;line-height:1.6}
.warn{border-color:#e0a800!important;background:#fffaea!important;box-shadow:0 0 0 3px #e0a80026}
button{background:linear-gradient(180deg,var(--blue-2),var(--blue));color:#fff;border:0;
 border-radius:var(--r-sm);padding:12px 20px;font-size:15px;font-family:inherit;font-weight:700;
 cursor:pointer;margin-top:16px;box-shadow:0 6px 16px #0b4f712e;
 transition:transform .18s,box-shadow .18s,filter .18s}
button:hover{transform:translateY(-2px);box-shadow:0 10px 22px #0b4f7138;filter:brightness(1.06)}
button:active{transform:translateY(0);filter:brightness(.97)}
button:disabled{opacity:.55;cursor:not-allowed;transform:none;box-shadow:none;filter:grayscale(.25)}
button.sec{background:#f1f6f6;color:#123a4d;border:1px solid var(--line-2);box-shadow:none}
button.sec:hover{background:#e6efef;box-shadow:0 6px 14px #12324c14}
table{border-collapse:separate;border-spacing:0;width:100%;font-size:13px;margin-top:12px;
 background:#fff;border:1px solid var(--line);border-radius:var(--r);overflow:hidden}
th,td{padding:9px 10px;text-align:right;border-bottom:1px solid #edf2f2}
th{background:#f0f6f5;color:#38555f;font-weight:700;font-size:12.5px}
tbody tr:last-child td{border-bottom:0}
tbody tr:hover td{background:#f6fbfa}
#msg{margin:14px 0;padding:12px 14px;border-radius:var(--r-sm);display:none;white-space:pre-line;
 font-size:14px;border:1px solid transparent;border-inline-start:4px solid transparent}
.ok{background:#e8f7ef;border-color:#b7e2cd;border-inline-start-color:var(--ok);color:#0b5c3a;display:block!important}
.err{background:#fdeceb;border-color:#f2c3c1;border-inline-start-color:var(--err);color:#8f1f23;display:block!important}
.info{background:#eaf2fd;border-color:#c4d9f5;border-inline-start-color:var(--info);color:#17427a;display:block!important}
.badge{display:inline-block;padding:3px 10px;border-radius:var(--r-pill);font-size:12px;font-weight:700;
 background:#eef4f4;color:#2f545d}
img{max-width:100%;border-radius:var(--r);border:1px solid var(--line)}
::-webkit-scrollbar{width:11px;height:11px}
::-webkit-scrollbar-thumb{background:#c3d5d3;border-radius:8px;border:3px solid transparent;background-clip:content-box}
::-webkit-scrollbar-thumb:hover{background:#a9c7c4;background-clip:content-box}
@media (prefers-reduced-motion:reduce){
 .flight,.radar-scope .sweep,.radar-scope .dot{animation:none!important}
 .flight{opacity:.35}button:hover{transform:none}}
/* ===== 1) تباين الهيدر: نص أبيض صريح لقراءة كبار السن ===== */
header .sub{color:#fff !important;font-weight:600}
/* ===== 2) الحقول منخفضة الثقة: أيقونة + حد واضح ===== */
.warn{border:1.5px solid #d9a400 !important}
.warn:before{content:"⚠️";position:absolute;left:8px;top:50%;transform:translateY(-50%);font-size:14px}
.field{position:relative}
.field input.warn,.field select.warn{padding-left:32px}
/* ===== 3) الطباعة: إخفاء الملاحة والأزرار ===== */
@media print{
 header .navlinks,.actions .btn,#msg,#go,#save,#cancel,#skip,#read,#scan{display:none !important}
 body{background:#fff}body:before{display:none !important}
 .card{box-shadow:none;border:1px solid #999;break-inside:avoid}
}
/* ===== 4) أرقام LTR داخل RTL: تُلف بـ bdi في الجافاسكربت ===== */
bdi.num{direction:ltr;unicode-bidi:isolate;display:inline-block;font-variant-numeric:tabular-nums}
"""


# Daily incoming workflow assets are kept separate from OCR and persistence.
UI_ROOT = BASE / "ui"
PAGE = (UI_ROOT / "index.html").read_text(encoding="utf-8")
UI_ASSETS = {
    "/assets/postal.css": (UI_ROOT / "postal.css", "text/css; charset=utf-8"),
    "/assets/postal.js": (UI_ROOT / "postal.js", "text/javascript; charset=utf-8"),
}


PAGE_DECISION = r"""<!doctype html>
<html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>تسجيل القرار</title>
<style>
""" + _UI_CSS + r"""
/* ===== خاص بصفحة القرار ===== */
main{padding:clamp(16px,3.6vw,40px);max-width:1280px;margin:auto}
.card{margin-bottom:20px}
header .page-title{position:relative;z-index:2;font-size:clamp(22px,3vw,34px);
 font-weight:800;line-height:1.4}
tr.pick{cursor:pointer;transition:background .15s}
tr.pick:hover{background:#eaf7f4}
tr.pick td:first-child{font-weight:700;color:var(--teal)}
#res{box-shadow:var(--sh-1)}
#res thead th{position:sticky;top:0;z-index:1}
#res td{max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
#edit{border:1px solid var(--line);box-shadow:var(--sh-2)}
#edit h3{margin:0 0 12px;font-size:16px}
#rnum{color:var(--teal)}
#rinfo{color:var(--muted);font-weight:400;font-size:13px}
.flex{display:flex;gap:12px;flex-wrap:wrap}.flex>div{flex:1;min-width:170px}
.search-row{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap}
.search-row>div{flex:1;min-width:240px}
.search-row>div:last-child{flex:0 0 auto;min-width:0}
.checkline{display:flex;align-items:center;gap:8px;margin:14px 0 0;font-size:13.5px;
 font-weight:600;color:#3d5b68}
.checkline input{width:18px;height:18px;accent-color:var(--teal)}
@media(max-width:700px){.flex>div{min-width:100%}}
</style></head><body>
<header><div class="header-deco" aria-hidden="true">
 <div class="radar-scope"><i class="ring"></i><i class="sweep"></i><i class="dot"></i><i class="dot b"></i></div>
 <span class="flight a">✈</span><span class="flight b">✈</span><span class="flight c">✈</span>
</div><span class="page-title">تسجيل رقم القرار</span><nav class="navlinks"><a href="/">رجوع لتسجيل الجواب</a></nav></header>
<main>
<div class="card">
 <div class="sec-title">البحث في السجل <span>دوس على الصف لتسجيل قراره</span></div>
 <div class="search-row">
  <div>
   <label>دوّر: رقم الوارد / الرقم التسلسلي / كلمة في الموضوع أو الجهة</label>
   <input id="q" placeholder="مثال: 1543 أو نظافة">
  </div>
  <div><label>&nbsp;</label><button id="go">دوّر</button></div>
 </div>
 <div class="checkline"><input type="checkbox" id="open"> اعرض اللي مفيش ليه قرار لسه</div>
 <div id="msg"></div>
 <table id="res"><thead><tr>
  <th>صف</th><th>التسلسلي</th><th>الوارد</th><th>تاريخ الوارد</th><th>الجهة</th><th>الموضوع</th>
  <th>الحالة</th><th>رقم القرار</th><th>تاريخ القرار</th></tr></thead><tbody></tbody></table>
</div>
<div class="card" id="edit" style="display:none">
 <h3 style="margin:0 0 6px">تسجيل القرار للصف <span id="rnum"></span> — <span id="rinfo"></span></h3>
 <div class="flex">
  <div style="flex:1"><label>رقم القرار</label><input id="dec_number"></div>
  <div style="flex:1"><label>تاريخ القرار</label><input type="date" id="dec_date"></div>
  <div style="flex:1"><label>الحالة</label><select id="status">
   <option>منفذة</option><option>مؤجل</option><option>مرفوض</option><option>محفوظ</option><option>منتظر</option></select></div>
 </div>
 <button id="save">حفظ القرار</button><button class="sec" id="cancel">إلغاء</button>
</div>
</main>
<script>
const $=id=>document.getElementById(id);
function msg(t,c){const m=$("msg");m.textContent=t;m.className=c||"";}
async function search(){
  msg("ببحث...","info");
  const q=encodeURIComponent($("q").value), open=$("open").checked?1:0;
  const j=await (await fetch("/rows?q="+q+"&open="+open)).json();
  if(j.error){msg(j.error,"err");return;}
  const rows=j.rows||[];
  document.querySelector("#res tbody").innerHTML=rows.map(r=>
   `<tr class="pick" data-row="${r.row}" data-info="${encodeURIComponent(r.serial+" — "+r.sender)}">`+
   `<td>${r.row}</td><td><bdi class="num">${r.serial}</bdi></td><td><bdi class="num">${r.number}</bdi></td><td><bdi class="num">${r.wared_date}</bdi></td>`+
   `<td>${r.sender}</td><td>${r.subject}</td><td>${r.status}</td><td><bdi class="num">${r.dec_number}</bdi></td><td><bdi class="num">${r.dec_date}</bdi></td></tr>`
  ).join("")||'<tr><td colspan="9">مفيش نتائج</td></tr>';
  msg(rows.length?rows.length+" نتيجة — دوس على الصف عشان تسجل قراره":"مفيش نتائج","info");
}
let picked=null;
document.addEventListener("click",e=>{
  const tr=e.target.closest("tr.pick"); if(!tr)return;
  picked=tr.dataset.row;
  $("rnum").textContent=picked;
  $("rinfo").textContent=decodeURIComponent(tr.dataset.info||"");
  $("dec_number").value=tr.children[7].textContent;
  $("dec_date").value=tr.children[8].textContent;
  $("status").value=tr.children[6].textContent||"منفذة";
  $("edit").style.display="block";$("edit").scrollIntoView();
});
$("go").onclick=search;
$("q").onkeydown=e=>{if(e.key==="Enter")search();};
let _qt=null;
$("q").oninput=()=>{clearTimeout(_qt);_qt=setTimeout(()=>{if($("q").value.trim().length>=2)search();},500);};
$("open").onchange=search;
$("cancel").onclick=()=>$("edit").style.display="none";
$("save").onclick=async()=>{
  if(!picked)return;
  const btn=$("save"); btn.disabled=true;
  try{
    const body=JSON.stringify({row:picked,dec_number:$("dec_number").value,
      dec_date:$("dec_date").value,status:$("status").value});
    let j=null;
    for(let k=0;k<3;k++){
      const r=await fetch("/decision",{method:"POST",headers:{"Content-Type":"application/json"},body});
      j=await r.json();
      if(!j.error)break;
      if(r.status!==409)break;
      msg("السجل مشغول لحظيًا... بحاول تاني ("+(k+1)+"/3)","info");
      await new Promise(res=>setTimeout(res,400*(k+1)));
    }
    if(j.error){msg(j.error,"err");return;}
    msg("اتسجل القرار في الصف "+j.row,"ok");$("edit").style.display="none";search();
  }finally{btn.disabled=false;}
};
search();
</script></body></html>
"""


def _err(h, code, text):
    body = json.dumps({"error": text}, ensure_ascii=False).encode("utf-8")
    h.send_response(code)
    h.send_header("Content-Type", "application/json; charset=utf-8")
    h.send_header("Content-Length", str(len(body)))
    h.end_headers()
    h.wfile.write(body)


def _json_safe(o):
    if isinstance(o, dict):
        return {k: _json_safe(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_json_safe(v) for v in o]
    if isinstance(o, (datetime, date)):
        return o.isoformat()
    return o


def _json(h, obj):
    body = json.dumps(_json_safe(obj), ensure_ascii=False).encode("utf-8")
    h.send_response(200)
    h.send_header("Content-Type", "application/json; charset=utf-8")
    h.send_header("Content-Length", str(len(body)))
    h.end_headers()
    h.wfile.write(body)


def _qs(path):
    u = urllib.parse.urlparse(path)
    return u.path, {k: v[0] for k, v in urllib.parse.parse_qs(u.query).items()}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _page(self, html):
        body = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path in UI_ASSETS:
            file, content_type = UI_ASSETS[self.path]
            data = file.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()
            self.wfile.write(data)
            return
        if self.path == "/":
            return self._page(PAGE)
        if self.path == "/decision":
            return self._page(PAGE_DECISION)
        if self.path == "/meta":
            try:
                _json(self, read_meta())
            except Exception as e:
                _err(self, 500, f"مش قادر أفتح ملف الإكسل: {e}")
        elif self.path.startswith("/inbox"):
            try:
                _json(self, {"files": [p.name for p in sorted(INBOX.iterdir())
                                       if p.suffix.lower() in EXTS]})
            except Exception as e:
                _err(self, 500, f"مفيش فولدر inbox: {e}")
        elif self.path.startswith("/rows"):
            try:
                _, args = _qs(self.path)
                _json(self, {"rows": find_rows(args.get("q", ""), args.get("open") == "1")})
            except Exception as e:
                _err(self, 500, f"فشل البحث: {e}")
        elif self.path.startswith("/setup"):
            try:
                done, sheet = ensure_workbook(apply=True)
                _json(self, {"done": done, "sheet": sheet})
            except PermissionError:
                _err(self, 409, "اقفل ملف الإكسل في إكسل وبعدين جرّب تاني.")
            except Exception as e:
                _err(self, 500, f"فشل ضبط السجل: {e}")
        elif self.path.startswith("/scanners"):
            try:
                _json(self, wia_list())
            except Exception as e:
                _err(self, 500, f"مش قادر أشوف الأجهزة: {e}")
        elif self.path == "/logo":
            if not LOGO.exists():
                return _err(self, 404, "شعار الشركة غير موجود")
            data = LOGO.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "image/png")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        elif self.path.startswith("/img/"):
            parts = self.path[5:].split("?")[0]
            # صورة المعاينة المعالَجة هي نفس الصورة التي خرجت منها إحداثيات OCR؛
            # لذلك يظل تحديد الدليل دقيقًا عند عرض أي صفحة من PDF.
            page_match = re.fullmatch(r"([0-9a-f]+)/page/(\d+)", parts)
            if page_match:
                uid, page_no = page_match.groups()
                file = TMP / uid / f"prep{int(page_no) - 1}.png"
                if not file.is_file():
                    return _err(self, 404, "مفيش صفحة بالرقم ده")
                data = file.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            if "/" in parts:                       # /img/<uid>/crop/<name> (قديم، للتوافق فقط)
                return _err(self, 404, "مفيش صورة")
            uid = re.sub(r"[^0-9a-f]", "", parts)
            files = list((TMP / uid).glob("view.*"))
            if not files:
                return _err(self, 404, "مفيش صورة")
            data = files[0].read_bytes()
            ctype = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                     ".tif": "image/tiff", ".tiff": "image/tiff"}.get(files[0].suffix.lower(), "image/png")
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        else:
            _err(self, 404, "غير موجود")

    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(n)
        if self.path == "/ocr":
            return self.handle_ocr(raw)
        if self.path == "/inbox":
            return self.handle_inbox(raw)
        if self.path == "/scan":
            return self.handle_scan(raw)
        if self.path == "/save":
            return self.handle_save(raw)
        if self.path == "/decision":
            return self.handle_decision(raw)
        _err(self, 404, "غير موجود")

    def _ocr_payload(self, raw: bytes, name: str, source_kind="upload"):
        ext = Path(name).suffix.lower()
        if ext not in EXTS:
            raise ValueError("نوع الملف مش مدعوم. استخدم PDF أو صورة.")
        uid = uuid.uuid4().hex
        d = TMP / uid
        d.mkdir(parents=True)
        src = d / ("orig" + ext)
        src.write_bytes(raw)
        if ext == ".pdf":
            try:
                pages = pdf_pages(src, d)
            except ImportError:
                raise ValueError("محتاج تنصّب PyMuPDF:  pip install pymupdf")
            if not pages:
                raise ValueError("الـ PDF فاضي أو مقفول.")
            shutil.copyfile(pages[0], d / "view.png")
        else:
            shutil.copy(src, d / ("view" + ext))
            pages = image_pages(src, d)
        prepped = []
        page_info = []
        for i, p in enumerate(pages):
            pp = d / f"prep{i}.png"
            prepare_image(p, pp)
            prepped.append(pp)
            quality = document_quality(pp)
            page_info.append({
                "index": i + 1,
                "width": quality.get("width"),
                "height": quality.get("height"),
                "quality_warnings": quality.get("warnings", []),
                "preview_url": f"/img/{uid}/page/{i + 1}",
            })
        engine = document_ai_engine()
        text, warnings, all_lines = engine.analyze(prepped)
        flat = [l for page in all_lines for l in page]
        try:
            jehat = read_meta()["jehat"]
        except Exception:
            jehat = []
        ctx = {"prepped": prepped[0] if prepped else None, "workdir": d,
               "prepared_pages": prepped}
        # للـ QR: نجرب الصورة المجهّزة وبعدها الأصلية/كل صفحات الـ PDF
        fields = extract_fields(text, jehat, flat, ctx,
                                orig_path=prepped[0] if prepped else None,
                                qr_paths=list(pages))
        quality_warnings = [warning for page in page_info
                            for warning in page["quality_warnings"]]
        # لا قصاصات محفوظة: القص كله مؤقت داخل الذاكرة أثناء OCR فقط.
        # مجلد الورقة النهائي = الصورة الأصلية + تقرير Word، والإكسل مرجع.
        payload = {"id": uid, "text": text, "fields": fields,
                "warnings": warnings + list(dict.fromkeys(quality_warnings)),
                "crops": [],
                "document": {
                    "page_count": len(page_info),
                    "pages": page_info,
                    "engine": engine.name,
                    "mode": engine.mode,
                }}
        (d / "result.json").write_text(json.dumps(
            {"source_kind": source_kind, "text": text, "fields": fields},
            ensure_ascii=False), encoding="utf-8")
        return payload

    def handle_ocr(self, raw):
        name = urllib.parse.unquote(self.headers.get("X-Filename", "scan"))
        try:
            _json(self, self._ocr_payload(raw, name))
        except ValueError as e:
            _err(self, 400, str(e))
        except Exception as e:
            _err(self, 500, f"حصل خطأ أثناء القراءة: {e}")

    def handle_inbox(self, raw):
        try:
            name = Path(str(json.loads(raw.decode("utf-8")).get("name", ""))).name
        except Exception:
            return _err(self, 400, "طلب غلط")
        f = INBOX / name
        if not name or not f.exists():
            return _err(self, 404, f"الملف {name} مش موجود في فولدر inbox")
        try:
            payload = self._ocr_payload(f.read_bytes(), name, "scanner")
            payload["inbox"] = name
            _json(self, payload)
        except ValueError as e:
            _err(self, 400, str(e))
        except Exception as e:
            _err(self, 500, f"حصل خطأ أثناء القراءة: {e}")

    def handle_save(self, raw):
        try:
            data = json.loads(raw.decode("utf-8"))
            uid = str(data.get("id") or "")
            if not re.fullmatch(r"[0-9a-f]{32}", uid):
                raise ValueError("اقرأ الجواب أولًا قبل الحفظ.")
            result_file = TMP / uid / "result.json"
            if not result_file.is_file():
                raise ValueError("جلسة قراءة الجواب غير موجودة؛ اقرأ الملف مرة أخرى.")
            result = json.loads(result_file.read_text(encoding="utf-8"))
            fields = data.get("fields")
            if not isinstance(fields, dict):
                raise ValueError("بيانات الجواب غير صالحة.")
            fields = {key: str(fields.get(key) or "").strip()
                      for key in ("number", "wared_date", "subject", "sender", "recipient")}
            fields['recipient'] = recipient_job_title(fields.get('recipient'))
            for key, label in (("subject", "نبذة الموضوع"),):
                if not str(fields.get(key) or "").strip():
                    raise ValueError(f"راجع الحقل الناقص: {label}.")
            fields["number"] = re.sub(r"\s+", "", fields["number"]).translate(AR_DIGITS)
            if fields["number"] and not re.fullmatch(r"\d+(?:[/-]\d+)*", fields["number"]):
                raise ValueError("رقم الوارد غير صالح؛ طابق الأرقام مع الختم.")
            if fields["wared_date"] and not _d(str(fields["wared_date"])):
                raise ValueError("تاريخ الوارد غير صالح.")
            extracted = result.get("fields", {})
            reviewed = data.get("reviewed", [])
            for key in ("number", "wared_date"):
                if (fields.get(key) and extracted.get("_confidence", {}).get(key) == "low"
                        and fields.get(key) == extracted.get(key) and key not in reviewed):
                    raise ValueError("طابق رقم الوارد وتاريخه مع الختم ثم أكد المراجعة.")
            fields["_confidence"] = {
                key: extracted.get("_confidence", {}).get(key, "medium")
                if fields[key] == extracted.get(key) else "reviewed"
                for key in fields if key != '_confidence'}
            origs = list((TMP / uid).glob("orig.*")) if uid else []
            source = origs[0] if origs else None
            if result["source_kind"] == "scanner" and source is None:
                raise ValueError("صورة المسح الأصلية غير موجودة؛ أعد قراءة الجواب.")
            if not result.get("text", "").strip():
                raise ValueError("لم يُستخرج نص من الجواب؛ أعد المسح بصورة أوضح.")
            serial, row = save_row(fields, source, result["text"],
                                   archive_source=result["source_kind"] == "scanner",
                                   request_id=uid)
            _json(self, {"serial": serial, "row": row})
        except PermissionError:
            _err(self, 409, "اقفل ملف الإكسل الأول وبعدين اضغط حفظ تاني.")
        except ValueError as e:
            _err(self, 400, str(e))
        except Exception as e:
            _err(self, 500, f"حصل خطأ في الحفظ: {e}")

    def handle_scan(self, raw):
        try:
            d = json.loads(raw.decode("utf-8")) if raw else {}
        except Exception:
            d = {}
        try:
            path = wia_scan(device=d.get("device", -1), dpi=d.get("dpi", 300),
                            color=d.get("color", "gray"),
                            source=d.get("source", "auto"))
        except ValueError as e:
            return _err(self, 400, str(e))
        except Exception as e:
            return _err(self, 500, f"حصل خطأ أثناء المسح: {e}")
        try:
            payload = self._ocr_payload(path.read_bytes(), path.name, "scanner")
            payload["scan"] = path.name
            _json(self, payload)
        except ValueError as e:
            _err(self, 400, str(e))
        except Exception as e:
            _err(self, 500, f"حصل خطأ أثناء القراءة: {e}")

    def handle_decision(self, raw):
        try:
            d = json.loads(raw.decode("utf-8"))
            info = save_decision(
                d.get("row", 0),
                dec_number=d.get("dec_number", ""),
                dec_date=d.get("dec_date", ""),
                status=d.get("status", "منفذة"))
            _json(self, {"row": info["row"], "serial": info["serial"]})
        except PermissionError:
            _err(self, 409, "اقفل ملف الإكسل الأول وبعدين احفظ تاني.")
        except (ValueError, TypeError) as e:
            _err(self, 400, str(e))
        except Exception as e:
            _err(self, 500, f"حصل خطأ في الحفظ: {e}")


def main():
    if not XLSX.exists():
        print(f"ملف الإكسل مش موجود: {XLSX}\nحطه في نفس فولدر الأداة.")
        sys.exit(1)
    try:
        import docx, openpyxl, PIL, fitz, rapidocr  # noqa: F401
    except ImportError as e:
        print(f"ناقص مكتبة ({getattr(e, 'name', e)}):\n"
              "  pip install -r requirements.txt")
        sys.exit(1)
    # تهيئة أول مرة: أعمدة ناقصة + ورقة القوائم (بتعمل نسخة احتياطية لوحدها)
    try:
        done, _sheet = ensure_workbook(apply=True)
        for d in done:
            print("تهيئة:", d)
    except Exception as e:
        print("تعذر ضبط الملف تلقائيًا:", e)
        print("اقفل ملف Excel لو مفتوح، ثم أعد تشغيل الأداة لإكمال تجهيز السجل.")
    if AR_MODEL.exists() and AR_DICT.exists():
        print("نموذج العربية: PP-OCRv5 (جاهز)")
    else:
        print("تنبيه: نموذج العربية مش موجود في فولدر models — شغّل models_dl.py.")
    print("فولدر المسح الجاهز:", INBOX)
    # الخادم يبدأ الأول (علشان البورت يبقى مفتوح) وبعدين تجهيز المحرك في الخلفية
    try:
        srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    except OSError:
        # شغّالة أصلًا (تانية) -> نفتح المتصفح ونتوقف
        try:
            webbrowser.open(f"http://127.0.0.1:{PORT}")
        except Exception:
            pass
        return
    threading.Thread(target=ocr_engine, daemon=True).start()
    url = f"http://127.0.0.1:{PORT}"
    print("الأداة شغالة على", url, "(اقفل النافذة دي لإيقافها)")
    if not os.environ.get("POSTAL_NO_BROWSER"):
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()

# ------------------------------------------------------------------ Scaling utilities
# Added scaling constants and helper functions as requested.
# DEFAULT_SCALE is used as the base scaling factor for digit cropping operations.
# It can be overridden per-call if needed.
DEFAULT_SCALE = 5

def _optimal_scale(image_path: Path, max_side: int = 1600) -> int:
    """Calculate a scaling factor to ensure the larger dimension of the image
    does not exceed ``max_side`` pixels. This helps keep OCR token consumption
    low while preserving readability.

    The function returns an integer ``scale`` that can be passed to the
    ``crop_digits_px`` helper. A ``scale`` of ``1`` means no up‑scaling.
    """
    from PIL import Image
    try:
        with Image.open(image_path) as im:
            w, h = im.size
    except Exception:
        return 1
    larger = max(w, h)
    if larger <= max_side:
        return 1
    # Compute the smallest integer scaling factor that brings ``larger`` down
    # to <= ``max_side``. We use integer division rounding up.
    factor = (larger + max_side - 1) // max_side
    # Ensure factor is at least 1.
    return max(1, factor)
