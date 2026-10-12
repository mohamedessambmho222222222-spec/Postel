# -*- coding: utf-8 -*-
"""تحميل نموذج العربية (PP-OCRv5) + قاموس الحروف من Mirror رسمي (ModelScope).

بيستخدم urllib الأول، ولو الشبكة رفضته (Proxy) بيجرّب PowerShell/Invoke-WebRequest.
"""
import hashlib
import subprocess
import sys
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
MODELS = HERE / "models"
MODELS.mkdir(exist_ok=True)

BASE = "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.9.2"
FILES = [
    (MODELS / "arabic_PP-OCRv5_rec_mobile.onnx",
     f"{BASE}/onnx/PP-OCRv5/rec/arabic_PP-OCRv5_rec_mobile.onnx"),
    (MODELS / "ppocrv5_arabic_dict.txt",
     f"{BASE}/paddle/PP-OCRv5/rec/arabic_PP-OCRv5_rec_mobile/ppocrv5_arabic_dict.txt"),
]


def via_powershell(url: str, dst: Path) -> bool:
    ps = ("$ProgressPreference='SilentlyContinue';"
          f"Invoke-WebRequest -Uri '{url}' -OutFile '{dst}' "
          "-UseBasicParsing -TimeoutSec 900")
    try:
        subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-Command", ps], timeout=1200)
    except Exception as e:
        print("   powershell failed:", e)
        return False
    return dst.exists() and dst.stat().st_size > 1000


def via_urllib(url: str, dst: Path) -> bool:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=600) as r, open(dst, "wb") as f:
            f.write(r.read())
        return dst.exists() and dst.stat().st_size > 1000
    except Exception as e:
        print("   urllib failed:", str(e)[:90])
        return False


def main():
    ok = True
    for dst, url in FILES:
        if dst.exists() and dst.stat().st_size > 1000:
            print(f"موجود بالفعل: {dst.name} ({dst.stat().st_size} bytes)")
            continue
        print("بتحميل", dst.name)
        if not (via_urllib(url, dst) or via_powershell(url, dst)):
            ok = False
            print("  فشل التحميل:", url)
            continue
        sha = hashlib.sha256(dst.read_bytes()).hexdigest()
        print(f"  تم: {dst.stat().st_size} bytes  sha256={sha[:16]}")
    if not ok:
        print("\nملاحظة: الشبكة ساعات بتحجب ModelScope/GitHub. الملفات مطلوبة مرة واحدة بس.")
        sys.exit(1)
    print("جاهز — النموذج العربي في فولدر models")


if __name__ == "__main__":
    main()

