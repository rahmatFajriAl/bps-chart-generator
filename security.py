"""Fungsi keamanan murni (tanpa Streamlit) supaya bisa dipakai app.py dan make_hash.py."""
import base64
import hashlib
import hmac
import io
import os
import re
import zipfile
from pathlib import Path

ALLOWED_EXT = {"xlsx", "xlsm", "csv"}
MAX_UPLOAD_BYTES = 10 * 1024 * 1024          # 10 MB per file
MAX_UNZIPPED_BYTES = 100 * 1024 * 1024       # batas ukuran setelah di-ekstrak (anti zip-bomb)
MAX_ZIP_ENTRIES = 2000

_CTRL_RE = re.compile(r"[\x00-\x1f\x7f]")
_MD_SPECIAL = re.compile(r"([\\`*_{}\[\]()#+\-.!|<>~$:])")


# ------------------------------------------------------------ password ----
def _b64(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode()


def _unb64(s: str) -> bytes:
    return base64.urlsafe_b64decode(s.encode())


def _scrypt(password: str, salt: bytes, n: int, r: int, p: int) -> bytes:
    return hashlib.scrypt(password.encode("utf-8"), salt=salt, n=n, r=r, p=p,
                          dklen=32, maxmem=128 * 1024 * 1024)


def hash_password(password: str, n: int = 2 ** 14, r: int = 8, p: int = 1) -> str:
    salt = os.urandom(16)
    return f"scrypt${n}${r}${p}${_b64(salt)}${_b64(_scrypt(password, salt, n, r, p))}"


def verify_password(password: str, stored: str) -> bool:
    """Selalu menghitung hash (walau format stored rusak) supaya waktu responsnya seragam."""
    try:
        scheme, n, r, p, salt, dk = stored.split("$")
        if scheme != "scrypt":
            raise ValueError
        expected = _unb64(dk)
        calc = _scrypt(password, _unb64(salt), int(n), int(r), int(p))
        return hmac.compare_digest(calc, expected)
    except Exception:
        _scrypt(password, b"0" * 16, 2 ** 14, 8, 1)
        return False


# ------------------------------------------------------------ sanitizers ----
def safe_name(s, maxlen: int = 120) -> str:
    """Buang karakter kontrol, rapikan spasi, batasi panjang."""
    s = _CTRL_RE.sub("", str(s or ""))
    s = re.sub(r"\s+", " ", s).strip()
    return s[:maxlen]


def safe_filename(name, ext: str) -> str:
    base = re.sub(r"[^\w\-. ()]", "_", safe_name(name), flags=re.UNICODE).strip(" .") or "tabel"
    if not base.lower().endswith("." + ext):
        base += "." + ext
    return base


def md_escape(s) -> str:
    """Escape karakter Markdown supaya teks dari pengguna tampil apa adanya."""
    return _MD_SPECIAL.sub(r"\\\1", str(s))


# ------------------------------------------------------------ upload ----
def validate_upload(filename: str, data: bytes):
    """Return None kalau aman, atau string pesan error."""
    ext = Path(filename).suffix.lower().lstrip(".")
    if ext not in ALLOWED_EXT:
        return "Format file tidak diizinkan. Gunakan .xlsx, .xlsm, atau .csv."
    if not data:
        return "File kosong."
    if len(data) > MAX_UPLOAD_BYTES:
        return f"Ukuran file melebihi batas {MAX_UPLOAD_BYTES // (1024 * 1024)} MB."

    if ext == "csv":
        if b"\x00" in data[:8192]:
            return "File CSV tidak valid (berisi data biner)."
        return None

    if not data.startswith(b"PK"):
        return "File Excel tidak valid (bukan format .xlsx yang sebenarnya)."
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            infos = zf.infolist()
            names = [i.filename for i in infos]
            if len(infos) > MAX_ZIP_ENTRIES:
                return "Struktur file Excel tidak wajar."
            if "[Content_Types].xml" not in names or not any(n.startswith("xl/") for n in names):
                return "File Excel tidak valid."
            if any(n.startswith(("/", "\\")) or ".." in n.replace("\\", "/").split("/") for n in names):
                return "File Excel berisi path yang tidak diizinkan."
            if any(n.lower().endswith("vbaproject.bin") for n in names):
                return "File yang berisi macro tidak diizinkan. Simpan ulang sebagai .xlsx biasa."
            if sum(i.file_size for i in infos) > MAX_UNZIPPED_BYTES:
                return "Isi file terlalu besar setelah diekstrak."
    except zipfile.BadZipFile:
        return "File Excel rusak atau bukan format yang benar."
    return None