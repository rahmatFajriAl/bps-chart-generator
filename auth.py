"""Login satu akun (admin). Kredensial awal dari .streamlit/secrets.toml atau environment variable.
Kata sandi bisa diganti dari dalam aplikasi; hash baru disimpan di database dan menggantikan hash di Secrets."""
import hmac
import math
import os
import time

import streamlit as st

import db
import security

IDLE_TIMEOUT = 30 * 60      # logout otomatis setelah 30 menit tidak aktif
MAX_FAILS = 5               # 5x salah -> terkunci
LOCK_SECONDS = 15 * 60      # selama 15 menit
MAX_INPUT_LEN = 128
MIN_PW_LEN = 10


def _secret_credentials():
    user = pw_hash = None
    ignore = False
    try:
        sec = st.secrets.get("auth", {})
        user, pw_hash = sec.get("username"), sec.get("password_hash")
        ignore = bool(sec.get("ignore_db_password", False))
    except Exception:
        pass
    return (user or os.environ.get("ADMIN_USERNAME"),
            pw_hash or os.environ.get("ADMIN_PASSWORD_HASH"),
            ignore or os.environ.get("IGNORE_DB_PASSWORD") == "1")


def _credentials():
    """(username, hash aktif). Hash di database (kalau ada) menang atas hash di Secrets."""
    user, secret_hash, ignore = _secret_credentials()
    override = None if ignore else db.get_password_override()[0]
    return user, (override or secret_hash)


def _current_stamp():
    _, _, ignore = _secret_credentials()
    return "" if ignore else db.get_credential_stamp()


def _log(status):
    try:
        db.log_login(status)
    except Exception:
        pass    # pencatatan tidak boleh menggagalkan login


def is_configured():
    u, h = _credentials()
    return bool(u and h)


def is_authenticated():
    if not st.session_state.get("auth_ok"):
        return False
    if time.time() - st.session_state.get("auth_last", 0) > IDLE_TIMEOUT:
        logout("Sesi berakhir karena tidak aktif. Silakan masuk lagi.")
        return False
    if st.session_state.get("auth_stamp", "") != _current_stamp():
        logout("Kata sandi telah diubah. Silakan masuk lagi.")
        return False
    st.session_state["auth_last"] = time.time()
    return True


def current_user():
    return st.session_state.get("auth_user", "")


def login(username, password):
    """Return (ok, pesan)."""
    cfg_user, cfg_hash = _credentials()
    if not cfg_user or not cfg_hash:
        return False, "Akun admin belum dikonfigurasi (lihat petunjuk setup)."

    remaining = db.lock_remaining()
    if remaining > 0:
        _log("locked")
        return False, f"Terlalu banyak percobaan gagal. Coba lagi dalam {math.ceil(remaining / 60)} menit."

    username, password = username.strip(), password
    too_long = len(username) > MAX_INPUT_LEN or len(password) > MAX_INPUT_LEN
    user_ok = hmac.compare_digest(username.encode(), cfg_user.encode())
    pw_ok = security.verify_password("" if too_long else password, cfg_hash)  # selalu dihitung

    if user_ok and pw_ok and not too_long:
        db.reset_failures()
        st.session_state["auth_ok"] = True
        st.session_state["auth_user"] = cfg_user
        st.session_state["auth_last"] = time.time()
        st.session_state["auth_stamp"] = _current_stamp()
        _log("success")
        return True, ""

    db.register_failure(MAX_FAILS, LOCK_SECONDS)
    _log("fail")
    time.sleep(1)  # perlambat brute force
    return False, "Username atau password salah."


def change_password(current, new, confirm):
    """Ganti kata sandi admin dari dalam aplikasi. Return (ok, pesan)."""
    if not st.session_state.get("auth_ok"):
        return False, "Sesi tidak valid. Silakan masuk lagi."
    cfg_user, cfg_hash = _credentials()
    if _secret_credentials()[2]:
        return False, "Ganti kata sandi dinonaktifkan (ignore_db_password aktif di Secrets)."
    if not cfg_hash:
        return False, "Akun admin belum dikonfigurasi."

    remaining = db.lock_remaining()
    if remaining > 0:
        return False, f"Terlalu banyak percobaan gagal. Coba lagi dalam {math.ceil(remaining / 60)} menit."

    # pemeriksaan murah dulu (tidak dihitung sebagai percobaan gagal)
    if not current or not new:
        return False, "Semua kolom wajib diisi."
    if max(len(current), len(new), len(confirm)) > MAX_INPUT_LEN:
        return False, f"Maksimal {MAX_INPUT_LEN} karakter."
    if new != confirm:
        return False, "Konfirmasi kata sandi baru tidak sama."
    if len(new) < MIN_PW_LEN:
        return False, f"Kata sandi baru minimal {MIN_PW_LEN} karakter."
    if new.lower() == (cfg_user or "").lower():
        return False, "Kata sandi tidak boleh sama dengan username."
    if hmac.compare_digest(new.encode(), current.encode()):
        return False, "Kata sandi baru harus berbeda dari yang lama."

    if not security.verify_password(current, cfg_hash):
        db.register_failure(MAX_FAILS, LOCK_SECONDS)   # sama seperti login: cegah tebak-tebakan
        _log("pw_fail")
        time.sleep(1)
        return False, "Kata sandi saat ini salah."

    db.set_password_override(security.hash_password(new))
    st.session_state["auth_stamp"] = _current_stamp()   # sesi ini tetap masuk, sesi lain otomatis keluar
    db.reset_failures()
    _log("pw_changed")
    return True, "Kata sandi berhasil diubah. Sesi login di perangkat lain otomatis keluar."


def logout(message=None):
    st.session_state.clear()   # bersihkan juga file yang sedang terbuka di memori sesi
    if message:
        st.session_state["auth_msg"] = message
