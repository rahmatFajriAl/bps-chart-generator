"""Login satu akun (admin). Kredensial dari .streamlit/secrets.toml atau environment variable."""
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


def _credentials():
    user = pw_hash = None
    try:
        sec = st.secrets.get("auth", {})
        user, pw_hash = sec.get("username"), sec.get("password_hash")
    except Exception:
        pass
    return user or os.environ.get("ADMIN_USERNAME"), pw_hash or os.environ.get("ADMIN_PASSWORD_HASH")


def is_configured():
    u, h = _credentials()
    return bool(u and h)


def is_authenticated():
    if not st.session_state.get("auth_ok"):
        return False
    if time.time() - st.session_state.get("auth_last", 0) > IDLE_TIMEOUT:
        logout("Sesi berakhir karena tidak aktif. Silakan masuk lagi.")
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
        return True, ""

    db.register_failure(MAX_FAILS, LOCK_SECONDS)
    time.sleep(1)  # perlambat brute force
    return False, "Username atau password salah."


def logout(message=None):
    st.session_state.clear()   # bersihkan juga file yang sedang terbuka di memori sesi
    if message:
        st.session_state["auth_msg"] = message