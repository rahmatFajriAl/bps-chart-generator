"""Buat hash password untuk secrets.toml.  Jalankan:  python make_hash.py"""
import getpass
import security

pw = getpass.getpass("Password baru: ")
if len(pw) < 10:
    raise SystemExit("Minimal 10 karakter.")
if pw != getpass.getpass("Ulangi password: "):
    raise SystemExit("Password tidak sama.")
print("\nTempel ke .streamlit/secrets.toml:\n")
print("[auth]")
print('username = "admin"')
print(f'password_hash = "{security.hash_password(pw)}"')