"""
iniciar_api_web.py — Inicia la API Movil de O-Leasing en el puerto 8008 y abre
un tunel HTTPS publico (Cloudflare Quick Tunnel) para que la aplicacion movil
se conecte por medio web desde cualquier red (Wi-Fi o datos moviles 4G/5G)
sin depender de una direccion IP local cambiante.
"""
import os
import re
import socket
import subprocess
import sys
import time
import threading

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TUNNEL_URL_FILE = os.path.join(BASE_DIR, "tunnel_url.txt")
QR_CONEXION_FILE = os.path.join(BASE_DIR, "qr_conexion_movil.png")
PORT = 8008

CLOUDFLARED_PATHS = [
    r"C:\Program Files (x86)\cloudflared\cloudflared.exe",
    r"C:\Program Files\cloudflared\cloudflared.exe",
    "cloudflared",
]


def is_port_in_use(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0


def find_cloudflared() -> str:
    for p in CLOUDFLARED_PATHS:
        if os.path.exists(p):
            return p
    return "cloudflared"


def save_connection_qr(public_url: str):
    try:
        import qrcode
        qr = qrcode.QRCode(version=1, box_size=6, border=2)
        qr.add_data(f"oleasing://conectar?api={public_url}")
        qr.make(fit=True)
        img = qr.make_image(fill_color="#E65C00", back_color="white")
        img.save(QR_CONEXION_FILE)
    except Exception:
        pass


def main():
    print("=" * 72)
    print("  O-LEASING MOBILE — SERVIDOR API + TUNEL WEB HTTPS AUTOMATICO")
    print("=" * 72)

    api_proc = None
    if is_port_in_use(PORT):
        print(f"[OK] La API de O-Leasing ya esta activa en http://127.0.0.1:{PORT}")
    else:
        print(f"[*] Iniciando API de O-Leasing en el puerto {PORT}...")
        api_proc = subprocess.Popen(
            [sys.executable, "-m", "uvicorn", "api:app", "--host", "0.0.0.0", "--port", str(PORT), "--reload"],
            cwd=BASE_DIR,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        for _ in range(20):
            if is_port_in_use(PORT):
                break
            time.sleep(0.5)
        print(f"[OK] API iniciada correctamente en http://127.0.0.1:{PORT}")

    cf_bin = find_cloudflared()
    print("[*] Creando enlace web seguro HTTPS (Cloudflare Tunnel)...")
    cf_proc = subprocess.Popen(
        [cf_bin, "tunnel", "--url", f"http://127.0.0.1:{PORT}"],
        cwd=BASE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    url_pattern = re.compile(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com")
    public_url = None

    try:
        for line in cf_proc.stdout:
            m = url_pattern.search(line)
            if m and not public_url:
                public_url = m.group(0)
                with open(TUNNEL_URL_FILE, "w", encoding="utf-8") as f:
                    f.write(public_url)
                save_connection_qr(public_url)
                print("\n" + "=" * 72)
                print("  CONEXION WEB ACTIVA (SIN IP LOCAL)")
                print("=" * 72)
                print(f"  URL Web para la App Movil : {public_url}")
                print(f"  Guardada automaticamente  : {TUNNEL_URL_FILE}")
                print(f"  QR de Conexion rapida     : {QR_CONEXION_FILE}")
                print("-" * 72)
                print("  Como funciona:")
                print("  1. Todos los codigos QR de Estado de Cuenta en la App Web y PDFs")
                print("     ahora incluyen automaticamente esta URL web.")
                print("  2. Al escanear cualquier QR de un contrato con la camara de tu")
                print("     telefono, la app movil se conectara sola por internet y")
                print("     abrira directamente el Estado de Cuenta de ese cliente.")
                print("=" * 72 + "\n")
    except KeyboardInterrupt:
        print("\n[*] Cerrando tunel web y servidor API...")
    finally:
        cf_proc.terminate()
        if api_proc:
            api_proc.terminate()


if __name__ == "__main__":
    main()
