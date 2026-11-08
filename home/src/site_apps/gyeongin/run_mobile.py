"""
경인사무소 현장 모바일 검사 입력 앱 실행기
스마트폰 / 태블릿 / PC 어디서나 브라우저로 접속 가능합니다.
"""
import re
import socket
import subprocess
import sys
import threading
import time
from pathlib import Path

# Add home/src to sys.path
HOME_SRC = Path(__file__).resolve().parents[2]
if str(HOME_SRC) not in sys.path:
    sys.path.insert(0, str(HOME_SRC))

PROJECT_ROOT = Path(__file__).resolve().parents[4]
CLOUDFLARED_EXE = PROJECT_ROOT / "tools" / "cloudflared.exe"

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def start_tunnel(port):
    """Cloudflare Tunnel을 실행하여 외부 접속용 공개 HTTPS 주소를 획득합니다."""
    if not CLOUDFLARED_EXE.exists():
        print(f" [경고] {CLOUDFLARED_EXE} 파일을 찾을 수 없어 로컬 모드로만 실행됩니다.", flush=True)
        return None, None

    cmd = [str(CLOUDFLARED_EXE), "tunnel", "--url", f"http://127.0.0.1:{port}"]
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        encoding="utf-8",
        errors="replace"
    )

    public_url = None
    url_pattern = re.compile(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com")

    start_time = time.time()
    for line in iter(proc.stdout.readline, ''):
        match = url_pattern.search(line)
        if match:
            public_url = match.group(0)
            break
        if time.time() - start_time > 15:
            break

    def drain_pipe():
        try:
            for _ in iter(proc.stdout.readline, ''):
                pass
        except Exception:
            pass

    t = threading.Thread(target=drain_pipe, daemon=True)
    t.start()

    return proc, public_url

def main():
    import uvicorn
    local_ip = get_local_ip()
    port = 8001  # Gyeongin uses port 8001 to run side-by-side with Central (8000)

    print(" [안내] 경인사무소 외부 접속 터널을 시작하는 중입니다...", flush=True)
    tunnel_proc, public_url = start_tunnel(port)

    print("=" * 65, flush=True)
    print("   🚀 한국지역난방공사 경인사무소 모바일 현장 입력 시스템", flush=True)
    print("=" * 65, flush=True)
    print(f" [사내/PC 브라우저 접속]        👉 http://localhost:{port}", flush=True)
    print(f" [사내 Wi-Fi 스마트폰/태블릿]   👉 http://{local_ip}:{port}", flush=True)
    if public_url:
        print(f" [🌟 외부 현장/LTE·5G 어디서나] 👉 {public_url}", flush=True)
    print("=" * 65, flush=True)
    print(" * 현장 스마트폰(LTE/5G)에서 위 [외부 현장] 주소로 접속하시면", flush=True)
    print("   경인사무소 일보를 실시간으로 등록하실 수 있습니다.", flush=True)
    print(" * 스마트폰 브라우저에서 '홈 화면에 추가'를 누르면 앱처럼 실행됩니다.", flush=True)
    print("=" * 65, flush=True)
    print(" 서버를 종료하려면 터미널에서 Ctrl + C를 누르세요.\n", flush=True)

    try:
        uvicorn.run("site_apps.gyeongin.mobile.server:app", host="0.0.0.0", port=port, reload=False)
    finally:
        if tunnel_proc:
            try:
                tunnel_proc.terminate()
            except Exception:
                pass

if __name__ == "__main__":
    main()
