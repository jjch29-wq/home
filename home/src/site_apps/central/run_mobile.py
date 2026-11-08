"""
중앙지사 현장 모바일 검사 입력 앱 실행기
스마트폰 / 태블릿 / PC 어디서나 브라우저로 접속 가능합니다.
"""
import socket
import sys
from pathlib import Path

# Add home/src to sys.path
HOME_SRC = Path(__file__).resolve().parents[2]
if str(HOME_SRC) not in sys.path:
    sys.path.insert(0, str(HOME_SRC))

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def main():
    import uvicorn
    local_ip = get_local_ip()
    port = 8000

    print("=" * 60)
    print("   🚀 한국지역난방공사 중앙지사 모바일 현장 입력 시스템")
    print("=" * 60)
    print(f" [PC 브라우저 접속]        👉 http://localhost:{port}")
    print(f" [현장 스마트폰/태블릿 접속] 👉 http://{local_ip}:{port}")
    print("=" * 60)
    print(" * 동일한 Wi-Fi(또는 핫스팟/네트워크)에 연결된 스마트폰에서")
    print(f"   위 스마트폰 접속 주소(http://{local_ip}:{port})로 접속하시면 됩니다.")
    print(" * 스마트폰 브라우저에서 '홈 화면에 추가'를 누르면 앱처럼 실행됩니다.")
    print("=" * 60)
    print(" 서버를 종료하려면 터미널에서 Ctrl + C를 누르세요.\n")

    uvicorn.run("site_apps.central.mobile.server:app", host="0.0.0.0", port=port, reload=True)

if __name__ == "__main__":
    main()
