import http.server
import socketserver
import socket
import os

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # doesn't even have to be reachable
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    return IP

PORT = 8000
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

ip_address = get_ip()

print("="*50)
print("GitMaster Pro - 모바일 접속 도우미")
print("="*50)
print(f"\n1. 같은 Wi-Fi에 연결된 휴대폰을 준비하세요.")
print(f"2. 휴대폰 브라우저(삼성 브라우저 등)를 엽니다.")
print(f"3. 아래 주소를 입력창에 입력하세요:\n")
print(f"   http://{ip_address}:{PORT}")
print(f"\n[주의] 반드시 https가 아닌 'http'로 접속해야 합니다.")
print(f"       (일부 브라우저가 자동으로 https로 연결하려고 할 수 있으니 확인해 주세요.)")
print(f"\n" + "="*50)
print("서버가 실행 중입니다... (종료하려면 Ctrl+C)")

with socketserver.TCPServer(("", PORT), Handler) as httpd:
    httpd.serve_forever()
