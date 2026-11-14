"""한국지역난방 중앙지사 실행기 - 모바일/클라우드 자동 동기화 지원."""
from pathlib import Path
import os
import sys
import threading

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from site_apps.central.src.app import MaterialManager
from site_apps.central.src.runtime_profile import is_development

class CentralMaterialManager(MaterialManager):
    def __init__(self, root):
        if is_development():
            super().__init__(root)
            self.root.title('[개발용 · 운영 동기화 차단] 중앙지사 자재관리')
            return
        # 1. 앱 실행 시 모바일 및 구글 클라우드 내역 선제 자동 동기화
        try:
            from site_apps.central.src.sync_google_sheet import sync_all_central_sources
            init_res = sync_all_central_sources()
            if init_res.get("status") != "success":
                print(f"[중앙지사 모바일] 시작 동기화 오류: {init_res.get('message', '')}")
            elif init_res.get("added_count", 0) > 0:
                print(f"[중앙지사 모바일] 시작 시 신규 {init_res['added_count']}건 자동 동기화 완료")
        except Exception as e:
            print(f"[중앙지사 모바일] 시작 동기화 건너뜀: {e}")

        super().__init__(root)
        self.root.title("한국지역난방공사 중앙지사 - 자재 및 소모품 관리 시스템")

        # 2. 주기적 백그라운드 자동 동기화 (앱 켜둔 상태에서도 30초마다 신규 모바일 입력 자동 감지)
        self.root.after(15000, self._start_periodic_sync)

    def open_google_sync_dialog(self):
        """구글 웹 앱 URL 설정 대화상자"""
        if is_development():
            self.trigger_mobile_sync()
            return
        import tkinter.simpledialog as sd
        from tkinter import messagebox
        from site_apps.central.src.sync_google_sheet import load_sync_config, save_sync_config
        cfg = load_sync_config()
        current_url = cfg.get("web_app_url", "")
        new_url = sd.askstring(
            "구글 클라우드 웹 앱 URL 설정",
            "Google Apps Script에서 배포한 웹 앱 URL을 입력하세요:\n(예: https://script.google.com/macros/s/.../exec)",
            initialvalue=current_url,
            parent=self.root
        )
        if new_url is not None:
            cfg["web_app_url"] = new_url.strip()
            if save_sync_config(cfg):
                messagebox.showinfo("설정 완료", "구글 클라우드 웹 앱 URL이 성공적으로 저장되었습니다.")
                if new_url.strip():
                    self.trigger_mobile_sync(show_msg=True)

    def trigger_mobile_sync(self, show_msg=True):
        """사용자가 '📱 모바일 동기화' 버튼 클릭 시 즉시 수동 동기화 실행"""
        if is_development():
            if show_msg:
                from tkinter import messagebox
                messagebox.showinfo('개발용', '운영 클라우드 동기화가 차단되어 있습니다. 테스트 모바일은 http://127.0.0.1:8001 을 사용하세요.')
            return
        try:
            from site_apps.central.src.sync_google_sheet import sync_all_central_sources, load_sync_config
            cfg = load_sync_config()
            if not cfg.get("web_app_url") and show_msg:
                from tkinter import messagebox
                if messagebox.askyesno("구글 웹 앱 URL 미등록", "구글 클라우드 웹 앱 URL이 아직 등록되지 않았습니다.\n지금 URL을 등록하여 24시간 연동을 활성화하시겠습니까?"):
                    self.open_google_sync_dialog()
                    return

            res = sync_all_central_sources()
            if res.get("status") != "success":
                if show_msg:
                    from tkinter import messagebox
                    messagebox.showerror("동기화 오류", res.get("message", "모바일 동기화에 실패했습니다."))
                return
            added = res.get("added_count", 0)
            if added > 0:
                self.load_data()
                if hasattr(self, 'update_daily_usage_view'):
                    self.update_daily_usage_view()
                if show_msg:
                    from tkinter import messagebox
                    messagebox.showinfo("모바일 동기화 완료", f"📱 모바일/클라우드에서 {added}건의 신규 작업이 동기화되었습니다.")
            else:
                if show_msg:
                    from tkinter import messagebox
                    messagebox.showinfo("동기화 확인", f"모든 모바일 데이터가 이미 최신 상태입니다.\n({res.get('message', '')})")
        except Exception as ex:
            if show_msg:
                from tkinter import messagebox
                messagebox.showerror("동기화 오류", f"동기화 중 오류가 발생했습니다:\n{ex}")

    def _start_periodic_sync(self):
        """백그라운드 스레드에서 모바일 신규 저장 내역 감지"""
        if is_development():
            return
        def worker():
            try:
                from site_apps.central.src.sync_google_sheet import sync_all_central_sources
                res = sync_all_central_sources()
                if res.get("status") != "success":
                    self.root.after(0, self._on_periodic_sync_error, res.get("message", "모바일 동기화 실패"))
                elif res.get("added_count", 0) > 0:
                    added = res["added_count"]
                    self.root.after(0, self._on_periodic_sync_found, added)
            except Exception:
                pass
            finally:
                try:
                    self.root.after(30000, self._start_periodic_sync)
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

    def _on_periodic_sync_error(self, message):
        """백그라운드 동기화 실패를 숨기지 않고 상태 영역과 콘솔에 표시한다."""
        if hasattr(self, 'daily_save_status') and self.daily_save_status:
            self.daily_save_status.config(text="📱 모바일 동기화 오류", foreground="red")
        print(f"[중앙지사 모바일] 백그라운드 동기화 오류: {message}")

    def _on_periodic_sync_found(self, count):
        """신규 모바일 데이터가 감지되었을 때 UI 자동 갱신"""
        try:
            self.load_data()
            if hasattr(self, 'update_daily_usage_view'):
                self.update_daily_usage_view()
            if hasattr(self, 'daily_save_status') and self.daily_save_status:
                self.daily_save_status.config(text=f"📱 모바일 신규 {count}건 자동 반영됨", foreground="green")
            print(f"[중앙지사 모바일] 백그라운드 자동 동기화 완료: {count}건 반영")
        except Exception as ex:
            print(f"[중앙지사 모바일] UI 갱신 실패: {ex}")

def main():
    import tkinter as tk
    root = tk.Tk()
    app = CentralMaterialManager(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        try:
            root.destroy()
        except Exception:
            pass

if __name__ == '__main__':
    main()
