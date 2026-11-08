"""경인사무소 실행기 - 모바일/클라우드 자동 동기화 지원."""
from pathlib import Path
import os
import sys
import threading

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

# Set Gyeongin-specific environment / data paths
os.environ["PMI_SITE"] = "gyeongin"

from site_apps.gyeongin.src.app import MaterialManager

class GyeonginMaterialManager(MaterialManager):
    def __init__(self, root):
        # 1. 앱 실행 시 모바일 및 구글 클라우드 내역 선제 자동 동기화
        try:
            from site_apps.gyeongin.src.sync_google_sheet import sync_all_gyeongin_sources
            init_res = sync_all_gyeongin_sources()
            if init_res.get("added_count", 0) > 0:
                print(f"[경인 모바일] 시작 시 신규 {init_res['added_count']}건 자동 동기화 완료")
        except Exception as e:
            print(f"[경인 모바일] 시작 동기화 건너뜀: {e}")

        super().__init__(root)
        self.root.title("경인사무소 - 자재 및 소모품 관리 시스템")

        # 2. 주기적 백그라운드 자동 동기화 (앱 켜둔 상태에서도 30초마다 신규 모바일 입력 자동 감지)
        self.root.after(15000, self._start_periodic_sync)

    def trigger_mobile_sync(self, show_msg=True):
        """사용자가 '📱 모바일 동기화' 버튼 클릭 시 즉시 수동 동기화 실행"""
        try:
            from site_apps.gyeongin.src.sync_google_sheet import sync_all_gyeongin_sources
            res = sync_all_gyeongin_sources()
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
                    messagebox.showinfo("동기화 확인", "모든 모바일 데이터가 이미 최신 상태입니다.")
        except Exception as ex:
            if show_msg:
                from tkinter import messagebox
                messagebox.showerror("동기화 오류", f"동기화 중 오류가 발생했습니다:\n{ex}")

    def _start_periodic_sync(self):
        """백그라운드 스레드에서 모바일 신규 저장 내역 감지"""
        def worker():
            try:
                from site_apps.gyeongin.src.sync_google_sheet import sync_all_gyeongin_sources
                res = sync_all_gyeongin_sources()
                added = res.get("added_count", 0)
                if added > 0:
                    self.root.after(0, self._on_periodic_sync_found, added)
            except Exception:
                pass
            finally:
                # 30초 후 재확인
                try:
                    self.root.after(30000, self._start_periodic_sync)
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

    def _on_periodic_sync_found(self, count):
        """신규 모바일 데이터가 감지되었을 때 UI 자동 갱신"""
        try:
            self.load_data()
            if hasattr(self, 'update_daily_usage_view'):
                self.update_daily_usage_view()
            if hasattr(self, 'daily_save_status') and self.daily_save_status:
                self.daily_save_status.config(text=f"📱 모바일 신규 {count}건 자동 반영됨", foreground="green")
            print(f"[경인 모바일] 백그라운드 자동 동기화 완료: {count}건 반영")
        except Exception as ex:
            print(f"[경인 모바일] UI 갱신 실패: {ex}")

def main():
    import tkinter as tk
    root = tk.Tk()
    app = GyeonginMaterialManager(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        try:
            root.destroy()
        except Exception:
            pass

if __name__ == '__main__':
    main()
