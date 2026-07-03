import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import xlsxwriter

BASE_PATH = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_PATH, 'make_gian_config.json')

# ── 기본값 ──────────────────────────────────────────────────────────────────
DEFAULTS = {
    'drafter':    '이기중 상무',
    'dept':       '중무팀',
    'date':       '2026년 06월 23일',
    'body':       '서부서수소 가스공사서비(기산-기향 천연가스 공급권비)으로 아래와 같이 이동하오니 검토하시어 재가하여 주시기 바랍니다.',
    'items': [
        {'name': '컨테이너 3X9m\n운반비 및 상하차비',
         'supply': 2500000, 'vat': 250000,
         'note': '서부서수소 CGN대산에너지공장에서\n경인서비스소 기스공시(기향)\n→ 컨테이너 이동\n(2026.06.23 컨테이너 이동)'}
    ],
    'move_list':  '시주동 컨테이너 3X9m 1동(2024.4.30 중고 구매 350만원)\n숙계동 컨테이너 3X9m 1동(2025.6.18 요산서주소 공정통공업단지에서 인수)\n자재, 책상, 의자등 컨테이너 3X9m 1동(2024.7.09 중부서주소 신세종에서에서 인수)\n냉방/에어컨 3EA\n스탠드용 냉난방기 1EA\n사무용 책상, 의자 등 포함\n저장탕 현상망고, 건조기 등 포함',
    'special':    '경인사무소에서 컨테이너 이동 비용처리 예정',
    'attach1':    '컨테이너 이동 견적서 1부',
    'attach2':    '컨테이너 사진 대장 1부  곤',
    'slogan':     '"고객과 함께 미래를 선도하는 일류기업 지향"',
    'company':    'SITCO 서울검사 부자이하',
    'out_file':   os.path.join(BASE_PATH, '기안지.xlsx'),
}

def load_config():
    cfg = dict(DEFAULTS)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                cfg.update(saved)
        except: pass
    return cfg

def save_config(cfg):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=4)
    except: pass


# ══════════════════════════════════════════════════════════════════════════════
class GianApp:
    def __init__(self, root):
        self.root = root
        self.root.title("기안지 자동 생성")
        self.root.geometry("700x820")
        self.root.resizable(True, True)
        self.cfg = load_config()
        self._build_ui()

    # ── UI 구성 ────────────────────────────────────────────────────────────
    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill='both', expand=True, padx=8, pady=8)

        # ── Tab 1: 기본 정보 ──
        t1 = ttk.Frame(nb); nb.add(t1, text=' 기본 정보 ')
        self._tab_basic(t1)

        # ── Tab 2: 운송 비용표 ──
        t2 = ttk.Frame(nb); nb.add(t2, text=' 운송 비용표 ')
        self._tab_cost(t2)

        # ── Tab 3: 이동 목록 / 특이사항 ──
        t3 = ttk.Frame(nb); nb.add(t3, text=' 이동 목록 / 특이사항 ')
        self._tab_list(t3)

        # ── Tab 4: 하단/저장 ──
        t4 = ttk.Frame(nb); nb.add(t4, text=' 하단 / 저장 ')
        self._tab_footer(t4)

        # ── 실행 버튼 (항상 하단 고정) ──
        btn = tk.Button(self.root, text='📄  기안지 엑셀 생성',
                        font=('맑은 고딕', 12, 'bold'), bg='#2C7BE5', fg='white',
                        activebackground='#1a5cbf', cursor='hand2',
                        command=self.generate)
        btn.pack(fill='x', padx=8, pady=(0, 8), ipady=6)

    def _lf(self, parent, text, **kwargs):
        f = tk.LabelFrame(parent, text=text, font=('맑은 고딕', 9, 'bold'),
                          padx=8, pady=6, **kwargs)
        f.pack(fill='x', padx=8, pady=4)
        return f

    def _row(self, parent, label, var, row, width=50):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='e', padx=4, pady=3)
        ttk.Entry(parent, textvariable=var, width=width).grid(row=row, column=1, sticky='w', pady=3)

    # ── Tab 1 ──────────────────────────────────────────────────────────────
    def _tab_basic(self, parent):
        f = self._lf(parent, '결재 정보')
        self.v_drafter = tk.StringVar(value=self.cfg['drafter'])
        self.v_dept    = tk.StringVar(value=self.cfg['dept'])
        self.v_date    = tk.StringVar(value=self.cfg['date'])
        self._row(f, '기  한  자:', self.v_drafter, 0)
        self._row(f, '협  조  처:', self.v_dept,    1)
        self._row(f, '시행일자:', self.v_date,    2)

        f2 = self._lf(parent, '본문 내용')
        self.t_body = tk.Text(f2, height=8, wrap='word', font=('맑은 고딕', 10))
        self.t_body.insert('1.0', self.cfg['body'])
        self.t_body.pack(fill='both', expand=True)

    # ── Tab 2 ──────────────────────────────────────────────────────────────
    def _tab_cost(self, parent):
        f = self._lf(parent, '운송 비용 항목 (단일 행)')
        item = self.cfg['items'][0] if self.cfg['items'] else {}
        self.v_iname  = tk.StringVar(value=item.get('name', ''))
        self.v_supply = tk.StringVar(value=str(item.get('supply', 0)))
        self.v_vat    = tk.StringVar(value=str(item.get('vat', 0)))
        self.v_note   = tk.StringVar(value=item.get('note', ''))

        self._row(f, '품    명:', self.v_iname,  0, 55)
        self._row(f, '공급가액(원):', self.v_supply, 1, 20)
        self._row(f, '부가가치세(원):', self.v_vat,    2, 20)
        ttk.Label(f, text='비    고:').grid(row=3, column=0, sticky='ne', padx=4, pady=3)
        self.t_note = tk.Text(f, height=5, width=52, font=('맑은 고딕', 10))
        self.t_note.insert('1.0', item.get('note', ''))
        self.t_note.grid(row=3, column=1, sticky='w', pady=3)

        hint = tk.Label(f, text='※ 합계는 공급가액+부가가치세로 자동 계산됩니다.',
                        fg='gray', font=('맑은 고딕', 9))
        hint.grid(row=4, column=0, columnspan=2, pady=4)

    # ── Tab 3 ──────────────────────────────────────────────────────────────
    def _tab_list(self, parent):
        f = self._lf(parent, '이동 목록 (줄바꿈으로 항목 구분)')
        self.t_move = tk.Text(f, height=10, wrap='word', font=('맑은 고딕', 10))
        self.t_move.insert('1.0', self.cfg['move_list'])
        self.t_move.pack(fill='both', expand=True)

        f2 = self._lf(parent, '특이사항')
        self.t_special = tk.Text(f2, height=4, wrap='word', font=('맑은 고딕', 10))
        self.t_special.insert('1.0', self.cfg['special'])
        self.t_special.pack(fill='both', expand=True)

    # ── Tab 4 ──────────────────────────────────────────────────────────────
    def _tab_footer(self, parent):
        f = self._lf(parent, '첨부 문서')
        self.v_att1 = tk.StringVar(value=self.cfg['attach1'])
        self.v_att2 = tk.StringVar(value=self.cfg['attach2'])
        self._row(f, '첨부 1:', self.v_att1, 0, 55)
        self._row(f, '첨부 2:', self.v_att2, 1, 55)

        f2 = self._lf(parent, '하단 슬로건 / 회사명')
        self.v_slogan  = tk.StringVar(value=self.cfg['slogan'])
        self.v_company = tk.StringVar(value=self.cfg['company'])
        self._row(f2, '슬로건:', self.v_slogan,  0, 55)
        self._row(f2, '회사명:', self.v_company, 1, 55)

        f3 = self._lf(parent, '저장 위치')
        self.v_out = tk.StringVar(value=self.cfg['out_file'])
        row3 = tk.Frame(f3); row3.pack(fill='x')
        ttk.Entry(row3, textvariable=self.v_out, width=52).pack(side='left')
        tk.Button(row3, text='찾아보기', command=self._browse_save).pack(side='left', padx=4)

    def _browse_save(self):
        p = filedialog.asksaveasfilename(
            initialdir=os.path.dirname(self.v_out.get()),
            initialfile=os.path.basename(self.v_out.get()),
            defaultextension='.xlsx',
            filetypes=[('Excel', '*.xlsx')])
        if p: self.v_out.set(p)

    # ── 설정 수집 ──────────────────────────────────────────────────────────
    def _collect(self):
        try: supply = int(self.v_supply.get().replace(',', ''))
        except: supply = 0
        try: vat = int(self.v_vat.get().replace(',', ''))
        except: vat = 0
        return {
            'drafter':   self.v_drafter.get().strip(),
            'dept':      self.v_dept.get().strip(),
            'date':      self.v_date.get().strip(),
            'body':      self.t_body.get('1.0', 'end-1c'),
            'items':     [{'name': self.v_iname.get(),
                           'supply': supply, 'vat': vat,
                           'note': self.t_note.get('1.0', 'end-1c')}],
            'move_list': self.t_move.get('1.0', 'end-1c'),
            'special':   self.t_special.get('1.0', 'end-1c'),
            'attach1':   self.v_att1.get().strip(),
            'attach2':   self.v_att2.get().strip(),
            'slogan':    self.v_slogan.get().strip(),
            'company':   self.v_company.get().strip(),
            'out_file':  self.v_out.get().strip(),
        }

    # ── 엑셀 생성 ──────────────────────────────────────────────────────────
    def generate(self):
        cfg = self._collect()
        save_config(cfg)
        out = cfg['out_file']
        if not out.lower().endswith('.xlsx'):
            out += '.xlsx'

        try:
            wb = xlsxwriter.Workbook(out)
            ws = wb.add_worksheet()

            ws.set_paper(9)       # A4
            ws.set_portrait()
            ws.set_margins(left=0.9, right=0.7, top=0.7, bottom=0.7)

            # ── 열 너비 설정 ──────────────────────────────────────────────
            # A:J = 10열 사용
            ws.set_column('A:A', 3)
            ws.set_column('B:B', 12)
            ws.set_column('C:C', 14)
            ws.set_column('D:D', 14)
            ws.set_column('E:E', 14)
            ws.set_column('F:F', 24)
            ws.set_column('G:J', 3)

            # ── 공통 포맷 ─────────────────────────────────────────────────
            def fmt(**kw):
                base = {'font_name': '맑은 고딕', 'font_size': 10,
                        'valign': 'vcenter', 'text_wrap': True}
                base.update(kw)
                return wb.add_format(base)

            f_title   = fmt(font_size=18, bold=True, align='center', border=0)
            f_hdr_lbl = fmt(bold=True, align='center', border=1, bg_color='#D9D9D9')
            f_hdr_val = fmt(align='left', border=1)
            f_center  = fmt(align='center', border=1)
            f_left    = fmt(align='left', border=1)
            f_right   = fmt(align='right', border=1)
            f_money   = fmt(align='right', border=1, num_format='#,##0')
            f_tbl_hdr = fmt(bold=True, align='center', border=1, bg_color='#BDD7EE')
            f_body    = fmt(align='justify', border=0)
            f_section = fmt(bold=True, align='left', border=0, font_size=10)
            f_plain   = fmt(align='left', border=0)
            f_footer  = fmt(align='center', border=1, italic=True, fg_color='#595959')

            R = 0  # 현재 행 (0-indexed)

            # ── 제목 ──────────────────────────────────────────────────────
            ws.set_row(R, 36)
            ws.merge_range(R, 0, R, 6, '기    안    지', f_title)
            R += 1

            # ── 헤더 (기한자 / 협조처 / 시행일자) ────────────────────────
            for label, val in [('기  한  자', cfg['drafter']),
                                ('협  조  처', cfg['dept']),
                                ('시 행 일 자', cfg['date'])]:
                ws.set_row(R, 20)
                ws.merge_range(R, 0, R, 0, '', f_hdr_lbl)
                ws.write(R, 0, label, f_hdr_lbl)
                # label cell spans B
                ws.write(R, 1, label, f_hdr_lbl)
                ws.merge_range(R, 1, R, 1, label, f_hdr_lbl)
                ws.merge_range(R, 2, R, 6, val, f_hdr_val)
                R += 1

            # ── 구분선 ────────────────────────────────────────────────────
            ws.set_row(R, 6)
            ws.merge_range(R, 0, R, 6, '', fmt(border=0, bottom=2))
            R += 1

            # ── 본문 ──────────────────────────────────────────────────────
            ws.set_row(R, 80)
            ws.merge_range(R, 0, R, 6, cfg['body'], f_body)
            R += 1

            # ── 아 래 ─────────────────────────────────────────────────────
            ws.set_row(R, 22)
            ws.merge_range(R, 0, R, 6, '아    래', fmt(align='center', bold=True, border=0))
            R += 1

            # ── 섹션 1 제목 ───────────────────────────────────────────────
            ws.set_row(R, 18)
            ws.merge_range(R, 0, R, 6, '1. 컨테이너 운송', f_section)
            R += 1

            # ── 비용 표 헤더 ──────────────────────────────────────────────
            ws.set_row(R, 22)
            ws.merge_range(R, 0, R, 1, '품  명',   f_tbl_hdr)
            ws.write(R, 2, '공급가액', f_tbl_hdr)
            ws.write(R, 3, '부가가치세', f_tbl_hdr)
            ws.write(R, 4, '합  계', f_tbl_hdr)
            ws.write(R, 5, '비  고', f_tbl_hdr)
            ws.write(R, 6, '', f_tbl_hdr)
            R += 1

            # ── 비용 표 데이터 ────────────────────────────────────────────
            for itm in cfg['items']:
                supply = itm['supply']; vat = itm['vat']; total = supply + vat
                ws.set_row(R, 60)
                ws.merge_range(R, 0, R, 1, itm['name'], f_center)
                ws.write(R, 2, supply, f_money)
                ws.write(R, 3, vat,    f_money)
                ws.write(R, 4, total,  f_money)
                ws.merge_range(R, 5, R, 6, itm['note'], f_left)
                R += 1

            ws.set_row(R, 6)
            R += 1  # 여백

            # ── 섹션 2: 이동 목록 ─────────────────────────────────────────
            ws.set_row(R, 18)
            ws.merge_range(R, 0, R, 6, '2. 컨테이너 이동 목록', f_section)
            R += 1

            for line in cfg['move_list'].split('\n'):
                ws.set_row(R, 16)
                ws.merge_range(R, 0, R, 6, f'  {line}', f_plain)
                R += 1

            ws.set_row(R, 6); R += 1

            # ── 섹션 3: 특이사항 ──────────────────────────────────────────
            ws.set_row(R, 18)
            ws.merge_range(R, 0, R, 6, '3. 특이사항', f_section)
            R += 1

            for line in cfg['special'].split('\n'):
                ws.set_row(R, 16)
                ws.merge_range(R, 0, R, 6, f'  {line}', f_plain)
                R += 1

            ws.set_row(R, 8); R += 1

            # ── 구분선 ────────────────────────────────────────────────────
            ws.merge_range(R, 0, R, 6, '', fmt(border=0, bottom=1)); R += 1

            # ── 첨부 ──────────────────────────────────────────────────────
            ws.set_row(R, 16)
            ws.merge_range(R, 0, R, 6, f'※ 첨  부 :  {cfg["attach1"]}', f_plain); R += 1
            ws.set_row(R, 16)
            ws.merge_range(R, 0, R, 6, f'          {cfg["attach2"]}', f_plain); R += 1

            ws.set_row(R, 8); R += 1

            # ── 하단 슬로건 ───────────────────────────────────────────────
            ws.set_row(R, 30)
            bottom_txt = f'{cfg["slogan"]}  {cfg["company"]}'
            ws.merge_range(R, 0, R, 6, bottom_txt,
                           fmt(align='center', border=2, bold=False,
                               font_size=9, fg_color='#404040',
                               bg_color='#F2F2F2'))
            R += 1

            wb.close()
            messagebox.showinfo('완료', f'기안지가 생성되었습니다!\n\n{out}')

        except Exception as e:
            messagebox.showerror('오류', f'생성 실패:\n{e}')


if __name__ == '__main__':
    root = tk.Tk()
    app = GianApp(root)
    root.mainloop()
