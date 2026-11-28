"""Client, Contract, Flexible Billing & Template Library management tab for NDT-Pro."""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from site_apps.central.src.services.client_contract_store import (
    CONTRACT_TYPE_LABELS,
    delete_client,
    delete_contract,
    load_client_contract_store,
    upsert_client,
    upsert_contract,
)
from site_apps.central.src.services.flexible_billing_service import (
    calculate_flexible_billing,
    export_flexible_billing_excel,
)
from site_apps.central.src.services.template_library import (
    DEFAULT_CLIENT_KEY,
    SUPPORTED_TEMPLATE_MODES,
    list_templates,
    register_template,
    remove_template,
)
from site_apps.central.src.site_profile import get_active_profile


class ClientContractTab(ttk.Frame):
    """Unified tab for managing clients, annual/per-job contracts, custom rates, and templates."""

    def __init__(self, parent, main_app=None):
        super().__init__(parent, padding=8)
        self.main_app = main_app
        self.profile = getattr(main_app, "site_profile", None) or get_active_profile()
        self._clients_by_id: dict[str, dict] = {}
        self._contracts_by_id: dict[str, dict] = {}

        self._build_ui()
        self.refresh_all()

    def _build_ui(self):
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # Left pane: Client master
        left = ttk.LabelFrame(paned, text="🏢 거래처 마스터", padding=8)
        paned.add(left, weight=2)

        left_btns = ttk.Frame(left)
        left_btns.pack(fill="x", pady=(0, 6))
        ttk.Button(left_btns, text="+ 거래처 추가", command=self._add_client, style="Primary.TButton").pack(side="left", padx=2)
        ttk.Button(left_btns, text="✏️ 수정", command=self._edit_selected_client).pack(side="left", padx=2)
        ttk.Button(left_btns, text="🗑️ 삭제", command=self._delete_selected_client, style="Danger.TButton").pack(side="left", padx=2)
        ttk.Button(left_btns, text="🔄 업체목록 반영", command=self._sync_companies_to_app, style="Secondary.TButton").pack(side="right", padx=2)

        c_cols = ("name", "biz_no", "contact", "terms")
        self.client_tree = ttk.Treeview(left, columns=c_cols, show="headings", selectmode="browse")
        self.client_tree.heading("name", text="거래처명")
        self.client_tree.heading("biz_no", text="사업자번호")
        self.client_tree.heading("contact", text="담당자 / 연락처")
        self.client_tree.heading("terms", text="결제조건")
        self.client_tree.column("name", width=150, anchor="w")
        self.client_tree.column("biz_no", width=110, anchor="center")
        self.client_tree.column("contact", width=140, anchor="w")
        self.client_tree.column("terms", width=130, anchor="w")
        self.client_tree.pack(fill="both", expand=True)
        self.client_tree.bind("<<TreeviewSelect>>", self._on_client_select)
        self.client_tree.bind("<Double-1>", lambda _e: self._edit_selected_client())

        # Right pane: Contracts (top) + Templates (bottom)
        right = ttk.PanedWindow(paned, orient="vertical")
        paned.add(right, weight=3)

        # Contracts section
        ctr_box = ttk.LabelFrame(right, text="📑 계약 및 자유 단가표", padding=8)
        right.add(ctr_box, weight=3)

        ctr_btns = ttk.Frame(ctr_box)
        ctr_btns.pack(fill="x", pady=(0, 6))
        ttk.Button(ctr_btns, text="+ 계약 추가", command=self._add_contract, style="Primary.TButton").pack(side="left", padx=2)
        ttk.Button(ctr_btns, text="✏️ 계약·단가표 편집", command=self._edit_selected_contract).pack(side="left", padx=2)
        ttk.Button(ctr_btns, text="🗑️ 삭제", command=self._delete_selected_contract, style="Danger.TButton").pack(side="left", padx=2)
        ttk.Button(ctr_btns, text="📄 기성·청구서 발행", command=self._export_selected_contract_billing, style="Success.TButton").pack(side="right", padx=2)

        k_cols = ("contract_name", "contract_type", "period", "budget", "rate_cnt")
        self.contract_tree = ttk.Treeview(ctr_box, columns=k_cols, show="headings", selectmode="browse")
        self.contract_tree.heading("contract_name", text="계약명")
        self.contract_tree.heading("contract_type", text="계약유형")
        self.contract_tree.heading("period", text="계약기간")
        self.contract_tree.heading("budget", text="계약총액")
        self.contract_tree.heading("rate_cnt", text="단가항목")
        self.contract_tree.column("contract_name", width=190, anchor="w")
        self.contract_tree.column("contract_type", width=105, anchor="center")
        self.contract_tree.column("period", width=155, anchor="center")
        self.contract_tree.column("budget", width=110, anchor="e")
        self.contract_tree.column("rate_cnt", width=70, anchor="center")
        self.contract_tree.pack(fill="both", expand=True)
        self.contract_tree.bind("<Double-1>", lambda _e: self._edit_selected_contract())

        # Template library section
        tpl_box = ttk.LabelFrame(right, text="📂 거래처별 성적서·사진대장 템플릿 라이브러리", padding=8)
        right.add(tpl_box, weight=2)

        tpl_btns = ttk.Frame(tpl_box)
        tpl_btns.pack(fill="x", pady=(0, 6))
        self.tpl_mode_var = tk.StringVar(value="RT")
        ttk.Label(tpl_btns, text="모드:").pack(side="left", padx=(0, 4))
        ttk.Combobox(
            tpl_btns,
            textvariable=self.tpl_mode_var,
            values=SUPPORTED_TEMPLATE_MODES,
            state="readonly",
            width=10,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(tpl_btns, text="+ 템플릿 등록", command=self._register_template, style="Primary.TButton").pack(side="left", padx=2)
        ttk.Button(tpl_btns, text="📂 파일 열기", command=self._open_selected_template, style="Secondary.TButton").pack(side="left", padx=2)
        ttk.Button(tpl_btns, text="🗑️ 삭제", command=self._delete_selected_template, style="Danger.TButton").pack(side="left", padx=2)

        t_cols = ("scope", "mode", "filename", "updated_at")
        self.tpl_tree = ttk.Treeview(tpl_box, columns=t_cols, show="headings", selectmode="browse")
        self.tpl_tree.heading("scope", text="적용 대상")
        self.tpl_tree.heading("mode", text="검사 모드")
        self.tpl_tree.heading("filename", text="등록 파일명")
        self.tpl_tree.heading("updated_at", text="갱신일시")
        self.tpl_tree.column("scope", width=130, anchor="center")
        self.tpl_tree.column("mode", width=85, anchor="center")
        self.tpl_tree.column("filename", width=240, anchor="w")
        self.tpl_tree.column("updated_at", width=140, anchor="center")
        self.tpl_tree.pack(fill="both", expand=True)

    def _selected_client_id(self) -> str | None:
        sel = self.client_tree.selection()
        return sel[0] if sel else None

    def _selected_contract_id(self) -> str | None:
        sel = self.contract_tree.selection()
        return sel[0] if sel else None

    def refresh_all(self):
        store = load_client_contract_store(self.profile)
        self._clients_by_id = {c["client_id"]: c for c in store["clients"]}
        self._contracts_by_id = {ct["contract_id"]: ct for ct in store["contracts"]}

        prev_client = self._selected_client_id()
        self.client_tree.delete(*self.client_tree.get_children())
        for c in store["clients"]:
            contact = f"{c.get('contact_name', '')} {c.get('contact_phone', '')}".strip()
            self.client_tree.insert(
                "",
                "end",
                iid=c["client_id"],
                values=(c.get("name", ""), c.get("biz_no", ""), contact, c.get("payment_terms", "")),
            )
        if prev_client and prev_client in self._clients_by_id:
            self.client_tree.selection_set(prev_client)
        self._refresh_contracts_and_templates()

    def _on_client_select(self, _event=None):
        self._refresh_contracts_and_templates()

    def _refresh_contracts_and_templates(self):
        client_id = self._selected_client_id()
        self.contract_tree.delete(*self.contract_tree.get_children())
        for ct in self._contracts_by_id.values():
            if client_id and ct.get("client_id") != client_id:
                continue
            period = f"{ct.get('start_date', '')} ~ {ct.get('end_date', '')}".strip(" ~")
            self.contract_tree.insert(
                "",
                "end",
                iid=ct["contract_id"],
                values=(
                    ct.get("contract_name", ""),
                    CONTRACT_TYPE_LABELS.get(ct.get("contract_type", ""), ct.get("contract_type", "")),
                    period,
                    f"{int(ct.get('budget_amount', 0)):,}원",
                    f"{len(ct.get('rate_items', []))}개",
                ),
            )

        self.tpl_tree.delete(*self.tpl_tree.get_children())
        entries = list_templates(client_id=client_id, profile=self.profile) if client_id else list_templates(profile=self.profile)
        for item in entries:
            cid = item.get("client_id", DEFAULT_CLIENT_KEY)
            scope = "기본 공용" if cid == DEFAULT_CLIENT_KEY else self._clients_by_id.get(cid, {}).get("name", cid)
            iid = f"{cid}::{item['mode']}"
            self.tpl_tree.insert(
                "",
                "end",
                iid=iid,
                values=(scope, item.get("mode", ""), item.get("filename", ""), item.get("updated_at", "")),
            )

    def _open_client_dialog(self, existing: dict | None = None):
        dlg = tk.Toplevel(self)
        dlg.title("거래처 수정" if existing else "거래처 추가")
        dlg.geometry("430x330")
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)

        fields = [
            ("name", "거래처명 *"),
            ("biz_no", "사업자번호"),
            ("ceo", "대표자"),
            ("contact_name", "담당자명"),
            ("contact_phone", "연락처"),
            ("contact_email", "이메일"),
            ("payment_terms", "결제조건"),
            ("note", "비고"),
        ]
        vars_map: dict[str, tk.StringVar] = {}
        for row, (key, label) in enumerate(fields):
            ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=(existing or {}).get(key, "월말 마감 익월 15일" if key == "payment_terms" else ""))
            vars_map[key] = var
            ttk.Entry(frm, textvariable=var, width=32).grid(row=row, column=1, sticky="ew", pady=3)
        frm.columnconfigure(1, weight=1)

        def save():
            payload = {k: v.get() for k, v in vars_map.items()}
            if existing:
                payload["client_id"] = existing["client_id"]
            try:
                upsert_client(payload, profile=self.profile)
            except Exception as err:
                messagebox.showerror("저장 오류", str(err), parent=dlg)
                return
            dlg.destroy()
            self.refresh_all()

        btns = ttk.Frame(frm)
        btns.grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="취소", command=dlg.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="저장", command=save, style="Primary.TButton").pack(side="right")

    def _add_client(self):
        self._open_client_dialog(None)

    def _edit_selected_client(self):
        cid = self._selected_client_id()
        if not cid or cid not in self._clients_by_id:
            messagebox.showinfo("안내", "수정할 거래처를 선택하세요.", parent=self)
            return
        self._open_client_dialog(self._clients_by_id[cid])

    def _delete_selected_client(self):
        cid = self._selected_client_id()
        if not cid or cid not in self._clients_by_id:
            return
        name = self._clients_by_id[cid].get("name", cid)
        if not messagebox.askyesno("삭제 확인", f"거래처 '{name}' 및 소속 계약을 삭제하시겠습니까?", parent=self):
            return
        delete_client(cid, profile=self.profile)
        self.refresh_all()

    def _sync_companies_to_app(self):
        if not self.main_app or not hasattr(self.main_app, "companies"):
            return
        added = 0
        for c in self._clients_by_id.values():
            name = c.get("name", "").strip()
            if name and name not in self.main_app.companies:
                self.main_app.companies.append(name)
                added += 1
        if hasattr(self.main_app, "update_registration_combos"):
            try:
                self.main_app.update_registration_combos()
            except Exception:
                pass
        messagebox.showinfo("동기화 완료", f"거래처 {added}건을 앱 업체 목록에 반영했습니다.", parent=self)

    def _open_contract_dialog(self, existing: dict | None = None):
        client_id = (existing or {}).get("client_id") or self._selected_client_id()
        if not client_id:
            messagebox.showwarning("거래처 선택", "먼저 좌측에서 거래처를 선택하세요.", parent=self)
            return

        dlg = tk.Toplevel(self)
        dlg.title("계약·단가표 편집" if existing else "새 계약 추가")
        dlg.geometry("680x520")
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)

        name_var = tk.StringVar(value=(existing or {}).get("contract_name", ""))
        type_var = tk.StringVar(value=(existing or {}).get("contract_type", "annual_unit"))
        start_var = tk.StringVar(value=(existing or {}).get("start_date", ""))
        end_var = tk.StringVar(value=(existing or {}).get("end_date", ""))
        budget_var = tk.StringVar(value=str((existing or {}).get("budget_amount", 0)))
        trip_var = tk.StringVar(value=str((existing or {}).get("default_trip_fee", 0)))
        lump_var = tk.StringVar(value=str((existing or {}).get("default_lump_sum", 0)))

        top = ttk.Frame(frm)
        top.pack(fill="x")
        ttk.Label(top, text="계약명 *:").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(top, textvariable=name_var, width=28).grid(row=0, column=1, sticky="ew", padx=4, pady=3)
        ttk.Label(top, text="계약유형:").grid(row=0, column=2, sticky="w", pady=3)
        ttk.Combobox(top, textvariable=type_var, values=tuple(CONTRACT_TYPE_LABELS.keys()), state="readonly", width=16).grid(row=0, column=3, sticky="ew", padx=4, pady=3)

        ttk.Label(top, text="시작일:").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(top, textvariable=start_var, width=16).grid(row=1, column=1, sticky="ew", padx=4, pady=3)
        ttk.Label(top, text="종료일:").grid(row=1, column=2, sticky="w", pady=3)
        ttk.Entry(top, textvariable=end_var, width=16).grid(row=1, column=3, sticky="ew", padx=4, pady=3)

        ttk.Label(top, text="계약총액:").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(top, textvariable=budget_var, width=16).grid(row=2, column=1, sticky="ew", padx=4, pady=3)
        ttk.Label(top, text="기본 출장비 / 일식:").grid(row=2, column=2, sticky="w", pady=3)
        fee_row = ttk.Frame(top)
        fee_row.grid(row=2, column=3, sticky="ew", padx=4, pady=3)
        ttk.Entry(fee_row, textvariable=trip_var, width=9).pack(side="left", padx=(0, 4))
        ttk.Entry(fee_row, textvariable=lump_var, width=9).pack(side="left")
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=1)

        ttk.Label(
            frm,
            text="자유 단가표 (한 줄당: 품목코드, 품목명, 검사방법, 규격, 단위, 단가, 계약수량)",
            font=("Malgun Gothic", 9, "bold"),
        ).pack(anchor="w", pady=(10, 4))

        rate_text = tk.Text(frm, height=12, font=("Consolas", 10))
        rate_text.pack(fill="both", expand=True)
        existing_rates = (existing or {}).get("rate_items") or []
        if existing_rates:
            lines = [
                f"{r.get('code','')}, {r.get('name','')}, {r.get('method','')}, {r.get('spec','')}, {r.get('unit','EA')}, {r.get('unit_price',0)}, {r.get('contract_qty',0)}"
                for r in existing_rates
            ]
            rate_text.insert("1.0", "\n".join(lines))
        else:
            rate_text.insert("1.0", "RT-01, 배관 RT 촬영, RT, 100A, 매, 25000, 100\nMT-01, 자분탐상검사, MT, 일반, m, 30000, 50")

        def save():
            parsed_rates = []
            for raw_line in rate_text.get("1.0", "end").splitlines():
                line = raw_line.strip()
                if not line:
                    continue
                parts = [p.strip() for p in line.split(",")]
                while len(parts) < 7:
                    parts.append("")
                parsed_rates.append(
                    {
                        "code": parts[0],
                        "name": parts[1],
                        "method": parts[2],
                        "spec": parts[3],
                        "unit": parts[4] or "EA",
                        "unit_price": parts[5] or 0,
                        "contract_qty": parts[6] or 0,
                    }
                )
            payload = {
                "client_id": client_id,
                "contract_name": name_var.get(),
                "contract_type": type_var.get(),
                "start_date": start_var.get(),
                "end_date": end_var.get(),
                "budget_amount": budget_var.get(),
                "default_trip_fee": trip_var.get(),
                "default_lump_sum": lump_var.get(),
                "rate_items": parsed_rates,
            }
            if existing:
                payload["contract_id"] = existing["contract_id"]
            try:
                upsert_contract(payload, profile=self.profile)
            except Exception as err:
                messagebox.showerror("저장 오류", str(err), parent=dlg)
                return
            dlg.destroy()
            self.refresh_all()

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="취소", command=dlg.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="계약 저장", command=save, style="Primary.TButton").pack(side="right")

    def _add_contract(self):
        self._open_contract_dialog(None)

    def _edit_selected_contract(self):
        ct_id = self._selected_contract_id()
        if not ct_id or ct_id not in self._contracts_by_id:
            messagebox.showinfo("안내", "편집할 계약을 선택하세요.", parent=self)
            return
        self._open_contract_dialog(self._contracts_by_id[ct_id])

    def _delete_selected_contract(self):
        ct_id = self._selected_contract_id()
        if not ct_id or ct_id not in self._contracts_by_id:
            return
        name = self._contracts_by_id[ct_id].get("contract_name", ct_id)
        if not messagebox.askyesno("삭제 확인", f"계약 '{name}'을(를) 삭제하시겠습니까?", parent=self):
            return
        delete_contract(ct_id, profile=self.profile)
        self.refresh_all()

    def _export_selected_contract_billing(self):
        ct_id = self._selected_contract_id()
        if not ct_id or ct_id not in self._contracts_by_id:
            messagebox.showinfo("안내", "청구서를 발행할 계약을 선택하세요.", parent=self)
            return
        contract = self._contracts_by_id[ct_id]
        client = self._clients_by_id.get(contract.get("client_id", ""))
        fp = filedialog.asksaveasfilename(
            title="유연 기성·청구서 저장",
            defaultextension=".xlsx",
            initialfile=f"{contract.get('contract_name', '기성청구서')}.xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            parent=self,
        )
        if not fp:
            return
        usage = [
            {"code": item["code"], "qty": item.get("contract_qty", 0)}
            for item in contract.get("rate_items", [])
        ]
        if contract.get("default_trip_fee", 0) or contract.get("default_lump_sum", 0):
            usage.append(
                {
                    "name": "기본 출장비/일식",
                    "qty": 0,
                    "trip_fee": contract.get("default_trip_fee", 0),
                    "lump_sum": contract.get("default_lump_sum", 0),
                }
            )
        result = calculate_flexible_billing(contract, usage)
        export_flexible_billing_excel(contract, client, result, fp)
        messagebox.showinfo(
            "발행 완료",
            f"공급가액: {result['supply_total']:,}원\n"
            f"부가세(10%): {result['vat_total']:,}원\n"
            f"총 청구금액: {result['grand_total']:,}원",
            parent=self,
        )

    def _register_template(self):
        fp = filedialog.askopenfilename(
            title="성적서/사진대장 템플릿 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")],
            parent=self,
        )
        if not fp:
            return
        client_id = self._selected_client_id() or DEFAULT_CLIENT_KEY
        mode = self.tpl_mode_var.get()
        try:
            register_template(fp, mode, client_id=client_id, profile=self.profile)
        except Exception as err:
            messagebox.showerror("등록 오류", str(err), parent=self)
            return
        self._refresh_contracts_and_templates()

    def _open_selected_template(self):
        sel = self.tpl_tree.selection()
        if not sel:
            return
        cid, mode = sel[0].split("::", 1)
        entries = list_templates(client_id=cid, profile=self.profile)
        target = next((e for e in entries if e.get("mode") == mode), None)
        if target and os.path.isfile(target.get("path", "")):
            try:
                os.startfile(target["path"])
            except Exception as err:
                messagebox.showerror("열기 오류", str(err), parent=self)

    def _delete_selected_template(self):
        sel = self.tpl_tree.selection()
        if not sel:
            return
        cid, mode = sel[0].split("::", 1)
        remove_template(mode, client_id=cid, profile=self.profile)
        self._refresh_contracts_and_templates()

