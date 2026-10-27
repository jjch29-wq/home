from pathlib import Path
import json
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QTableWidget, QTableWidgetItem, QPushButton, QSpinBox, QHeaderView, QMessageBox,
    QListWidget, QFileDialog, QSplitter, QWidget, QInputDialog)
from PySide6.QtCore import Qt
from report_link import ReportBatch, load_reports


class ReportDialog(QDialog):
    def __init__(self, reports, sources, parent=None):
        super().__init__(parent)
        self.reports, self.sources = list(reports), list(sources)
        self.batch = ReportBatch(sources)
        self.result_records = None
        self.slots = []
        self.rendering = False
        self.setWindowTitle('여러 RT 보고서 → 촬영 순서별 사진 배정')
        self.resize(1250, 900)
        layout = QVBoxLayout(self)
        hint = QLabel('① 보고서 추가·순서 변경 → ② 전체 순서 배정 → ③ 보고서별 첫·마지막 사진과 누락 확인')
        hint.setWordWrap(True); layout.addWidget(hint)
        source_row = QHBoxLayout()
        self.report_combo = QComboBox(); self.fill_catalog()
        source_row.addWidget(self.report_combo, 1)
        self.add_button(source_row, '보고서 추가', self.add_report)
        self.add_button(source_row, '다른 이력 / JSON 파일 추가', self.import_reports)
        layout.addLayout(source_row)
        split = QSplitter(Qt.Horizontal)
        queue_panel = QWidget(); queue_layout = QVBoxLayout(queue_panel)
        queue_layout.setContentsMargins(0, 0, 8, 0)
        queue_layout.addWidget(QLabel('보고서 작업 순서 · 선택하면 항목·사진 배정 표시'))
        self.queue = QListWidget(); queue_layout.addWidget(self.queue, 1)
        order_row = QHBoxLayout()
        self.add_button(order_row, '위로', lambda: self.move_report(-1))
        self.add_button(order_row, '아래로', lambda: self.move_report(1))
        self.add_button(order_row, '목록에서 제외', self.remove_report)
        queue_layout.addLayout(order_row)
        self.add_button(queue_layout, '전체 목록 순서로 사진 배정', self.generate_all)
        queue_note = QLabel('순서 이동은 기존 사진 연결을 유지합니다.\n촬영 순서대로 다시 연결하려면 전체 배정을 누르세요.\n목록 제외는 저장된 보고서 원본에 영향을 주지 않습니다.')
        queue_note.setWordWrap(True); queue_layout.addWidget(queue_note)
        split.addWidget(queue_panel)
        detail_panel = QWidget(); detail_layout = QVBoxLayout(detail_panel)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        self.current_label = QLabel('보고서를 추가하세요.'); detail_layout.addWidget(self.current_label)
        self.items = QTableWidget(0, 6)
        self.items.setHorizontalHeaderLabels(['항목', '라인번호', 'JOINT', '구경', '용접사', '촬영 매수'])
        self.items.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.items.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        detail_layout.addWidget(self.items, 1)
        controls = QHBoxLayout(); controls.addWidget(QLabel('이 보고서 시작 사진 번호'))
        self.report_start = QSpinBox(); self.report_start.setRange(1, max(1, len(sources)+1))
        self.report_start.setToolTip('전체 사진 목록에서 1부터 센 번호입니다. 이 보고서만 다시 배정할 때 사용합니다.')
        controls.addWidget(self.report_start)
        self.add_button(controls, '선택 보고서만 다시 배정', self.generate)
        detail_layout.addLayout(controls)
        self.boundary = QLabel('-'); self.boundary.setWordWrap(True); detail_layout.addWidget(self.boundary)
        self.mapping = QTableWidget(0, 4)
        self.mapping.setHorizontalHeaderLabels(['항목', '라인 / JOINT / 구간', '배정 사진 (누락 선택 가능)', '순번'])
        self.mapping.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.mapping.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        detail_layout.addWidget(self.mapping, 2)
        split.addWidget(detail_panel); split.setSizes([340, 870]); layout.addWidget(split, 1)
        self.summary = QLabel('매수 정보가 없는 원촬영은 3장, 보수는 0장 제안입니다. 실제 매수를 확인하세요.')
        self.summary.setWordWrap(True); layout.addWidget(self.summary)
        footer = QHBoxLayout()
        self.add_button(footer, '전체 배정 적용 후 사진 확인', self.apply_mapping)
        self.add_button(footer, '취소', self.reject); layout.addLayout(footer)
        self.queue.currentRowChanged.connect(self.show_report)
        if self.reports: self.add_report()

    @staticmethod
    def add_button(layout, text, callback):
        button = QPushButton(text); button.clicked.connect(callback); layout.addWidget(button)
        return button

    def fill_catalog(self):
        self.report_combo.clear()
        for report in self.reports:
            self.report_combo.addItem(f"{report.get('report_no') or 'RT 미리보기'} · {report.get('issued_at', '')} · {len(report['items'])}항목")

    def import_reports(self):
        paths, _ = QFileDialog.getOpenFileNames(self, '추가할 RT 이력 / 미리보기 JSON', '', '보고서 데이터 (*.json)')
        if not paths: return
        try:
            incoming = [report for path in paths for report in load_reports(path)]
            identities = {r.get('id') or json.dumps(r, sort_keys=True, ensure_ascii=False) for r in self.reports}
            added = 0
            for report in incoming:
                key = report.get('id') or json.dumps(report, sort_keys=True, ensure_ascii=False)
                if key not in identities:
                    self.reports.append(report); identities.add(key); added += 1
            self.fill_catalog()
            if added: self.report_combo.setCurrentIndex(len(self.reports)-added)
            self.summary.setText(f'선택 후보에 {added}개 보고서를 불러왔습니다. 원하는 보고서를 골라 보고서 추가를 누르세요.')
        except Exception as exc: QMessageBox.warning(self, '보고서 불러오기', str(exc))

    def add_report(self):
        index = self.report_combo.currentIndex()
        if index < 0: return
        try:
            self.batch.add(self.reports[index]); self.refresh_queue(len(self.batch.entries)-1)
            self.summary.setText(f'작업 보고서 {len(self.batch.entries)}개 · 순서를 정하고 전체 배정을 실행하세요.')
        except Exception as exc: QMessageBox.warning(self, '보고서 추가', str(exc))

    def queue_text(self, index):
        entry = self.batch.entries[index]
        used = [a for a in entry.assignments if a is not None]
        name = entry.report.get('report_no') or 'RT 미리보기'
        state = '배정 필요' if entry.dirty else f'배정 {len(used)} / 누락 {entry.assignments.count(None)}'
        return f'{index+1}. {name}\n예상 {sum(entry.counts)}장 · {state}'

    def refresh_queue(self, selected=None):
        if selected is None: selected = self.queue.currentRow()
        self.queue.blockSignals(True); self.queue.clear()
        for index in range(len(self.batch.entries)): self.queue.addItem(self.queue_text(index))
        selected = min(max(0, selected), len(self.batch.entries)-1)
        self.queue.setCurrentRow(selected); self.queue.blockSignals(False)
        self.show_report()

    def move_report(self, offset):
        index = self.queue.currentRow()
        if index < 0: return
        self.refresh_queue(self.batch.move(index, offset))
        self.summary.setText('보고서 순서를 변경했습니다. 기존 사진 배정은 유지됩니다. 촬영 순서로 재연결하려면 전체 배정을 누르세요.')

    def remove_report(self):
        index = self.queue.currentRow()
        if index < 0: return
        del self.batch.entries[index]; self.refresh_queue(index)
        self.summary.setText('작업 목록에서만 제외했습니다. 다른 보고서의 사진 배정은 유지됩니다.')

    def show_report(self, *_):
        self.rendering = True
        try:
            self.items.setRowCount(0); self.mapping.setRowCount(0); self.slots = []
            index = self.queue.currentRow()
            if index < 0:
                self.current_label.setText('보고서를 추가하세요.'); self.boundary.setText('-'); return
            entry = self.batch.entries[index]
            self.current_label.setText(f"선택 보고서: {entry.report.get('report_no') or 'RT 미리보기'}")
            self.items.setRowCount(len(entry.report['items']))
            for row, item in enumerate(entry.report['items']):
                for col, value in enumerate([str(row+1), item.get('Dwg',''), item.get('Joint',''), item.get('Size',''), item.get('Welder','')]):
                    cell = QTableWidgetItem(str(value)); cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                    self.items.setItem(row, col, cell)
                count = QSpinBox(); count.setRange(0,100); count.setValue(entry.counts[row])
                count.valueChanged.connect(lambda value, r=row: self.change_count(r, value))
                self.items.setCellWidget(row, 5, count)
            self.report_start.setValue(min(entry.start+1, self.report_start.maximum()))
            if not entry.dirty: self.render_mapping(entry)
            self.update_boundaries()
        finally: self.rendering = False

    def change_count(self, row, value):
        if self.rendering: return
        index = self.queue.currentRow()
        if index < 0: return
        entry = self.batch.entries[index]; entry.counts[row] = value; entry.dirty = True
        self.slots = []; self.mapping.setRowCount(0)
        self.queue.item(index).setText(self.queue_text(index)); self.update_boundaries()
        self.summary.setText('선택 보고서의 촬영 매수가 변경됐습니다. 선택 보고서만 다시 배정하세요. 다른 보고서 연결은 유지됩니다.')

    def render_mapping(self, entry):
        self.slots = entry.slots; self.mapping.setRowCount(len(entry.slots))
        for row, slot in enumerate(entry.slots):
            for col, text in ((0, slot['report_item']), (1, f"{slot['line_no']} / {slot['joint_no']} / {slot['section'] or '구간 선택 필요'}"), (3,str(slot['sequence']))):
                cell = QTableWidgetItem(text); cell.setFlags(cell.flags() & ~Qt.ItemIsEditable); self.mapping.setItem(row,col,cell)
            combo = QComboBox(); combo.addItem('누락 / 배정 안 함', None)
            for index, source in enumerate(self.sources): combo.addItem(f'{index+1}. {Path(source).name}', index)
            assigned = entry.assignments[row]
            combo.setCurrentIndex(0 if assigned is None else assigned+1)
            combo.currentIndexChanged.connect(lambda _, r=row: self.mapping_changed(r))
            self.mapping.setCellWidget(row,2,combo)

    def mapping_changed(self, row):
        if self.rendering: return
        index = self.queue.currentRow()
        if index < 0: return
        entry = self.batch.entries[index]
        entry.assignments[row] = self.mapping.cellWidget(row,2).currentData()
        self.queue.item(index).setText(self.queue_text(index)); self.update_boundaries()

    def update_boundaries(self):
        index = self.queue.currentRow()
        if index < 0: return
        entry = self.batch.entries[index]
        if entry.dirty: self.boundary.setText('촬영 매수를 확인하고 배정하세요. 다른 보고서의 배정에는 영향을 주지 않습니다.'); return
        used = [a for a in entry.assignments if a is not None]
        first = f'{used[0]+1}. {Path(self.sources[used[0]]).name}' if used else '없음'
        last = f'{used[-1]+1}. {Path(self.sources[used[-1]]).name}' if used else '없음'
        self.boundary.setText(f'첫 사진: {first}\n마지막 사진: {last} · 예상 {len(entry.slots)}장 / 배정 {len(used)}장 / 누락 {entry.assignments.count(None)}장')

    def generate(self):
        index = self.queue.currentRow()
        if index < 0: return
        try:
            self.batch.generate(index, self.report_start.value()-1); self.refresh_queue(index)
            self.summary.setText('선택 보고서만 다시 배정했습니다. 다른 보고서의 연결은 유지됩니다. 누락 사진을 직접 조정하세요.')
        except Exception as exc: QMessageBox.warning(self, '사진 배정', str(exc))

    def generate_all(self):
        if not self.batch.entries: return
        if any(entry.slots for entry in self.batch.entries):
            if QMessageBox.question(self,'전체 다시 배정','전체 보고서를 현재 순서대로 처음부터 배정합니다. 기존 수동 사진 연결을 대체할까요?', QMessageBox.Yes|QMessageBox.No, QMessageBox.No) != QMessageBox.Yes: return
        try:
            self.batch.generate_all(); self.refresh_queue()
            self.summary.setText(f'보고서 {len(self.batch.entries)}개 / 예상 {sum(len(e.slots) for e in self.batch.entries)}장 / 전체 사진 {len(self.sources)}장 · 보고서별 첫·마지막 사진을 확인하세요.')
        except Exception as exc: QMessageBox.warning(self, '사진 배정', str(exc))

    def apply_mapping(self):
        try:
            records = self.batch.records()
            if not records: raise ValueError('배정된 사진이 없습니다.')
            for index, entry in enumerate(self.batch.entries):
                if any(not slot['report_no'] for slot in entry.slots):
                    number, ok = QInputDialog.getText(self, '보고서번호 확인', f'작업 순서 {index+1}번 보고서의 번호를 입력하세요.')
                    if not ok or not number.strip(): return
                    for slot in entry.slots:
                        if not slot['report_no']: slot['report_no'] = number.strip()
            records = self.batch.records()
            missing = sum(e.assignments.count(None) for e in self.batch.entries)
            extra = len(self.sources)-len(records)
            if QMessageBox.question(self,'전체 배정 확인',f'보고서 {len(self.batch.entries)}개 · 배정 {len(records)}장 / 누락 {missing}장 / 제외 사진 {extra}장\n보고서별 첫·마지막 사진과 누락을 확인했나요?\n제외 사진은 저장하지 않습니다.', QMessageBox.Yes|QMessageBox.No, QMessageBox.No) != QMessageBox.Yes: return
            self.result_records = records; self.accept()
        except Exception as exc: QMessageBox.warning(self,'배정 확인',str(exc))
