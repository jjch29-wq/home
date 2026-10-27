from pathlib import Path
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QTableWidget, QTableWidgetItem, QPushButton, QSpinBox, QHeaderView, QMessageBox)
from PySide6.QtCore import Qt
from report_link import build_slots, suggested_count, map_photos


class ReportDialog(QDialog):
    def __init__(self, reports, sources, parent=None):
        super().__init__(parent)
        self.reports, self.sources = reports, sources
        self.result_records = None
        self.slots = []
        self.setWindowTitle("저장된 RT 보고서 → 사진 배정")
        self.resize(1100, 760)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("① 보고서와 촬영 매수를 확인 → ② 순서 배정 → 사진을 보며 누락·보수 배정을 수정하세요."))
        self.report_combo = QComboBox()
        for report in reports:
            self.report_combo.addItem(f"{report.get('report_no') or 'RT 미리보기'} · {report.get('issued_at', '')} · {len(report['items'])}항목")
        layout.addWidget(self.report_combo)
        self.items = QTableWidget(0, 6)
        self.items.setHorizontalHeaderLabels(["순서", "라인번호", "JOINT", "구경", "용접사", "촬영 매수"])
        self.items.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.items.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        layout.addWidget(self.items)
        row = QHBoxLayout()
        button = QPushButton("② 촬영 매수로 순서 배정 / 다시 배정")
        button.clicked.connect(self.generate)
        row.addWidget(button)
        self.summary = QLabel("매수 정보가 없는 원촬영은 3장 제안, 보수는 0장입니다. 반드시 확인하세요.")
        self.summary.setWordWrap(True)
        row.addWidget(self.summary, 1)
        layout.addLayout(row)
        self.mapping = QTableWidget(0, 4)
        self.mapping.setHorizontalHeaderLabels(["보고서 항목", "라인 / JOINT / 구간", "배정할 원본 사진 (누락 선택 가능)", "사진 순번"])
        self.mapping.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.mapping.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        layout.addWidget(self.mapping, 2)
        footer = QHBoxLayout()
        apply = QPushButton("배정 적용 후 메인 화면에서 사진 확인")
        apply.clicked.connect(self.apply_mapping)
        cancel = QPushButton("취소"); cancel.clicked.connect(self.reject)
        footer.addWidget(apply); footer.addWidget(cancel); layout.addLayout(footer)
        self.report_combo.currentIndexChanged.connect(self.show_report)
        self.show_report()

    def show_report(self):
        self.slots = []; self.mapping.setRowCount(0)
        report = self.reports[self.report_combo.currentIndex()]
        self.items.setRowCount(len(report['items']))
        for row, item in enumerate(report['items']):
            for col, value in enumerate([str(row+1), item.get('Dwg',''), item.get('Joint',''), item.get('Size',''), item.get('Welder','')]):
                cell = QTableWidgetItem(str(value)); cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                self.items.setItem(row, col, cell)
            count = QSpinBox(); count.setRange(0,100); count.setValue(suggested_count(item))
            count.valueChanged.connect(self.invalidate)
            self.items.setCellWidget(row, 5, count)

    def invalidate(self):
        self.slots = []; self.mapping.setRowCount(0)
        self.summary.setText("촬영 매수가 변경됐습니다. 순서 배정을 다시 누르세요.")

    def generate(self):
        try:
            report = self.reports[self.report_combo.currentIndex()]
            self.slots = build_slots(report, [self.items.cellWidget(r,5).value() for r in range(self.items.rowCount())])
            self.mapping.setRowCount(len(self.slots))
            for row, slot in enumerate(self.slots):
                for col, text in ((0,slot['report_item']), (1,f"{slot['line_no']} / {slot['joint_no']} / {slot['section'] or '구간 선택 필요'}"), (3,str(slot['sequence']))):
                    cell = QTableWidgetItem(text); cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                    self.mapping.setItem(row,col,cell)
                combo = QComboBox(); combo.addItem("누락 / 배정 안 함", None)
                for index, source in enumerate(self.sources): combo.addItem(Path(source).name,index)
                combo.setCurrentIndex(row+1 if row < len(self.sources) else 0)
                self.mapping.setCellWidget(row,2,combo)
            self.summary.setText(f"보고서 예상 {len(self.slots)}장 / 사진 {len(self.sources)}장 · 중간 누락은 이후 사진을 직접 재배정하세요.")
        except Exception as exc: QMessageBox.warning(self,"배정 확인",str(exc))

    def apply_mapping(self):
        try:
            if not self.slots: raise ValueError("촬영 매수를 확인하고 순서 배정을 먼저 실행하세요.")
            assignments = [self.mapping.cellWidget(r,2).currentData() for r in range(len(self.slots))]
            records = map_photos(self.slots,self.sources,assignments)
            if not records: raise ValueError("배정된 사진이 없습니다.")
            missing = assignments.count(None); extra = len(self.sources)-len(records)
            if QMessageBox.question(self,"배정 확인",f"배정 {len(records)}장 / 누락 {missing}장 / 제외 사진 {extra}장\n제외 사진은 저장하지 않습니다. 이 배정을 적용할까요?", QMessageBox.Yes|QMessageBox.No, QMessageBox.No) != QMessageBox.Yes: return
            self.result_records = records
            self.accept()
        except Exception as exc: QMessageBox.warning(self,"배정 확인",str(exc))
