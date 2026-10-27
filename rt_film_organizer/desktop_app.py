from __future__ import annotations

import sys
from pathlib import Path

import cv2
import easyocr
from PySide6.QtCore import QObject, QPoint, QRect, Qt, QThread, Signal
from PySide6.QtGui import QAction, QColor, QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QProgressBar, QPushButton,
    QSplitter, QStatusBar, QTableWidget, QTableWidgetItem, QToolBar,
    QVBoxLayout, QWidget, QRubberBand,
)

from core import (
    FilmRecord, assign_sequences, crop_image, detect_film_rect, export_records,
    imread, imwrite_jpg, list_images, parse_marking, prepare_for_ocr,
    recognize_film_section, render_filename, unique_path,
)


APP_DIR = Path(__file__).resolve().parent
MODEL_DIR = APP_DIR / "models"
DEFAULT_INPUT = r"F:\내 드라이브\01_업무_및_회사\Office\착공\2026년 중앙지사 열수송관 비파괴검사 단가계약\영상화장비사진\예제"


class OcrWorker(QObject):
    progress = Signal(int, int, object)
    completed = Signal()
    failed = Signal(str)

    def __init__(self, records, rect, indices):
        super().__init__()
        self.records, self.rect, self.indices = records, rect, indices

    def run(self):
        try:
            reader = easyocr.Reader(["en"], gpu=False, verbose=False,
                                    model_storage_directory=str(MODEL_DIR))
            for position, index in enumerate(self.indices, 1):
                record = self.records[index]
                cropped = crop_image(imread(record.source), self.rect)
                detected_section = recognize_film_section(reader, cropped)
                candidates = []
                for variant in prepare_for_ocr(cropped):
                    found = reader.readtext(
                        variant, detail=1, paragraph=False,
                        allowlist="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-/",
                    )
                    text = " ".join(item[1] for item in found)
                    confidence = sum(float(item[2]) for item in found) / max(len(found), 1)
                    candidates.append((len(text) * max(confidence, .15), text, confidence))
                _, text, confidence = max(candidates, default=(0, "", 0))
                parsed = parse_marking(text)
                record.raw_text, record.confidence = text, confidence
                record.line_no, record.joint_no = parsed["line_no"], parsed["joint_no"]
                record.welder_no, record.section = parsed["welder_no"], parsed["section"]
                if detected_section:
                    record.section = detected_section
                    record.section_source = "납마킹 OCR"
                else:
                    record.section_source = ""
                missing = [v for v in (record.line_no, record.joint_no, record.welder_no, record.section) if not v]
                record.status = "확인" if not missing else "검토필요"
                self.progress.emit(position, len(self.indices), index)
            assign_sequences(self.records)
            self.completed.emit()
        except Exception as exc:
            self.failed.emit(str(exc))


class ImageLabel(QLabel):
    selection_made = Signal(tuple)

    def __init__(self):
        super().__init__("사진을 불러오세요")
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumSize(360, 300)
        self.setStyleSheet("background:#111820;color:#aab7c4;border-radius:8px;")
        self.original_pixmap = None
        self.selection_enabled = False
        self.drag_origin = QPoint()
        self.rubber_band = QRubberBand(QRubberBand.Rectangle, self)
        self.setCursor(Qt.ArrowCursor)

    def set_cv_image(self, image):
        rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        h, w, channels = rgb.shape
        qimage = QImage(rgb.data, w, h, channels * w, QImage.Format_RGB888).copy()
        self.original_pixmap = QPixmap.fromImage(qimage)
        self.fit_image()

    def fit_image(self):
        if self.original_pixmap:
            target_size = self.contentsRect().size()
            if target_size.width() <= 0 or target_size.height() <= 0:
                return
            self.setPixmap(self.original_pixmap.scaled(
                target_size,
                Qt.KeepAspectRatio, Qt.SmoothTransformation,
            ))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.rubber_band.hide()
        self.fit_image()

    def start_selection(self):
        if not self.original_pixmap:
            return False
        self.selection_enabled = True
        self.rubber_band.hide()
        self.setCursor(Qt.CrossCursor)
        return True

    def cancel_selection(self):
        self.selection_enabled = False
        self.rubber_band.hide()
        self.setCursor(Qt.ArrowCursor)

    def displayed_image_rect(self):
        pixmap = self.pixmap()
        if not pixmap:
            return QRect()
        area = self.contentsRect()
        x = area.x() + (area.width() - pixmap.width()) // 2
        y = area.y() + (area.height() - pixmap.height()) // 2
        return QRect(x, y, pixmap.width(), pixmap.height())

    def mousePressEvent(self, event):
        if self.selection_enabled and event.button() == Qt.LeftButton:
            image_rect = self.displayed_image_rect()
            if image_rect.contains(event.position().toPoint()):
                self.drag_origin = event.position().toPoint()
                self.rubber_band.setGeometry(QRect(self.drag_origin, self.drag_origin))
                self.rubber_band.show()
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.selection_enabled and self.rubber_band.isVisible():
            current = event.position().toPoint()
            selection = QRect(self.drag_origin, current).normalized().intersected(self.displayed_image_rect())
            self.rubber_band.setGeometry(selection)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.selection_enabled and event.button() == Qt.LeftButton and self.rubber_band.isVisible():
            selected = self.rubber_band.geometry().intersected(self.displayed_image_rect())
            image_rect = self.displayed_image_rect()
            self.cancel_selection()
            if selected.width() >= 10 and selected.height() >= 10 and image_rect.isValid():
                source_w, source_h = self.original_pixmap.width(), self.original_pixmap.height()
                shown = self.pixmap()
                scale_x = source_w / shown.width()
                scale_y = source_h / shown.height()
                x = round((selected.x() - image_rect.x()) * scale_x)
                y = round((selected.y() - image_rect.y()) * scale_y)
                width = round(selected.width() * scale_x)
                height = round(selected.height() * scale_y)
                self.selection_made.emit((x, y, width, height))
            event.accept()
            return
        super().mouseReleaseEvent(event)


class MainWindow(QMainWindow):
    columns = ["원본", "라인번호", "JOINT", "용접사", "구간", "구간판정", "순번", "저장 파일명", "상태", "OCR 원문"]
    section_options = ("", "(1-2)", "(2-3)", "(3-1)")

    def __init__(self):
        super().__init__()
        self.records: list[FilmRecord] = []
        self.rect = None
        self.show_cropped = True
        self.ocr_thread = None
        self.setWindowTitle("RT 필름 영상화 정리기")
        self.resize(1540, 900)
        self.setMinimumSize(1120, 700)
        self.build_ui()
        self.apply_style()

    def build_ui(self):
        toolbar = QToolBar("주요 기능")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        actions = [
            ("사진 불러오기", self.load_images), ("필름영역 검출", self.detect_crop),
            ("영역 직접 지정", self.start_manual_crop), ("일괄 자르기", self.batch_crop),
            ("OCR 일괄분석", self.start_ocr),
            ("OCR명으로 저장", self.save_outputs),
        ]
        for text, callback in actions:
            action = QAction(text, self); action.triggered.connect(callback); toolbar.addAction(action)
        toolbar.addSeparator()
        toggle = QAction("원본/크롭 보기", self); toggle.triggered.connect(self.toggle_preview); toolbar.addAction(toggle)

        central = QWidget(); outer = QVBoxLayout(central); outer.setContentsMargins(16, 14, 16, 10); outer.setSpacing(10)
        self.setCentralWidget(central)

        title = QLabel("RT 필름 영상화 정리기")
        title.setObjectName("title")
        subtitle = QLabel("고정좌표 일괄 크롭 · 납마킹 OCR · 검토 후 JPG 저장 · 원본 보존")
        subtitle.setObjectName("subtitle")
        outer.addWidget(title); outer.addWidget(subtitle)

        paths = QFrame(); paths.setObjectName("card"); path_layout = QVBoxLayout(paths)
        self.input_edit = QLineEdit(DEFAULT_INPUT)
        self.output_edit = QLineEdit()
        path_layout.addLayout(self.path_row("입력 폴더", self.input_edit, self.choose_input))
        path_layout.addLayout(self.path_row("출력 폴더", self.output_edit, self.choose_output))
        name_row = QHBoxLayout(); name_row.addWidget(QLabel("파일명 규칙"), 0)
        self.template_combo = QComboBox(); self.template_combo.setEditable(True)
        self.template_combo.addItems(["{joint}-{seq}", "{line}_{joint}_{section}-{seq}", "{section}_{line}_{joint}_{welder}-{seq}"])
        self.template_combo.currentTextChanged.connect(self.refresh_names)
        name_row.addWidget(self.template_combo, 1); name_row.addWidget(QLabel("치환자: {line} {joint} {welder} {section} {seq}"))
        path_layout.addLayout(name_row); outer.addWidget(paths)

        splitter = QSplitter(Qt.Horizontal)
        left = QFrame(); left.setObjectName("card"); left_layout = QVBoxLayout(left)
        left_layout.addWidget(QLabel("사진 미리보기")); self.preview = ImageLabel()
        self.preview.selection_made.connect(self.manual_crop_selected)
        left_layout.addWidget(self.preview, 1)
        self.image_info = QLabel("-"); self.image_info.setObjectName("hint"); left_layout.addWidget(self.image_info)
        splitter.addWidget(left)

        right = QFrame(); right.setObjectName("card"); right_layout = QVBoxLayout(right)
        right_layout.addWidget(QLabel("OCR 결과 - 셀을 더블클릭하여 수정하세요"))
        self.table = QTableWidget(0, len(self.columns)); self.table.setHorizontalHeaderLabels(self.columns)
        self.table.setSelectionBehavior(QTableWidget.SelectRows); self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False); self.table.itemSelectionChanged.connect(self.update_preview)
        self.table.itemChanged.connect(self.table_changed)
        header = self.table.horizontalHeader(); header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch); header.setSectionResizeMode(7, QHeaderView.Stretch)
        right_layout.addWidget(self.table, 1); splitter.addWidget(right); splitter.setSizes([520, 1000]); outer.addWidget(splitter, 1)

        progress_row = QHBoxLayout(); self.progress = QProgressBar(); self.progress.setVisible(False)
        self.status_label = QLabel("입력 폴더를 선택하고 사진을 불러오세요.")
        progress_row.addWidget(self.status_label, 1); progress_row.addWidget(self.progress, 0); outer.addLayout(progress_row)
        self.setStatusBar(QStatusBar())

    def path_row(self, label, edit, callback):
        row = QHBoxLayout(); text = QLabel(label); text.setFixedWidth(82); row.addWidget(text); row.addWidget(edit, 1)
        button = QPushButton("찾아보기"); button.clicked.connect(callback); row.addWidget(button); return row

    def apply_style(self):
        self.setStyleSheet("""
        QMainWindow{background:#f3f6f9} QToolBar{background:#17324d;padding:7px;spacing:7px;border:0}
        QToolButton{color:white;padding:8px 13px;border-radius:5px;background:#24547b;font-weight:600}
        QToolButton:hover{background:#2f78ae} QLabel#title{font-size:24px;font-weight:700;color:#17324d}
        QLabel#subtitle,QLabel#hint{color:#617283} QFrame#card{background:white;border:1px solid #d8e1e8;border-radius:9px}
        QLineEdit,QComboBox{padding:7px;border:1px solid #b9c7d3;border-radius:5px;background:white}
        QPushButton{padding:7px 12px;border:1px solid #aebdca;border-radius:5px;background:white}
        QPushButton:hover{background:#eaf2f8} QTableWidget{border:1px solid #d8e1e8;gridline-color:#e5ebf0;alternate-background-color:#f7fafc}
        QHeaderView::section{background:#e8f0f6;padding:7px;border:0;border-right:1px solid #d2dce5;font-weight:600}
        QProgressBar{min-width:260px;border:1px solid #b9c7d3;border-radius:5px;text-align:center} QProgressBar::chunk{background:#2374ab}
        """)

    def choose_input(self):
        folder = QFileDialog.getExistingDirectory(self, "입력 사진 폴더", self.input_edit.text())
        if folder: self.input_edit.setText(folder)

    def choose_output(self):
        folder = QFileDialog.getExistingDirectory(self, "출력 폴더", self.output_edit.text() or self.input_edit.text())
        if folder: self.output_edit.setText(folder)

    def load_images(self):
        try:
            paths = list_images(self.input_edit.text().strip())
            if not paths: raise ValueError("선택한 폴더에 이미지가 없습니다.")
            self.records = [FilmRecord(source=str(path)) for path in paths]; self.rect = None
            assign_sequences(self.records)
            if not self.output_edit.text(): self.output_edit.setText(str(Path(self.input_edit.text()) / "영상화_결과"))
            self.populate_table(); self.table.selectRow(0)
            self.status_label.setText(f"사진 {len(paths)}장을 불러왔습니다. 필름영역 검출을 실행하세요.")
        except Exception as exc: self.error(str(exc))

    def detect_crop(self):
        if not self.records: return self.error("먼저 사진을 불러오세요.")
        try:
            image = imread(self.records[0].source); self.rect = detect_film_rect(image); self.show_cropped = True
            x, y, w, h = self.rect; self.update_preview()
            text = f"고정 크롭 X={x}, Y={y}, {w:,}×{h:,} px"
            if w < 2048 or h < 1536:
                self.status_label.setText(text + " · 과업 최소 해상도 미충족")
                QMessageBox.warning(self, "크롭 해상도 부족", f"검출 영역은 {w:,}×{h:,}픽셀입니다.\n과업 최소 2,048×1,536을 충족하지 않습니다.\n카메라를 더 가까이 배치하여 재촬영하는 것을 권장합니다.")
            else: self.status_label.setText(text + " · 해상도 충족")
        except Exception as exc: self.error(str(exc))

    def start_manual_crop(self):
        if not self.records:
            return self.error("먼저 사진을 불러오세요.")
        self.show_cropped = False
        self.update_preview()
        if self.preview.start_selection():
            self.status_label.setText(
                "미리보기에서 자를 영역의 왼쪽 위부터 오른쪽 아래까지 마우스로 드래그하세요."
            )

    def manual_crop_selected(self, rect):
        try:
            image = imread(self.records[self.table.currentRow() if self.table.currentRow() >= 0 else 0].source)
            x, y, w, h = rect
            x = max(0, min(x, image.shape[1] - 1))
            y = max(0, min(y, image.shape[0] - 1))
            w = max(1, min(w, image.shape[1] - x))
            h = max(1, min(h, image.shape[0] - y))
            self.rect = (x, y, w, h)
            self.show_cropped = True
            self.update_preview()
            compliance = "해상도 충족" if w >= 2048 and h >= 1536 else "과업 최소 해상도 미충족"
            self.status_label.setText(
                f"수동 고정 크롭 X={x}, Y={y}, {w:,}×{h:,} px · {compliance} · 전체 사진에 동일 적용"
            )
        except Exception as exc:
            self.error(str(exc))

    def start_ocr(self):
        if not self.records: return self.error("먼저 사진을 불러오세요.")
        if self.rect is None:
            try: self.rect = detect_film_rect(imread(self.records[0].source))
            except Exception as exc: return self.error(str(exc))
        self.set_busy(True); self.progress.setRange(0, len(self.records)); self.progress.setValue(0)
        self.status_label.setText("OCR 모델을 준비하고 있습니다...")
        self.ocr_thread = QThread(); self.worker = OcrWorker(self.records, self.rect, list(range(len(self.records))))
        self.worker.moveToThread(self.ocr_thread); self.ocr_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.ocr_progress); self.worker.completed.connect(self.ocr_complete); self.worker.failed.connect(self.ocr_failed)
        self.worker.completed.connect(self.ocr_thread.quit); self.worker.failed.connect(self.ocr_thread.quit)
        self.ocr_thread.start()

    def batch_crop(self):
        """Crop every loaded image with the single fixed rectangle, without OCR or renaming."""
        if not self.records:
            return self.error("먼저 사진을 불러오세요.")
        if self.rect is None:
            try:
                self.rect = detect_film_rect(imread(self.records[0].source))
                self.show_cropped = True
                self.update_preview()
            except Exception as exc:
                return self.error(str(exc))
        output_root = self.output_edit.text().strip()
        if not output_root:
            return self.error("출력 폴더를 지정하세요.")

        crop_dir = Path(output_root) / "일괄자르기"
        crop_dir.mkdir(parents=True, exist_ok=True)
        self.progress.setVisible(True)
        self.progress.setRange(0, len(self.records))
        written = []
        try:
            for index, record in enumerate(self.records, 1):
                image = imread(record.source)
                cropped = crop_image(image, self.rect)
                destination = unique_path(crop_dir / f"{Path(record.source).stem}.JPG")
                imwrite_jpg(destination, cropped)
                written.append(destination)
                self.progress.setValue(index)
                self.status_label.setText(
                    f"일괄 자르기 {index}/{len(self.records)} · {Path(record.source).name}"
                )
                QApplication.processEvents()
            self.status_label.setText(
                f"일괄 자르기 완료: {len(written)}장 · {crop_dir}"
            )
            QMessageBox.information(
                self,
                "일괄 자르기 완료",
                f"동일한 고정 좌표로 {len(written)}장을 잘랐습니다.\n"
                f"원본 파일은 변경하지 않았습니다.\n\n저장 위치:\n{crop_dir}",
            )
        except Exception as exc:
            self.error(str(exc))
        finally:
            self.progress.setVisible(False)

    def ocr_progress(self, current, total, index):
        self.progress.setValue(current); self.status_label.setText(f"OCR 분석 중 {current}/{total} · {Path(self.records[index].source).name}")
        self.update_row(index)

    def ocr_complete(self):
        self.populate_table(); self.set_busy(False); self.status_label.setText("OCR 완료. 추출값을 확인·수정한 뒤 저장하세요.")

    def ocr_failed(self, message): self.set_busy(False); self.error(message)

    def set_busy(self, busy):
        self.progress.setVisible(busy); self.table.setEnabled(not busy)

    def populate_table(self):
        self.table.blockSignals(True); self.table.setRowCount(len(self.records))
        for row in range(len(self.records)): self.update_row(row, blocked=True)
        self.table.blockSignals(False)

    def update_row(self, row, blocked=False):
        if not blocked: self.table.blockSignals(True)
        record = self.records[row]; values = [Path(record.source).name, record.line_no, record.joint_no, record.welder_no,
            record.section, record.section_source, str(record.sequence), render_filename(self.template_combo.currentText(), record), record.status, record.raw_text]
        for col, value in enumerate(values):
            item = QTableWidgetItem(value)
            if col in (0, 4, 5, 7, 8, 9): item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            if col == 5 and ("확인필요" in value or value in ("순서추정", "촬영순서")):
                item.setBackground(QColor("#fff2cc"))
            if record.status == "검토필요": item.setBackground(QColor("#fff2cc"))
            self.table.setItem(row, col, item)
        section_combo = QComboBox()
        section_combo.addItems(self.section_options)
        if record.section in self.section_options:
            section_combo.setCurrentText(record.section)
        else:
            section_combo.setCurrentIndex(0)
        section_combo.setToolTip("납마킹을 확인하여 구간을 선택하세요.")
        section_combo.currentTextChanged.connect(
            lambda value, target_row=row: self.section_combo_changed(target_row, value)
        )
        self.table.setCellWidget(row, 4, section_combo)
        if not blocked: self.table.blockSignals(False)

    def section_combo_changed(self, row, value):
        if not (0 <= row < len(self.records)):
            return
        record = self.records[row]
        record.section = value
        record.section_source = "수동확인" if value else ""
        record.status = "수동확인" if value else "검토필요"
        self.table.blockSignals(True)
        self.table.item(row, 4).setText(value)
        self.table.item(row, 5).setText(record.section_source)
        self.table.item(row, 7).setText(render_filename(self.template_combo.currentText(), record))
        self.table.item(row, 8).setText(record.status)
        self.table.blockSignals(False)

    def table_changed(self, item):
        row, col = item.row(), item.column()
        if row >= len(self.records): return
        mapping = {1:"line_no", 2:"joint_no", 3:"welder_no", 6:"sequence"}
        if col in mapping:
            value = item.text().strip()
            if col == 6:
                try: value = int(value)
                except ValueError: value = 1
            setattr(self.records[row], mapping[col], value)
            self.records[row].status = "수동확인"; self.update_row(row)

    def refresh_names(self):
        if self.records: self.populate_table()

    def update_preview(self):
        if not self.records: return
        row = self.table.currentRow(); row = row if row >= 0 else 0
        try:
            image = imread(self.records[row].source)
            if self.show_cropped and self.rect: image = crop_image(image, self.rect)
            self.preview.set_cv_image(image)
            self.image_info.setText(f"{Path(self.records[row].source).name} · {image.shape[1]:,}×{image.shape[0]:,} px · {'크롭' if self.show_cropped and self.rect else '원본'}")
        except Exception as exc: self.image_info.setText(str(exc))

    def toggle_preview(self): self.show_cropped = not self.show_cropped; self.update_preview()

    def save_outputs(self):
        if not self.records or self.rect is None: return self.error("사진을 불러오고 필름영역을 검출하세요.")
        if not self.output_edit.text().strip(): return self.error("출력 폴더를 지정하세요.")
        if any(not r.joint_no for r in self.records):
            if QMessageBox.question(self, "미확인 항목", "JOINT 미확인 사진이 있습니다. 검토본을 계속 저장할까요?") != QMessageBox.Yes: return
        if self.rect[2] < 2048 or self.rect[3] < 1536:
            if QMessageBox.question(self, "해상도 기준 미충족", "크롭 해상도가 과업 기준보다 작습니다. 공식 납품본이 아닌 검토본으로 저장할까요?") != QMessageBox.Yes: return
        try:
            written = export_records(self.records, self.output_edit.text().strip(), self.rect, self.template_combo.currentText())
            QMessageBox.information(self, "저장 완료", f"{len(written)}장을 별도 폴더에 저장했습니다.\n원본은 변경하지 않았습니다.\n\n{self.output_edit.text()}")
            self.status_label.setText(f"저장 완료: {len(written)}장")
        except Exception as exc: self.error(str(exc))

    def error(self, message): QMessageBox.critical(self, "오류", message)


def main():
    app = QApplication(sys.argv); app.setStyle("Fusion")
    window = MainWindow(); window.show(); sys.exit(app.exec())


if __name__ == "__main__": main()

