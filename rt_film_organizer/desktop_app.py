from __future__ import annotations

import sys
from pathlib import Path

import cv2
import easyocr
from PySide6.QtCore import QObject, QPoint, QRect, Qt, QThread, Signal, QSettings
from PySide6.QtGui import QAction, QColor, QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QProgressBar, QPushButton,
    QSplitter, QStatusBar, QTableWidget, QTableWidgetItem, QToolBar,
    QVBoxLayout, QWidget, QRubberBand, QDialog, QInputDialog, QCheckBox, QMenu, QToolButton,
)
from report_link import load_reports
from report_dialog import ReportDialog

from core import (
    FilmRecord, assign_sequences, clamp_rect, crop_image, detect_film_rect, export_records,
    imread, imwrite_jpg, list_images, parse_marking, prepare_for_ocr,
    recognize_film_section, render_filename, unique_path,
    save_project, load_project,
)


APP_DIR = Path(__file__).resolve().parent
MODEL_DIR = APP_DIR / "models"
UI_SETTINGS_PATH = APP_DIR.parent / ".work-sync" / "rt-film-ui.ini"


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
                cropped = crop_image(imread(record.source), record.crop_rect or self.rect)
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
                if record.report_item:
                    differences = [label for field, label in (("line_no", "라인"), ("joint_no", "JOINT"), ("welder_no", "용접사"))
                                   if parsed[field] and parsed[field] != getattr(record, field)]
                    if detected_section and record.section and detected_section != record.section:
                        differences.append("구간")
                    record.status = "확인필요(OCR불일치: " + ",".join(differences) + ")" if differences else "보고서 배정·OCR 검토필요"
                    self.progress.emit(position, len(self.indices), index)
                    continue
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
            assign_sequences([r for r in self.records if not r.report_item])
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
    columns = ["원본", "라인번호", "JOINT", "용접사", "구간", "구간판정", "순번", "저장 파일명", "상태", "OCR 원문", "보고서번호", "보고서 항목", "구경", "검사일", "적용 영역"]
    section_options = ("", "(1-2)", "(2-3)", "(3-1)")

    def __init__(self):
        super().__init__()
        self.records: list[FilmRecord] = []
        self.source_photos = []
        self.rect = None
        self.initial_rect = None
        self.processing = False
        self.selection_base = None
        self.selection_record = None
        self.show_cropped = True
        self.ocr_thread = None
        UI_SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
        self.path_settings = QSettings(str(UI_SETTINGS_PATH), QSettings.IniFormat)
        self.setWindowTitle("RT 필름 영상화 정리기")
        self.resize(1540, 900)
        self.setMinimumSize(1120, 850)
        self.build_ui()
        self.apply_style()

    def build_ui(self):
        toolbar = QToolBar("주요 기능")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        actions = [
            ("① 사진 불러오기", self.load_images),
            ("보고서 순서 배정", self.link_report),
            ("작업 저장", self.save_work), ("작업 열기", self.open_work),
        ]
        for text, callback in actions:
            action = QAction(text, self); action.triggered.connect(callback); toolbar.addAction(action)
        toolbar.addSeparator()
        toggle = QAction("원본/크롭 보기", self); toggle.triggered.connect(self.toggle_preview); toolbar.addAction(toggle)
        self.more_menu = QMenu(self)
        for text, callback in [("OCR 분석 (보조 검증)", self.start_ocr),
                               ("현재 영역에서 더 자르기", self.refine_target_crop),
                               ("공통 자르기 되돌리기 (최초 지정)", self.restore_initial_crop),
                               ("원본 이름으로 잘라 저장", self.batch_crop)]:
            action = self.more_menu.addAction(text); action.triggered.connect(callback)
        self.more_menu.addSeparator()
        self.filename_settings_action = self.more_menu.addAction("파일명 설정")
        self.filename_settings_action.setCheckable(True)
        self.filename_settings_action.toggled.connect(lambda visible: self.filename_settings.setVisible(visible))
        more_button = QToolButton(); more_button.setText("더보기")
        more_button.setMenu(self.more_menu); more_button.setPopupMode(QToolButton.InstantPopup)
        toolbar.addWidget(more_button)

        central = QWidget(); outer = QVBoxLayout(central); outer.setContentsMargins(16, 14, 16, 10); outer.setSpacing(10)
        self.setCentralWidget(central)

        title = QLabel("RT 필름 영상화 정리기")
        title.setObjectName("title")
        subtitle = QLabel("사진 불러오기 → 보고서 배정 → 영역 지정 → 확인 후 저장")
        subtitle.setObjectName("subtitle")
        outer.addWidget(title); outer.addWidget(subtitle)

        paths = QFrame(); paths.setObjectName("card"); path_layout = QVBoxLayout(paths)
        self.input_edit = QLineEdit(str(self.path_settings.value("paths/input", "")))
        self.output_edit = QLineEdit(str(self.path_settings.value("paths/output", "")))
        self.input_edit.setPlaceholderText("찾아보기에서 이 PC의 사진 폴더를 선택하세요.")
        self.input_edit.editingFinished.connect(self.save_paths)
        self.output_edit.editingFinished.connect(self.save_paths)
        path_layout.addLayout(self.path_row("입력 폴더", self.input_edit, self.choose_input))
        path_layout.addLayout(self.path_row("출력 폴더", self.output_edit, self.choose_output))
        self.filename_settings = QFrame()
        name_row = QHBoxLayout(self.filename_settings); name_row.setContentsMargins(0,0,0,0)
        name_row.addWidget(QLabel("파일명 규칙"), 0)
        self.template_combo = QComboBox(); self.template_combo.setEditable(True)
        self.template_combo.addItems(["{joint}{section}", "{joint}-{seq}", "{line}_{joint}_{section}-{seq}", "{section}_{line}_{joint}_{welder}-{seq}"])
        self.template_combo.currentTextChanged.connect(self.refresh_names)
        name_row.addWidget(self.template_combo, 1); name_row.addWidget(QLabel("치환자: {report} {line} {joint} {welder} {section} {seq} {size} {date}"))
        path_layout.addWidget(self.filename_settings); self.filename_settings.hide(); outer.addWidget(paths)

        self.edit_controls = QFrame(); self.edit_controls.setObjectName("card")
        edit_layout = QHBoxLayout(self.edit_controls)
        edit_layout.addWidget(QLabel("② 영역 편집 · 적용 대상"))
        self.crop_target = QComboBox(); self.crop_target.addItems(["전체 사진 (공통)", "선택 사진 한 장"])
        self.crop_target.currentIndexChanged.connect(lambda _: self.update_preview())
        self.crop_target.setToolTip("공통 영역을 바꿔도 개별 지정 사진의 영역은 유지됩니다.")
        edit_layout.addWidget(self.crop_target)
        for text, callback in [("영역 지정", self.start_target_crop), ("자동 검출", self.detect_target_crop),
                               ("개별 지정 해제", self.reset_individual_crop)]:
            button = QPushButton(text); button.clicked.connect(callback); edit_layout.addWidget(button)
            if text == "개별 지정 해제": button.setToolTip("선택 사진의 개별 영역을 해제하고 공통 영역으로 돌아갑니다.")
        edit_layout.addStretch(); outer.addWidget(self.edit_controls)

        splitter = QSplitter(Qt.Horizontal)
        left = QFrame(); left.setObjectName("card"); left_layout = QVBoxLayout(left)
        left_layout.addWidget(QLabel("사진 미리보기")); self.preview = ImageLabel()
        self.preview.selection_made.connect(self.manual_crop_selected)
        left_layout.addWidget(self.preview, 1)
        self.image_info = QLabel("-"); self.image_info.setObjectName("hint"); left_layout.addWidget(self.image_info)
        splitter.addWidget(left)

        right = QFrame(); right.setObjectName("card"); right_layout = QVBoxLayout(right)
        heading = QHBoxLayout(); heading.addWidget(QLabel("③ 사진별 내용 확인 · 셀을 더블클릭하여 수정"), 1)
        self.details_toggle = QCheckBox("OCR·보고서 상세")
        self.details_toggle.toggled.connect(self.toggle_details); heading.addWidget(self.details_toggle)
        right_layout.addLayout(heading)
        self.table = QTableWidget(0, len(self.columns)); self.table.setHorizontalHeaderLabels(self.columns)
        self.table.setSelectionBehavior(QTableWidget.SelectRows); self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False); self.table.itemSelectionChanged.connect(self.update_preview)
        self.table.itemChanged.connect(self.table_changed)
        header = self.table.horizontalHeader(); header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch); header.setSectionResizeMode(7, QHeaderView.Stretch)
        self.toggle_details(False)
        right_layout.addWidget(self.table, 1); splitter.addWidget(right); splitter.setSizes([520, 1000]); outer.addWidget(splitter, 1)

        self.save_controls = QFrame(); save_row = QHBoxLayout(self.save_controls)
        save_row.addWidget(QLabel("④ 결과 저장 · 편집 내용은 저장 버튼을 눌러야 파일에 반영됩니다."), 1)
        for text, callback in [("확인 후 저장", self.save_outputs)]:
            button = QPushButton(text); button.clicked.connect(callback); save_row.addWidget(button)
        outer.addWidget(self.save_controls)
        progress_row = QHBoxLayout(); self.progress = QProgressBar(); self.progress.setVisible(False)
        self.status_label = QLabel("입력 폴더를 선택하고 사진을 불러오세요.")
        self.status_label.setWordWrap(True)
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
        if folder:
            old_input = self.input_edit.text().strip()
            automatic_output = str(Path(old_input) / "영상화_결과") if old_input else ""
            if not self.output_edit.text().strip() or self.output_edit.text() == automatic_output:
                self.output_edit.setText(str(Path(folder) / "영상화_결과"))
            self.input_edit.setText(folder)
            self.save_paths()

    def toggle_details(self, visible):
        for col in (3, 5, 6, 9, 10, 11, 12, 13): self.table.setColumnHidden(col, not visible)

    def set_common_crop(self, rect):
        self.rect = tuple(rect)
        if self.initial_rect is None: self.initial_rect = self.rect

    def start_target_crop(self):
        row = self.table.currentRow()
        if self.crop_target.currentIndex() == 1 and not 0 <= row < len(self.records):
            return self.error("표에서 사진 한 장을 선택하세요.")
        self.start_manual_crop()
        if self.records and self.crop_target.currentIndex() == 1:
            self.selection_record = self.records[row]
            self.status_label.setText("원본에서 남길 영역을 드래그하세요. 선택 사진 한 장에만 적용됩니다.")

    def refine_target_crop(self):
        if self.crop_target.currentIndex() == 1: self.start_individual_crop()
        else: self.start_additional_crop()

    def detect_target_crop(self):
        if self.crop_target.currentIndex() == 0: return self.detect_crop()
        row = self.table.currentRow()
        if not 0 <= row < len(self.records): return self.error("표에서 사진 한 장을 선택하세요.")
        try:
            self.records[row].crop_rect = detect_film_rect(imread(self.records[row].source))
            self.show_cropped = True; self.update_preview()
            self.status_label.setText("선택 사진 한 장의 영역을 자동 검출했습니다.")
        except Exception as exc: self.error(str(exc))

    def restore_initial_crop(self):
        if self.initial_rect is None: return self.error("아직 공통 영역을 지정하지 않았습니다.")
        self.rect = self.initial_rect; self.show_cropped = True; self.update_preview()
        self.status_label.setText("공통 영역을 최초 지정 상태로 복원했습니다. 개별 지정 사진은 유지됩니다.")

    def save_work(self):
        if not self.records: return self.error("먼저 사진을 불러오세요.")
        path, _ = QFileDialog.getSaveFileName(self, "자르기·보고서 배정 작업 저장", "영상화_작업.rtfilm.json", "영상화 작업 (*.rtfilm.json)")
        if not path: return
        try:
            save_project(path, self.records, common_rect=self.rect, initial_rect=self.initial_rect,
                         source_photos=self.source_photos, input_dir=self.input_edit.text(),
                         output_dir=self.output_edit.text(), template=self.template_combo.currentText())
            self.status_label.setText(f"작업 저장 완료 · {path}")
        except Exception as exc: self.error(str(exc))

    def open_work(self):
        path, _ = QFileDialog.getOpenFileName(self, "영상화 작업 열기", "", "영상화 작업 (*.rtfilm.json)")
        if not path: return
        try:
            data = load_project(path)
            self.records = data['records']; self.source_photos = data['source_photos']
            self.rect = data['common_rect']; self.initial_rect = data['initial_rect']
            self.input_edit.setText(data.get('input_dir', '')); self.output_edit.setText(data.get('output_dir', ''))
            self.template_combo.setCurrentText(data.get('template', '{joint}{section}'))
            self.show_cropped = True; self.populate_table()
            if self.records: self.table.selectRow(0)
            else: self.preview.clear(); self.image_info.setText("-")
            self.save_paths(); self.status_label.setText(f"작업 복원 완료 · 사진 {len(self.records)}장 · {path}")
        except Exception as exc: self.error(str(exc))

    def choose_output(self):
        folder = QFileDialog.getExistingDirectory(self, "출력 폴더", self.output_edit.text() or self.input_edit.text())
        if folder:
            self.output_edit.setText(folder)
            self.save_paths()

    def save_paths(self):
        self.path_settings.setValue("paths/input", self.input_edit.text().strip())
        self.path_settings.setValue("paths/output", self.output_edit.text().strip())
        self.path_settings.sync()

    def closeEvent(self, event):
        if self.processing or (self.ocr_thread and self.ocr_thread.isRunning()):
            self.error("처리 완료 후 창을 닫아주세요.")
            event.ignore()
            return
        self.save_paths()
        super().closeEvent(event)

    def load_images(self):
        try:
            paths = list_images(self.input_edit.text().strip())
            if not paths: raise ValueError("선택한 폴더에 이미지가 없습니다.")
            self.source_photos = [str(path) for path in paths]
            self.records = [FilmRecord(source=str(path)) for path in paths]; self.rect = None; self.initial_rect = None
            assign_sequences(self.records)
            if not self.output_edit.text(): self.output_edit.setText(str(Path(self.input_edit.text()) / "영상화_결과"))
            self.save_paths()
            self.populate_table(); self.table.selectRow(0)
            self.status_label.setText(f"사진 {len(paths)}장을 불러왔습니다. 필름영역 검출을 실행하세요.")
        except Exception as exc: self.error(str(exc))

    def link_report(self):
        if self.ocr_thread and self.ocr_thread.isRunning():
            return self.error("OCR 완료 후 보고서를 배정하세요.")
        if not self.source_photos:
            return self.error("먼저 사진을 불러오세요.")
        default = APP_DIR.parent / "home/src/report_apps/ndt_report/data/report_issue_history.json"
        path, _ = QFileDialog.getOpenFileName(self, "발행 이력 또는 RT 미리보기 JSON 선택", str(default), "보고서 데이터 (*.json)")
        if not path: return
        try:
            reports = load_reports(path)
            dialog = ReportDialog(reports, self.source_photos, self)
            if dialog.exec() != QDialog.Accepted: return
            records = dialog.result_records
            if any(not r.report_no for r in records):
                number, ok = QInputDialog.getText(self, "보고서번호", "저장 파일명에 사용할 보고서번호")
                if not ok or not number.strip(): return
                for record in records:
                    if not record.report_no: record.report_no = number.strip()
            crop_overrides = {r.source: r.crop_rect for r in self.records}
            for record in records: record.crop_rect = crop_overrides.get(record.source)
            self.records = records
            self.template_combo.setCurrentText("{joint}{section}")
            self.populate_table(); self.table.selectRow(0)
            self.status_label.setText(f"보고서 순서로 {len(records)}장 배정 · 사진과 항목을 확인하고 구간을 선택한 뒤 저장하세요.")
        except Exception as exc: self.error(str(exc))

    def detect_crop(self):
        if not self.records: return self.error("먼저 사진을 불러오세요.")
        try:
            image = imread(self.records[0].source); self.set_common_crop(detect_film_rect(image)); self.show_cropped = True
            x, y, w, h = self.rect; self.update_preview()
            text = f"고정 크롭 X={x}, Y={y}, {w:,}×{h:,} px"
            self.status_label.setText(text + " · 공통 영역 적용 (개별 지정 사진 유지)")
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

    def start_additional_crop(self):
        if not self.records or self.rect is None:
            return self.error("먼저 필름영역을 검출하거나 영역을 직접 지정하세요.")
        self.show_cropped = True
        self.update_preview()
        row = max(0, self.table.currentRow())
        try:
            image = imread(self.records[row].source)
            self.selection_base = clamp_rect(self.rect, image.shape[1], image.shape[0])
            self.preview.set_cv_image(crop_image(image, self.selection_base))
            if self.preview.start_selection():
                self.status_label.setText("현재 크롭 사진에서 남길 영역을 드래그하세요. 전체 사진에 적용됩니다. 이후 일괄 자르기를 다시 눌러 저장하세요.")
        except Exception as exc:
            self.error(str(exc))

    def start_individual_crop(self):
        row = self.table.currentRow()
        if not 0 <= row < len(self.records):
            return self.error("표에서 자를 사진 한 장을 선택하세요.")
        record = self.records[row]
        self.show_cropped = True
        self.update_preview()
        try:
            image = imread(record.source)
            base = record.crop_rect or self.rect or (0, 0, image.shape[1], image.shape[0])
            self.selection_base = clamp_rect(base, image.shape[1], image.shape[0])
            self.selection_record = record
            if self.preview.start_selection():
                self.status_label.setText(f"{Path(record.source).name} 한 장에서 남길 영역을 드래그하세요. 다른 사진의 영역은 유지됩니다.")
        except Exception as exc: self.error(str(exc))

    def reset_individual_crop(self):
        row = self.table.currentRow()
        if not 0 <= row < len(self.records):
            return self.error("표에서 공통 영역으로 되돌릴 사진을 선택하세요.")
        record = self.records[row]
        previous = record.crop_rect
        record.crop_rect = None
        self.show_cropped = True
        self.update_preview()
        if previous is None:
            self.status_label.setText("이 사진은 개별 지정 영역이 없어 복귀 전후가 같습니다. 공통 영역은 ‘영역 직접 지정’으로 다시 설정할 수 있습니다.")
        elif self.rect is None:
            self.status_label.setText("개별 영역을 해제했습니다. 공통 영역이 없어 불러온 원본 전체를 표시합니다.")
        else:
            old_w, old_h = previous[2:]
            new_w, new_h = self.rect[2:]
            self.status_label.setText(f"{Path(record.source).name}: 개별 {old_w:,}×{old_h:,} → 공통 {new_w:,}×{new_h:,} px · 저장 파일에 반영하려면 다시 저장하세요.")

    def manual_crop_selected(self, rect):
        try:
            image = imread(self.records[self.table.currentRow() if self.table.currentRow() >= 0 else 0].source)
            x, y, w, h = rect
            if self.selection_base is not None:
                base_x, base_y, base_w, base_h = self.selection_base
                x, y, w, h = clamp_rect(rect, base_w, base_h)
                x, y = base_x + x, base_y + y
            x = max(0, min(x, image.shape[1] - 1))
            y = max(0, min(y, image.shape[0] - 1))
            w = max(1, min(w, image.shape[1] - x))
            h = max(1, min(h, image.shape[0] - y))
            individual = self.selection_record
            if individual is not None:
                individual.crop_rect = (x, y, w, h)
            else:
                self.set_common_crop((x, y, w, h))
            self.show_cropped = True
            self.update_preview()
            self.status_label.setText(
                f"수동 크롭 X={x}, Y={y}, {w:,}×{h:,} px · {'선택 사진만 적용' if individual is not None else '공통 영역 적용 (개별 지정 사진 유지)'}"
            )
        except Exception as exc:
            self.error(str(exc))

    def start_ocr(self):
        if not self.records: return self.error("먼저 사진을 불러오세요.")
        if self.rect is None and any(r.crop_rect is None for r in self.records):
            try: self.set_common_crop(detect_film_rect(imread(self.records[0].source)))
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
        if self.rect is None and any(r.crop_rect is None for r in self.records):
            try:
                self.set_common_crop(detect_film_rect(imread(self.records[0].source)))
                self.show_cropped = True
                self.update_preview()
            except Exception as exc:
                return self.error(str(exc))
        output_root = self.output_edit.text().strip()
        if not output_root:
            return self.error("출력 폴더를 지정하세요.")

        crop_dir = Path(output_root) / "일괄자르기"
        crop_dir.mkdir(parents=True, exist_ok=True)
        self.set_busy(True)
        self.progress.setVisible(True)
        self.progress.setRange(0, len(self.records))
        written = []
        try:
            for index, record in enumerate(self.records, 1):
                image = imread(record.source)
                cropped = crop_image(image, record.crop_rect or self.rect)
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
                f"공통·개별 지정 영역으로 {len(written)}장을 잘랐습니다.\n"
                f"원본 파일은 변경하지 않았습니다.\n\n저장 위치:\n{crop_dir}",
            )
        except Exception as exc:
            self.error(str(exc))
        finally:
            self.set_busy(False)

    def ocr_progress(self, current, total, index):
        self.progress.setValue(current); self.status_label.setText(f"OCR 분석 중 {current}/{total} · {Path(self.records[index].source).name}")
        self.update_row(index)

    def ocr_complete(self):
        self.populate_table(); self.set_busy(False); self.status_label.setText("OCR 완료. 추출값을 확인·수정한 뒤 저장하세요.")

    def ocr_failed(self, message): self.set_busy(False); self.error(message)

    def set_busy(self, busy):
        self.processing = busy
        self.progress.setVisible(busy); self.table.setEnabled(not busy)
        for toolbar in self.findChildren(QToolBar): toolbar.setEnabled(not busy)
        for widget in (self.input_edit, self.output_edit, self.template_combo): widget.setEnabled(not busy)
        for widget in (self.edit_controls, self.save_controls, self.details_toggle): widget.setEnabled(not busy)
        for button in self.centralWidget().findChildren(QPushButton): button.setEnabled(not busy)

    def populate_table(self):
        self.table.blockSignals(True); self.table.setRowCount(len(self.records))
        for row in range(len(self.records)): self.update_row(row, blocked=True)
        self.table.blockSignals(False)

    def update_row(self, row, blocked=False):
        if not blocked: self.table.blockSignals(True)
        record = self.records[row]; values = [Path(record.source).name, record.line_no, record.joint_no, record.welder_no,
            record.section, record.section_source, str(record.sequence), render_filename(self.template_combo.currentText(), record), record.status, record.raw_text,
            record.report_no, record.report_item, record.size, record.exam_date,
            self.crop_description(record)]
        for col, value in enumerate(values):
            item = QTableWidgetItem(value)
            if col in (0, 4, 5, 7, 8, 9, 11, 14): item.setFlags(item.flags() & ~Qt.ItemIsEditable)
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
        mapping = {1:"line_no", 2:"joint_no", 3:"welder_no", 6:"sequence", 10:"report_no", 12:"size", 13:"exam_date"}
        if col in mapping:
            value = item.text().strip()
            if col == 6:
                try: value = int(value)
                except ValueError: value = 1
            setattr(self.records[row], mapping[col], value)
            self.records[row].status = "수동확인"; self.update_row(row)

    def refresh_names(self):
        if self.records: self.populate_table()

    def crop_description(self, record):
        rect = record.crop_rect or self.rect
        if rect is None: return "미지정"
        return f"{'개별' if record.crop_rect else '공통'} · {rect[2]}×{rect[3]}"

    def update_preview(self):
        self.preview.cancel_selection()
        self.selection_base = None
        self.selection_record = None
        if not self.records: return
        self.table.blockSignals(True)
        for index, record in enumerate(self.records):
            item = self.table.item(index, 14)
            if item: item.setText(self.crop_description(record))
        self.table.blockSignals(False)
        row = self.table.currentRow(); row = row if row >= 0 else 0
        try:
            image = imread(self.records[row].source)
            original_w, original_h = image.shape[1], image.shape[0]
            effective_rect = self.records[row].crop_rect or self.rect
            if self.show_cropped and effective_rect: image = crop_image(image, effective_rect)
            self.preview.set_cv_image(image)
            self.image_info.setText(f"{Path(self.records[row].source).name} · 원본 {original_w:,}×{original_h:,} px\n현재 {image.shape[1]:,}×{image.shape[0]:,} px · {'개별 크롭' if self.show_cropped and self.records[row].crop_rect else '공통 크롭' if self.show_cropped and effective_rect else '원본'}")
        except Exception as exc: self.image_info.setText(str(exc))

    def toggle_preview(self): self.show_cropped = not self.show_cropped; self.update_preview()

    def save_outputs(self):
        if self.ocr_thread and self.ocr_thread.isRunning(): return self.error("OCR 완료 후 저장하세요.")
        if not self.records or any(not (r.crop_rect or self.rect) for r in self.records): return self.error("사진을 불러오고 필름영역을 검출하세요.")
        if not self.output_edit.text().strip(): return self.error("출력 폴더를 지정하세요.")
        if any(r.report_item for r in self.records):
            if any(not r.section or not r.line_no or not r.joint_no or not r.report_no or r.sequence < 1 for r in self.records):
                return self.error("보고서 배정 사진의 보고서번호·라인·JOINT·구간·순번을 모두 확인하세요. 보수구간은 직접 선택해야 합니다.")
            mismatches = sum("OCR불일치" in r.status for r in self.records)
            if QMessageBox.question(self, "보고서와 사진 최종 확인", f"{len(self.records)}장의 사진과 보고서 항목 연결을 모두 확인했나요?\nOCR 불일치 표시: {mismatches}장\n확인한 배정과 현재 파일명으로 저장합니다.", QMessageBox.Yes|QMessageBox.No, QMessageBox.No) != QMessageBox.Yes: return
        if any(not r.joint_no for r in self.records):
            if QMessageBox.question(self, "미확인 항목", "JOINT 미확인 사진이 있습니다. 검토본을 계속 저장할까요?") != QMessageBox.Yes: return
        self.set_busy(True)
        try:
            written = export_records(self.records, self.output_edit.text().strip(), self.rect, self.template_combo.currentText())
            QMessageBox.information(self, "저장 완료", f"{len(written)}장을 별도 폴더에 저장했습니다.\n원본은 변경하지 않았습니다.\n\n{self.output_edit.text()}")
            self.status_label.setText(f"저장 완료: {len(written)}장")
        except Exception as exc: self.error(str(exc))
        finally: self.set_busy(False)

    def error(self, message): QMessageBox.critical(self, "오류", message)


def main():
    app = QApplication(sys.argv); app.setStyle("Fusion")
    window = MainWindow(); window.show(); sys.exit(app.exec())


if __name__ == "__main__": main()

