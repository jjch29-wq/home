"""Native Windows UI for ISO Drawer (PySide6)."""
from __future__ import annotations

import sys
from PySide6.QtCore import QPointF, Qt, Signal
from PySide6.QtGui import QAction, QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDockWidget, QFileDialog, QFormLayout, QInputDialog,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton, QToolBar, QVBoxLayout,
    QWidget,
)

# Keep imports explicit so PyInstaller detects the local module reliably.
from core import Point, Project, export_dxf, export_pdf, load_project, save_project, snap_iso


class DrawingCanvas(QWidget):
    changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.project = Project()
        self.scale = 1.0
        self.offset = QPointF(80, 80)
        self.cursor_world = None
        self.pan_anchor = None
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setCursor(Qt.CrossCursor)

    def to_world(self, p):
        return ((p.x() - self.offset.x()) / self.scale, (p.y() - self.offset.y()) / self.scale)

    def to_screen(self, p):
        return QPointF(p.x * self.scale + self.offset.x(), p.y * self.scale + self.offset.y())

    def mousePressEvent(self, event):
        if event.button() == Qt.MiddleButton:
            self.pan_anchor = (event.position(), QPointF(self.offset))
            self.setCursor(Qt.ClosedHandCursor)
            return
        if event.button() == Qt.RightButton:
            self.cursor_world = None
            self.update()
            return
        if event.button() != Qt.LeftButton:
            return
        x, y = self.to_world(event.position())
        if not self.project.points:
            self.project.points.append(Point(x, y))
            self.changed.emit(); self.update(); return
        last = self.project.points[-1]
        x, y = snap_iso((last.x, last.y), (x, y))
        if ((x-last.x)**2 + (y-last.y)**2)**0.5 < 5 / self.scale:
            return
        value, ok = QInputDialog.getDouble(self, "실제 길이", "이 구간의 실제 길이 (mm):", 1000, 0.1, 100000000, 1)
        if ok:
            self.project.points.append(Point(x, y, value))
            self.changed.emit(); self.update()

    def mouseMoveEvent(self, event):
        if self.pan_anchor:
            start, original = self.pan_anchor
            self.offset = original + event.position() - start
        else:
            self.cursor_world = self.to_world(event.position())
        self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MiddleButton:
            self.pan_anchor = None
            self.setCursor(Qt.CrossCursor)

    def wheelEvent(self, event):
        before = self.to_world(event.position())
        self.scale = max(.1, min(10, self.scale * (1.15 if event.angleDelta().y() > 0 else 1/1.15)))
        self.offset = QPointF(event.position().x() - before[0]*self.scale, event.position().y() - before[1]*self.scale)
        self.update()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.cursor_world = None; self.update()
        elif event.matches(event.StandardKey.Undo):
            self.undo()
        else:
            super().keyPressEvent(event)

    def undo(self):
        if self.project.points:
            self.project.points.pop(); self.changed.emit(); self.update()

    def fit(self):
        if not self.project.points:
            self.scale, self.offset = 1, QPointF(80, 80)
        else:
            xs = [p.x for p in self.project.points]; ys = [p.y for p in self.project.points]
            self.scale = min((self.width()-140)/max(max(xs)-min(xs), 1), (self.height()-140)/max(max(ys)-min(ys), 1))
            self.scale = max(.1, min(10, self.scale))
            self.offset = QPointF(70-min(xs)*self.scale, 70-min(ys)*self.scale)
        self.update()

    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing); p.fillRect(self.rect(), QColor("#17232d"))
        step = 50*self.scale
        if step >= 12:
            p.setPen(QPen(QColor("#263743"), 1)); x = self.offset.x() % step
            while x < self.width(): p.drawLine(QPointF(x,0), QPointF(x,self.height())); x += step
            y = self.offset.y() % step
            while y < self.height(): p.drawLine(QPointF(0,y), QPointF(self.width(),y)); y += step
        p.setFont(QFont("Segoe UI", 9))
        for i in range(1, len(self.project.points)):
            a, b = self.to_screen(self.project.points[i-1]), self.to_screen(self.project.points[i])
            p.setPen(QPen(QColor("#64e6db"), 3)); p.drawLine(a,b)
            p.setPen(QColor("white")); p.drawText((a+b)/2 + QPointF(4,-10), f"{self.project.points[i].actual_length:g} mm")
        for point in self.project.points:
            q = self.to_screen(point); p.setPen(QPen(QColor("#ff7b7b"),2)); p.drawEllipse(q,4,4)
            if point.component != "NONE":
                p.setPen(QColor("#ffd166")); p.drawText(q+QPointF(7,15), point.component)
        if self.project.points and self.cursor_world:
            last=self.project.points[-1]; q=snap_iso((last.x,last.y),self.cursor_world)
            pen=QPen(QColor("#ffbf69"),2,Qt.DashLine); p.setPen(pen); p.drawLine(self.to_screen(last),self.to_screen(Point(*q)))


class MainWindow(QMainWindow):
    COMPONENTS = ("NONE","ELBOW_90","ELBOW_45","TEE","GATE_VALVE","BALL_VALVE","CHECK_VALVE","FLANGE","REDUCER","WELD")

    def __init__(self):
        super().__init__(); self.setWindowTitle("ISO Drawer - 배관 아이소메트릭"); self.resize(1280,780)
        self.canvas=DrawingCanvas(); self.setCentralWidget(self.canvas); self.canvas.changed.connect(self.refresh_status)
        self._toolbar(); self._properties(); self.statusBar().showMessage("포인트 0개 · 구간 0개")

    def action(self, text, slot, shortcut=None):
        a=QAction(text,self); a.triggered.connect(slot)
        if shortcut: a.setShortcut(shortcut)
        self.toolbar.addAction(a)

    def _toolbar(self):
        self.toolbar=QToolBar("도구"); self.toolbar.setMovable(False); self.addToolBar(self.toolbar)
        for args in [("새 도면",self.new,"Ctrl+N"),("열기",self.open,"Ctrl+O"),("저장",self.save,"Ctrl+S"),("DXF 출력",self.dxf,None),("PDF 출력",self.pdf,None),("실행 취소",self.canvas.undo,"Ctrl+Z"),("전체 맞춤",self.canvas.fit,"F")]: self.action(*args)

    def _properties(self):
        dock=QDockWidget("도면 정보 및 부속",self); dock.setMinimumWidth(250); box=QWidget(); lay=QVBoxLayout(box); form=QFormLayout()
        self.line=QLineEdit("LINE-001"); self.size_edit=QLineEdit('4"'); self.spec=QLineEdit(); self.comp=QComboBox(); self.comp.addItems(self.COMPONENTS)
        form.addRow("라인 번호",self.line); form.addRow("배관 구경",self.size_edit); form.addRow("SPEC",self.spec); form.addRow("끝점 부속",self.comp); lay.addLayout(form)
        apply=QPushButton("끝점에 부속 적용"); apply.clicked.connect(self.apply_component); lay.addWidget(apply)
        help=QLabel("<b>작업 방법</b><br><br>1. 시작점을 클릭합니다.<br>2. 다음 방향을 클릭합니다.<br>3. 실제 길이(mm)를 입력합니다.<br>4. 필요할 때 부속을 적용합니다.<br><br>우클릭/ESC: 입력 종료<br>휠: 확대·축소<br>가운데 드래그: 화면 이동"); help.setWordWrap(True); help.setAlignment(Qt.AlignTop); lay.addWidget(help,1)
        dock.setWidget(box); self.addDockWidget(Qt.RightDockWidgetArea,dock)

    def sync(self):
        q=self.canvas.project; q.line_no=self.line.text(); q.size=self.size_edit.text(); q.spec=self.spec.text()
    def refresh_status(self):
        n=len(self.canvas.project.points); self.statusBar().showMessage(f"포인트 {n}개 · 구간 {max(0,n-1)}개")
    def new(self):
        if QMessageBox.question(self,"새 도면","현재 도면을 지우고 새로 시작할까요?")==QMessageBox.Yes:
            self.canvas.project=Project(); self.canvas.cursor_world=None; self.canvas.update(); self.refresh_status()
    def apply_component(self):
        if self.canvas.project.points:
            self.canvas.project.points[-1].component=self.comp.currentText(); self.canvas.update()
    def save(self):
        self.sync(); path,_=QFileDialog.getSaveFileName(self,"프로젝트 저장",self.line.text()+".json","ISO 프로젝트 (*.json)")
        if path: save_project(self.canvas.project,path)
    def open(self):
        path,_=QFileDialog.getOpenFileName(self,"프로젝트 열기","","ISO 프로젝트 (*.json)")
        if not path:return
        try:
            q=load_project(path); self.canvas.project=q; self.line.setText(q.line_no); self.size_edit.setText(q.size); self.spec.setText(q.spec); self.canvas.fit(); self.refresh_status()
        except Exception as e: QMessageBox.critical(self,"열기 실패",str(e))
    def _export(self, ext, fn, title, filt):
        if len(self.canvas.project.points)<2: return QMessageBox.warning(self,"출력 불가","포인트를 두 개 이상 입력하세요.")
        self.sync(); path,_=QFileDialog.getSaveFileName(self,title,self.line.text()+ext,filt)
        if path:
            try: fn(self.canvas.project,path); QMessageBox.information(self,"출력 완료",path)
            except Exception as e: QMessageBox.critical(self,"출력 실패",str(e))
    def dxf(self): self._export(".dxf",export_dxf,"DXF 출력","DXF 도면 (*.dxf)")
    def pdf(self): self._export(".pdf",export_pdf,"PDF 출력","PDF 도면 (*.pdf)")


def main():
    app=QApplication(sys.argv); app.setStyle("Fusion"); window=MainWindow(); window.show(); return app.exec()


if __name__ == "__main__": raise SystemExit(main())
