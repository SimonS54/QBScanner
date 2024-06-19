import re
import spacy
from fuzzywuzzy import fuzz
from PyQt5 import QtCore, QtGui, QtWidgets
import pyautogui
import pyperclip
import pygetwindow as gw
import time
import os
import sys
import ctypes
import mss
import cv2
import numpy as np
from google.cloud import vision
import io
import keyboard
import winsound

nlp = spacy.load("en_core_web_sm")

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "service_account_key.json"
client = vision.ImageAnnotatorClient()

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, ' '.join(sys.argv), None, 1)

CURRENT_DIR = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
IMAGE_PATH = os.path.join(CURRENT_DIR, "image.png")

class OCRWorker(QtCore.QThread):
    resultReady = QtCore.pyqtSignal(str)
    notifyReady = QtCore.pyqtSignal(str)

    def __init__(self, region=None):
        super().__init__()
        self.running = False
        self.screenshot_path = os.path.join(CURRENT_DIR, "screenshot.png")
        self.region = region if region else (350, 200, 1000, 300)

    def run(self):
        self.running = True
        while self.running:
            try:
                screenshot = pyautogui.screenshot(region=self.region)
                screenshot.save(self.screenshot_path)

                text = self.image_to_text(self.screenshot_path)

                self.process_text(text)

                if self.error and self.qb_link and self.product:
                    browser_link = self.get_current_browser_url()
                    if browser_link:
                        if self.qb_id:
                            message = f"/qbissue product: {self.product} ticket_link: {browser_link} qb_link: {self.qb_link} issue: {self.error} qb_id: {self.qb_id}"
                        else:
                            message = f"/qbissue product: {self.product} ticket_link: {browser_link} qb_link: {self.qb_link} issue: {self.error}"
                        self.resultReady.emit(message)
                        self.notifyReady.emit("All information gathered successfully!")
                        self.stop()

                time.sleep(0.2)
            except Exception as e:
                print(f"Error during OCR processing: {e}")

    def image_to_text(self, image_path):
        with io.open(image_path, 'rb') as image_file:
            content = image_file.read()
        image = vision.Image(content=content)

        response = client.text_detection(image=image)
        texts = response.text_annotations
        if texts:
            return texts[0].description
        else:
            return ""

    def process_text(self, text):
        self.error = self.extract_issue(text) or self.error
        self.qb_link = self.extract_qb_link(text) or self.qb_link
        self.qb_id = self.extract_qb_id(text) or self.qb_id
        self.product = self.extract_product(text) or self.product

    def stop(self):
        self.running = False

    def extract_issue(self, text):
        issue_marker = "ISSUE"
        text_lines = text.split('\n')
        issue_index = -1
        for i, line in enumerate(text_lines):
            if issue_marker in line:
                issue_index = i
                break

        if issue_index != -1 and issue_index + 1 < len(text_lines):
            return text_lines[issue_index + 1].strip()
        return None

    def extract_qb_link(self, text):
        match = re.search(r'https?://\S+', text)
        return match.group(0) if match else None

    def extract_qb_id(self, text):
        match = re.search(r'\b[0-9a-fA-F]{8}\b', text)
        return match.group(0) if match else None

    def extract_product(self, text):
        products = {
            "R6 Full": ["r6 full", "rainbow six full", "rainbow full"],
            "R6 Lite": ["r6 lite", "rainbow six lite", "rainbow lite", "lite"],
            "XDefiant": ["xdefiant", "xd", "defiant"]
        }

        text = text.lower()
        best_match = ("r6_full", 0)

        for product, aliases in products.items():
            for alias in aliases:
                match_ratio = fuzz.partial_ratio(alias, text)
                if match_ratio > best_match[1]:
                    best_match = (product, match_ratio)

        if best_match[1] > 70:
            return best_match[0]
        else:
            return "r6_full"

    def get_current_browser_url(self):
        try:
            browsers = {
                'Opera': ('ctrl', 'l'),
                'Chrome': ('ctrl', 'l'),
                'Firefox': ('ctrl', 'l'),
                'Microsoft Edge': ('ctrl', 'l'),
            }

            window = None
            for browser_name, (hotkey1, hotkey2) in browsers.items():
                for w in gw.getWindowsWithTitle(browser_name):
                    if browser_name.lower() in w.title.lower():
                        window = w
                        break
                if window:
                    break

            if window is None:
                raise Exception("Supported browser window not found")

            window.activate()
            time.sleep(0.1)
            pyautogui.hotkey(hotkey1, hotkey2)
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.1)

            url = pyperclip.paste()
            return url

        except Exception as e:
            print(f"Error retrieving browser URL: {e}")
            return None

class CustomTitleBar(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(CustomTitleBar, self).__init__(parent)
        self.parent = parent
        self.setAutoFillBackground(True)
        self.setBackgroundRole(QtGui.QPalette.Window)
        self.initUI()

    def initUI(self):
        self.setFixedHeight(40)
        layout = QtWidgets.QHBoxLayout()
        layout.setContentsMargins(10, 0, 0, 0)
        self.setLayout(layout)

        self.titleLabel = QtWidgets.QLabel("QuickBuild Auto Scanner")
        self.titleLabel.setStyleSheet("color: white; font: bold 14px;")
        layout.addWidget(self.titleLabel)

        layout.addStretch()

        self.minimizeButton = QtWidgets.QPushButton("-")
        self.minimizeButton.setFixedSize(40, 40)
        self.minimizeButton.setStyleSheet(
            "QPushButton { background-color: transparent; border-image: url(minimize.png); color: white; font-weight: bold; }"
            "QPushButton:hover { background-color: #4a4a4a; }"
        )
        self.minimizeButton.clicked.connect(self.parent.showMinimized)
        layout.addWidget(self.minimizeButton)

        self.closeButton = QtWidgets.QPushButton("X")
        self.closeButton.setFixedSize(40, 40)
        self.closeButton.setStyleSheet(
            "QPushButton { background-color: transparent; border-image: url(close.png); color: white; font-weight: bold; }"
            "QPushButton:hover { background-color: #ff4a4a; }"
        )
        self.closeButton.clicked.connect(self.parent.close)
        layout.addWidget(self.closeButton)

        palette = self.palette()
        palette.setColor(QtGui.QPalette.Background, QtGui.QColor('#181c34'))
        self.setPalette(palette)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.old_pos = event.globalPos()

    def mouseMoveEvent(self, event):
        if self.old_pos is not None:
            delta = event.globalPos() - self.old_pos
            self.parent.move(self.parent.pos() + delta)
            self.old_pos = event.globalPos()

    def mouseReleaseEvent(self, event):
        self.old_pos = None

class App(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.worker = None
        self.hotkey = None
        self.scan_region = (350, 200, 1000, 300)

    def initUI(self):
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.setGeometry(100, 100, 400, 300)
        self.center()

        self.titleBar = CustomTitleBar(self)
        centralWidget = QtWidgets.QWidget()
        self.setCentralWidget(centralWidget)
        layout = QtWidgets.QVBoxLayout(centralWidget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.titleBar)

        content = QtWidgets.QWidget()
        contentLayout = QtWidgets.QVBoxLayout(content)
        layout.addWidget(content)

        palette = self.palette()
        palette.setColor(QtGui.QPalette.Background, QtGui.QColor('#181c34'))
        self.setPalette(palette)

        appIcon = QtGui.QIcon(os.path.join(CURRENT_DIR, 'icon.ico'))
        self.setWindowIcon(appIcon)

        pixmap = QtGui.QPixmap(IMAGE_PATH)
        pixmap = pixmap.scaled(100, 100, QtCore.Qt.KeepAspectRatio)

        self.logoLabel = QtWidgets.QLabel()
        self.logoLabel.setPixmap(pixmap)
        contentLayout.addWidget(self.logoLabel, alignment=QtCore.Qt.AlignCenter)

        self.usernameLabel = QtWidgets.QLabel("Hotkey:")
        self.usernameLabel.setStyleSheet("color: white;")
        contentLayout.addWidget(self.usernameLabel, alignment=QtCore.Qt.AlignCenter)

        self.hotkeyField = QtWidgets.QLineEdit()
        contentLayout.addWidget(self.hotkeyField, alignment=QtCore.Qt.AlignCenter)

        buttonLayout = QtWidgets.QHBoxLayout()
        contentLayout.addLayout(buttonLayout)

        self.setButton = QtWidgets.QPushButton("Set!")
        self.setButton.clicked.connect(self.set_hotkey)
        buttonLayout.addWidget(self.setButton, alignment=QtCore.Qt.AlignCenter)

        self.regionButton = QtWidgets.QPushButton("Set Scan Area")
        self.regionButton.clicked.connect(self.set_scan_area)
        buttonLayout.addWidget(self.regionButton, alignment=QtCore.Qt.AlignCenter)

        self.create_tray_icon()
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        centerPoint = QtWidgets.QApplication.desktop().screenGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def set_hotkey(self):
        new_hotkey = self.hotkeyField.text().lower()
        if new_hotkey:
            if self.hotkey:
                keyboard.remove_hotkey(self.hotkey)
            self.hotkey = new_hotkey
            keyboard.add_hotkey(self.hotkey, self.toggle_scan)
            QtWidgets.QMessageBox.information(self, "Hotkey Set", f'Hotkey "{self.hotkey}" has been set!')

    def set_scan_area(self):
        try:
            region = self.select_region()
            if region:
                self.scan_region = region
                QtWidgets.QMessageBox.information(self, "Scan Area Set", f'Scan area has been set to {self.scan_region}!')
            else:
                QtWidgets.QMessageBox.warning(self, "Error", 'Failed to set scan area: No region selected')
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f'Failed to set scan area: {e}')

    def select_region(self):
        try:
            QtWidgets.QMessageBox.information(self, "Instruction", "Press Escape to finalize the capture area selection")

            with mss.mss() as sct:
                monitor = sct.monitors[1]
                screenshot = sct.grab(monitor)
                img = np.array(screenshot)

                img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

                cv2.namedWindow("Select Region", cv2.WND_PROP_FULLSCREEN)
                cv2.setWindowProperty("Select Region", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
                r = cv2.selectROI("Select Region", img, fromCenter=False, showCrosshair=True)
                cv2.destroyAllWindows()

                if r[2] > 0 and r[3] > 0:
                    x, y, w, h = r
                    return (x, y, w, h)
                return None
        except Exception as e:
            print(f"Error during region selection: {e}")
            return None

    def toggle_scan(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
            self.worker = None
            self.play_stop_sound()
        else:
            self.worker = OCRWorker(region=self.scan_region)
            self.worker.resultReady.connect(self.handle_result)
            self.worker.notifyReady.connect(self.notify_user)
            self.worker.start()
            self.play_start_sound()

    def handle_result(self, message):
        print(message)
        pyperclip.copy(message)

    def notify_user(self, notification_message):
        QtWidgets.QMessageBox.information(self, "Scanning Complete", notification_message)

    def create_tray_icon(self):
        self.trayIcon = QtWidgets.QSystemTrayIcon(self)
        self.trayIcon.setIcon(QtGui.QIcon(IMAGE_PATH))
        self.trayIcon.setToolTip("QuickBuild Auto Scanner")

        showAction = QtWidgets.QAction("Show", self)
        showAction.triggered.connect(self.showNormal)

        quitAction = QtWidgets.QAction("Quit", self)
        quitAction.triggered.connect(self.close)

        trayMenu = QtWidgets.QMenu()
        trayMenu.addAction(showAction)
        trayMenu.addAction(quitAction)

        self.trayIcon.setContextMenu(trayMenu)
        self.trayIcon.show()

    def play_start_sound(self):
        winsound.Beep(500, 200)

    def play_stop_sound(self):
        winsound.Beep(300, 200)

if __name__ == '__main__':
    if not is_admin():
        run_as_admin()
        sys.exit()
    else:
        if os.name == 'nt':
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
        
        app = QtWidgets.QApplication(sys.argv)
        window = App()
        sys.exit(app.exec_())