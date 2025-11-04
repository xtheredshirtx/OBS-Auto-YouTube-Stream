import sys
import time
import datetime
import logging
import traceback
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass
from threading import Timer
import ctypes
import platform

import pyautogui
from PyQt5 import QtWidgets, QtGui, QtCore

"""
Automated Task Runner (Pro)
Author: xTheRedShirtx
"""

APP_TITLE = "Automated Task Runner (Pro)"
APP_AUTHOR = "xTheRedShirtx"
OBS_START_HOTKEY = 'up'
OBS_STOP_HOTKEY = 'down'
LOG_FILE = "automation_log.txt"

QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

VK_LCONTROL = 0xA2
VK_ESCAPE   = 0x1B
VK_DELETE   = 0x2E

def is_windows() -> bool:
    return platform.system().lower().startswith("win")

def key_pressed(vk_code: int) -> bool:
    if not is_windows():
        return False
    return (ctypes.windll.user32.GetAsyncKeyState(vk_code) & 0x8000) != 0

logger = logging.getLogger("automation")
logger.setLevel(logging.DEBUG)

_file = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
_file.setLevel(logging.DEBUG)
_file.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
if not any(isinstance(h, RotatingFileHandler) for h in logger.handlers):
    logger.addHandler(_file)

pyautogui.FAILSAFE = True

@dataclass
class ClickPoint:
    name: str
    x: int
    y: int

DEFAULT_POINTS = [
    ClickPoint("Step 1", 3514, 1640),
    ClickPoint("Step 2 (date field)", 1775, 596),
    ClickPoint("Step 3", 1474, 1649),
    ClickPoint("Step 5", 2875, 1640),
    ClickPoint("Step 7", 2674, 1640),
    ClickPoint("Step 8", 2066, 1100),
]

def load_points(settings: QtCore.QSettings) -> list:
    points = []
    size = settings.value("points/count", 0, int)
    if not size:
        return DEFAULT_POINTS.copy()
    for i in range(size):
        name = settings.value(f"points/{i}/name", f"Step {i+1}")
        x = int(settings.value(f"points/{i}/x".format(i=i), 0))
        y = int(settings.value(f"points/{i}/y".format(i=i), 0))
        points.append(ClickPoint(name, x, y))
    return points

def save_points(settings: QtCore.QSettings, points: list):
    settings.setValue("points/count", len(points))
    for i, p in enumerate(points):
        settings.setValue(f"points/{i}/name", p.name)
        settings.setValue(f"points/{i}/x", p.x)
        settings.setValue(f"points/{i}/y", p.y)

class WatchdogTimer:
    def __init__(self, timeout, error_callback):
        self.timeout = timeout
        self.error_callback = error_callback
        self.timer = None

    def reset(self):
        self.cancel()
        self.timer = Timer(self.timeout, self.error_callback)
        self.timer.daemon = True
        self.timer.start()

    def cancel(self):
        if self.timer:
            self.timer.cancel()
            self.timer = None

def clear_obs_broadcast_error(autoretry: bool = True):
    """Dismiss OBS/YouTube 'Live broadcast creation error' dialog if present."""
    try:
        import pygetwindow as gw
        titles = [t for t in gw.getAllTitles() if t]
        match_titles = [t for t in titles if ("Live broadcast creation error" in t) or ("Broadcast creation error" in t) or ("Forbidden" in t)]
        if match_titles:
            w = gw.getWindowsWithTitle(match_titles[0])[0]
            try:
                w.activate()
            except Exception:
                pass
            time.sleep(0.2)
            pyautogui.press("enter")  # OK
            logging.getLogger("automation").info("Dismissed OBS broadcast creation error dialog.")
            if autoretry:
                try:
                    # In case OBS thinks it's still live, send STOP then START
                    pyautogui.press(OBS_STOP_HOTKEY)
                    time.sleep(0.8)
                    pyautogui.press(OBS_START_HOTKEY)
                    logging.getLogger("automation").info("Sent OBS restart hotkeys (DOWN then UP).")
                except Exception as _:
                    logging.getLogger("automation").warning("Failed to send OBS hotkeys.")
            return True
    except Exception:
        # pygetwindow may not be installed or window ops may fail; ignore silently
        return False


def obs_start_stream():
    """Press UP arrow to start streaming."""
    try:
        pyautogui.press("up")
        logger.info("Sent UP arrow to start streaming in OBS.")
    except Exception as e:
        logger.error(f"Failed to send start hotkey: {e}")

def obs_stop_stream():
    """Press DOWN arrow to stop streaming."""
    try:
        pyautogui.press("down")
        logger.info("Sent DOWN arrow to stop streaming in OBS.")
    except Exception as e:
        logger.error(f"Failed to send stop hotkey: {e}")


class AutomationThread(QtCore.QThread):
    log_signal = QtCore.pyqtSignal(str)
    status_signal = QtCore.pyqtSignal(str)
    stop_signal = QtCore.pyqtSignal()
    update_timer_signal = QtCore.pyqtSignal(int)
    error_popup_signal = QtCore.pyqtSignal(str)

    def __init__(self, points: list, long_wait_seconds: int, step_delay: int,
                 max_retries: int, watchdog_seconds: int, step4_wait_sec: int,
                 dry_run: bool):
        super().__init__()
        self.points = points
        self.total_seconds = max(0, int(long_wait_seconds))
        self.step_delay = max(0, int(step_delay))
        self.max_retries = max(1, int(max_retries))
        self.step4_wait_sec = max(0, int(step4_wait_sec))
        self.dry_run = dry_run
        self.is_running = False
        self.watchdog = WatchdogTimer(max(1, int(watchdog_seconds)), self.watchdog_timeout)

    def stop(self):
        """Gracefully stop the automation thread."""
        self.is_running = False

    def watchdog_timeout(self):
        msg = "Error: Script unresponsive. Watchdog timeout."
        logger.error(msg)
        self.log_signal.emit(msg)
        self.status_signal.emit("Status: Error - Watchdog timeout")
        self.stop()

    def safe_sleep_with_interrupt(self, seconds: int):
        for _ in range(seconds):
            if not self.is_running:
                self.log_signal.emit("Automation interrupted during a delay.")
                return False
            time.sleep(1)
        return True

    def execute_click(self, x: int, y: int, description: str) -> bool:
        for attempt in range(1, self.max_retries + 1):
            try:
                self.log_signal.emit(f"[{description}] Attempt {attempt}/{self.max_retries}")
                logger.debug(f"Clicking at ({x}, {y}) - {description} - attempt {attempt}")
                if not self.dry_run:
                    pyautogui.click(x, y)
                else:
                    self.log_signal.emit(f"[DRY RUN] Would click at ({x}, {y})")
                # Pause watchdog during intentional per-step delay to avoid false timeouts
                try:
                    self.watchdog.cancel()
                except Exception:
                    pass
                if not self.safe_sleep_with_interrupt(self.step_delay):
                    return False
                # Resume watchdog after delay
                self.watchdog.reset()
                return True
            except pyautogui.FailSafeException:
                self.log_signal.emit("PyAutoGUI Fail-safe triggered (mouse to top-left). Stopping.")
                logger.exception("PyAutoGUI Fail-safe triggered")
                self.stop()
                return False
            except Exception as e:
                logger.exception(f"Error during {description}: {e}")
                self.log_signal.emit(f"Retrying {description} due to: {e}")
        return False

    def run(self):
        self.is_running = True
        try:
            self.status_signal.emit("Status: Running")
            while self.is_running:
                self.log_signal.emit("Starting new iteration.")

                self.watchdog.reset()
                if not self.execute_click(self.points[0].x, self.points[0].y, self.points[0].name):
                    self.log_signal.emit("Failed after retries: Step 1")
                    break

                self.watchdog.reset()
                if not self.execute_click(self.points[1].x, self.points[1].y, self.points[1].name):
                    self.log_signal.emit("Failed after retries: Step 2")
                    break
                current_date = datetime.datetime.now().strftime('%Y-%m-%d')
                if not self.dry_run:
                    pyautogui.typewrite(current_date)
                self.log_signal.emit(f"Entered date: {current_date}")
                if not self.safe_sleep_with_interrupt(2):
                    break

                self.watchdog.reset()
                if not self.execute_click(self.points[2].x, self.points[2].y, self.points[2].name):
                    self.log_signal.emit("Failed after retries: Step 3")
                    break

                # FIX: Pause watchdog for Step 4 wait to avoid false timeouts
                self.log_signal.emit(f"Step 4: Waiting for {self.step4_wait_sec} seconds.")
                self.watchdog.cancel()
                if not self.safe_sleep_with_interrupt(self.step4_wait_sec):
                    break
                self.watchdog.reset()

                self.watchdog.reset()
                if not self.execute_click(self.points[3].x, self.points[3].y, self.points[3].name):
                    self.log_signal.emit("Failed after retries: Step 5")
                    break

                hrs = self.total_seconds // 3600
                mins = (self.total_seconds % 3600) // 60
                self.log_signal.emit(f"Step 6: Long wait for {hrs}h {mins}m.")
                self.watchdog.cancel()
                remaining = self.total_seconds
                while remaining > 0 and self.is_running:
                    self.update_timer_signal.emit(remaining)
                    # During long wait, look for OBS error dialog and clear it if appears
                    clear_obs_broadcast_error(True)
                    time.sleep(1)
                    remaining -= 1
                if not self.is_running:
                    self.log_signal.emit("Automation interrupted during long wait.")
                    break

                self.log_signal.emit("Long wait completed.")
                self.watchdog.reset()
                if not self.safe_sleep_with_interrupt(2):
                    break

                # Before Step 7/8 try to clear any lingering OBS error dialog
                clear_obs_broadcast_error(True)

                self.watchdog.reset()
                if not self.execute_click(self.points[4].x, self.points[4].y, self.points[4].name):
                    self.log_signal.emit("Failed after retries: Step 7")
                    break

                clear_obs_broadcast_error(True)

                self.watchdog.reset()
                if not self.execute_click(self.points[5].x, self.points[5].y, self.points[5].name):
                    self.log_signal.emit("Failed after retries: Step 8")
                    break

                self.log_signal.emit("Iteration completed. Restarting loop.")
                self.watchdog.cancel()
                if not self.safe_sleep_with_interrupt(5):
                    break
                self.watchdog.reset()

        except Exception as e:
            logger.exception("An error occurred in AutomationThread")
            self.log_signal.emit(f"An error occurred: {e}")
            self.status_signal.emit("Status: Error")
            self.error_popup_signal.emit(str(e))
        finally:
            self.is_running = False
            self.watchdog.cancel()
            self.stop_signal.emit()

class CaptureOverlay(QtWidgets.QWidget):
    captured = QtCore.pyqtSignal(int, int)
    cancelled = QtCore.pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint
            | QtCore.Qt.WindowStaysOnTopHint
            | QtCore.Qt.Tool
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self._build()
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self._tick)
        self._last_ctrl = False

    def _build(self):
        self.resize(420, 120)
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        self.move(int(screen.center().x() - self.width()/2), 40)

        self.card = QtWidgets.QFrame(self)
        self.card.setGeometry(0, 0, 420, 120)
        self.card.setStyleSheet("""
            QFrame {
                background: rgba(255,255,255,0.92);
                border-radius: 12px;
                border: 1px solid #c9c9c9;
            }
        """)

        v = QtWidgets.QVBoxLayout(self.card)
        title = QtWidgets.QLabel("Coordinate Capture")
        title.setAlignment(QtCore.Qt.AlignCenter)
        title.setFont(QtGui.QFont("Segoe UI", 12, QtGui.QFont.Bold))
        v.addWidget(title)

        self.pos_lbl = QtWidgets.QLabel("Position: (0, 0)")
        self.pos_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self.pos_lbl.setFont(QtGui.QFont("Segoe UI", 11))
        v.addWidget(self.pos_lbl)

        tip = QtWidgets.QLabel("Place the mouse, then press LEFT CTRL to capture. Press ESC to cancel.")
        tip.setAlignment(QtCore.Qt.AlignCenter)
        tip.setFont(QtGui.QFont("Segoe UI", 10))
        v.addWidget(tip)

    def start(self):
        self.show()
        self.timer.start(25)

    def _tick(self):
        pos = pyautogui.position()
        self.pos_lbl.setText(f"Position: ({pos.x}, {pos.y})")
        if key_pressed(VK_ESCAPE):
            self.timer.stop()
            self.hide()
            self.cancelled.emit()
            return
        ctrl_now = key_pressed(VK_LCONTROL)
        if ctrl_now and not self._last_ctrl:
            self.timer.stop()
            self.hide()
            self.captured.emit(pos.x, pos.y)
            return
        self._last_ctrl = ctrl_now

class RunnerTab(QtWidgets.QWidget):
    start_clicked = QtCore.pyqtSignal()
    stop_clicked  = QtCore.pyqtSignal()

    def __init__(self):
        super().__init__()
        self._build()

    def _build(self):
        layout = QtWidgets.QVBoxLayout(self)

        title = QtWidgets.QLabel(f"{APP_TITLE} - by {APP_AUTHOR}")
        title.setAlignment(QtCore.Qt.AlignCenter)
        title.setFont(QtGui.QFont("Segoe UI", 20, QtGui.QFont.Bold))
        layout.addWidget(title)

        time_row = QtWidgets.QHBoxLayout()
        self.hours_input = QtWidgets.QSpinBox()
        self.hours_input.setRange(0, 48)
        self.hours_input.setPrefix("Hours: ")
        self.minutes_input = QtWidgets.QSpinBox()
        self.minutes_input.setRange(0, 59)
        self.minutes_input.setPrefix("Minutes: ")
        time_row.addWidget(self.hours_input)
        time_row.addWidget(self.minutes_input)
        layout.addLayout(time_row)

        settings_row = QtWidgets.QHBoxLayout()
        self.step_delay = QtWidgets.QSpinBox()
        self.step_delay.setRange(0, 120)
        self.step_delay.setPrefix("Step delay (s): ")

        self.max_retries = QtWidgets.QSpinBox()
        self.max_retries.setRange(1, 10)
        self.max_retries.setPrefix("Retries: ")

        self.watchdog_sec = QtWidgets.QSpinBox()
        self.watchdog_sec.setRange(3, 600)
        self.watchdog_sec.setPrefix("Watchdog (s): ")

        self.step4_wait = QtWidgets.QSpinBox()
        self.step4_wait.setRange(0, 300)
        self.step4_wait.setPrefix("Step 4 wait (s): ")

        settings_row.addWidget(self.step_delay)
        settings_row.addWidget(self.max_retries)
        settings_row.addWidget(self.watchdog_sec)
        settings_row.addWidget(self.step4_wait)
        layout.addLayout(settings_row)

        toggles_row = QtWidgets.QHBoxLayout()
        self.dry_run = QtWidgets.QCheckBox("Dry run (no actual clicks)")
        self.always_on_top = QtWidgets.QCheckBox("Always on top")
        toggles_row.addWidget(self.dry_run)
        toggles_row.addWidget(self.always_on_top)
        layout.addLayout(toggles_row)

        self.status_label = QtWidgets.QLabel("Status: Idle")
        self.status_label.setAlignment(QtCore.Qt.AlignCenter)
        self.status_label.setFont(QtGui.QFont("Segoe UI", 12))
        layout.addWidget(self.status_label)

        self.timer_label = QtWidgets.QLabel("Timer: Not Started")
        self.timer_label.setAlignment(QtCore.Qt.AlignCenter)
        self.timer_label.setFont(QtGui.QFont("Segoe UI", 14, QtGui.QFont.Bold))
        layout.addWidget(self.timer_label)

        btn_row = QtWidgets.QHBoxLayout()
        self.start_btn = QtWidgets.QPushButton("Start")
        self.stop_btn = QtWidgets.QPushButton("Stop")
        self.stop_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_clicked.emit)
        self.stop_btn.clicked.connect(self.stop_clicked.emit)
        btn_row.addWidget(self.start_btn)
        btn_row.addWidget(self.stop_btn)
        layout.addLayout(btn_row)

        self.progress = QtWidgets.QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        hotkey = QtWidgets.QLabel("Hotkeys: Delete = Emergency Stop • Capture: LEFT CTRL • Esc = cancel capture")
        hotkey.setAlignment(QtCore.Qt.AlignCenter)
        hotkey.setFont(QtGui.QFont("Segoe UI", 9))
        layout.addWidget(hotkey)

        self.log_view = QtWidgets.QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMinimumHeight(160)
        layout.addWidget(self.log_view)

        layout.addStretch()

class CoordinatesTab(QtWidgets.QWidget):
    def __init__(self, points: list):
        super().__init__()
        self.points = points
        self.overlay = None
        self._build()

    def _build(self):
        layout = QtWidgets.QVBoxLayout(self)
        info = QtWidgets.QLabel(
            "Edit coordinates. Click Pick, then press LEFT CTRL to capture.\nESC cancels. Use Test Click to fire one click."
        )
        layout.addWidget(info)

        self.table = QtWidgets.QTableWidget(len(self.points), 4)
        self.table.setHorizontalHeaderLabels(["Step", "X", "Y", "Actions"])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        for row, p in enumerate(self.points):
            name_item = QtWidgets.QTableWidgetItem(p.name)
            self.table.setItem(row, 0, name_item)

            x_spin = QtWidgets.QSpinBox()
            x_spin.setRange(0, 9999)
            x_spin.setValue(p.x)
            x_spin.valueChanged.connect(lambda val, r=row: self._update_point(r, "x", val))
            self.table.setCellWidget(row, 1, x_spin)

            y_spin = QtWidgets.QSpinBox()
            y_spin.setRange(0, 9999)
            y_spin.setValue(p.y)
            y_spin.valueChanged.connect(lambda val, r=row: self._update_point(r, "y", val))
            self.table.setCellWidget(row, 2, y_spin)

            actions = QtWidgets.QWidget()
            h = QtWidgets.QHBoxLayout(actions)
            h.setContentsMargins(0, 0, 0, 0)
            pick_btn = QtWidgets.QPushButton("Pick")
            test_btn = QtWidgets.QPushButton("Test Click")
            pick_btn.clicked.connect(lambda _=None, r=row: self._pick_coord_ctrl(r))
            test_btn.clicked.connect(lambda _=None, r=row: self._test_click(r))
            h.addWidget(pick_btn)
            h.addWidget(test_btn)
            self.table.setCellWidget(row, 3, actions)

        self.table.resizeColumnsToContents()
        layout.addWidget(self.table)

        row2 = QtWidgets.QHBoxLayout()
        self.save_btn = QtWidgets.QPushButton("Save Coordinates")
        self.load_btn = QtWidgets.QPushButton("Reload Saved")
        self.save_btn.clicked.connect(self.save_to_settings)
        self.load_btn.clicked.connect(self.load_from_settings)
        row2.addWidget(self.save_btn)
        row2.addWidget(self.load_btn)
        layout.addLayout(row2)

        layout.addStretch()

    def _update_point(self, row: int, field: str, value: int):
        if field == "x":
            self.points[row].x = value
        else:
            self.points[row].y = value

    def _pick_coord_ctrl(self, row: int):
        parent = self.window()
        parent.showMinimized()
        self.overlay = CaptureOverlay()
        self.overlay.captured.connect(lambda x, y, r=row: self._apply_capture(r, x, y))
        self.overlay.cancelled.connect(self._cancel_capture)
        self.overlay.start()

    def _apply_capture(self, row: int, x: int, y: int):
        self.points[row].x, self.points[row].y = x, y
        self.table.cellWidget(row, 1).setValue(x)
        self.table.cellWidget(row, 2).setValue(y)
        logger.info(f"Captured {self.points[row].name}: ({x}, {y})")
        self.window().showNormal()
        self.overlay = None

    def _cancel_capture(self):
        logger.info("Coordinate capture cancelled.")
        self.window().showNormal()
        self.overlay = None

    def _test_click(self, row: int):
        x, y = self.points[row].x, self.points[row].y
        try:
            pyautogui.click(x, y)
            logger.info(f"Test clicked at ({x}, {y}) for {self.points[row].name}")
        except pyautogui.FailSafeException:
            logger.exception("PyAutoGUI Fail-safe triggered during test click.")
            QtWidgets.QMessageBox.warning(self, "Test Click", "Fail-safe triggered.")
        except Exception as e:
            logger.exception(f"Test click failed: {e}")
            QtWidgets.QMessageBox.critical(self, "Test Click Failed", str(e))

    def save_to_settings(self):
        settings = QtCore.QSettings("xTheRedShirtx", "AutoRunnerPro")
        save_points(settings, self.points)
        QtWidgets.QMessageBox.information(self, "Saved", "Coordinates saved.")

    def load_from_settings(self):
        settings = QtCore.QSettings("xTheRedShirtx", "AutoRunnerPro")
        loaded = load_points(settings)
        if len(loaded) != len(self.points):
            QtWidgets.QMessageBox.warning(self, "Mismatch", "Saved set has different size; ignoring.")
            return
        self.points[:] = loaded
        for r, p in enumerate(self.points):
            self.table.item(r, 0).setText(p.name)
            self.table.cellWidget(r, 1).setValue(p.x)
            self.table.cellWidget(r, 2).setValue(p.y)
        QtWidgets.QMessageBox.information(self, "Loaded", "Coordinates reloaded.")

class DebugTab(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self._build()

    def _build(self):
        layout = QtWidgets.QVBoxLayout(self)
        self.log_view = QtWidgets.QTextEdit()
        self.log_view.setReadOnly(True)
        layout.addWidget(self.log_view)

        row = QtWidgets.QHBoxLayout()
        open_btn = QtWidgets.QPushButton("Open Log File")
        copy_btn = QtWidgets.QPushButton("Copy Logs")
        clear_btn = QtWidgets.QPushButton("Clear View")
        row.addWidget(open_btn)
        row.addWidget(copy_btn)
        row.addWidget(clear_btn)
        layout.addLayout(row)

        open_btn.clicked.connect(self._open_log)
        copy_btn.clicked.connect(self._copy_logs)
        clear_btn.clicked.connect(lambda: self.log_view.clear())

    def _open_log(self):
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(LOG_FILE))

    def _copy_logs(self):
        self.log_view.selectAll()
        self.log_view.copy()
        QtWidgets.QMessageBox.information(self, "Copied", "Logs copied to clipboard.")

class HelpTab(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        v = QtWidgets.QVBoxLayout(self)
        txt = QtWidgets.QTextBrowser()
        txt.setOpenExternalLinks(True)
        txt.setHtml(f"""
            <h2>How it works</h2>
            <p><b>Built by {APP_AUTHOR}</b></p>
            <ul>
                <li><b>Runner</b>: set waits, retries, watchdog; Start/Stop controls.</li>
                <li><b>Coordinates</b>: Pick with <b>Left Ctrl</b>. ESC cancels.</li>
                <li><b>Debug</b>: Live logs; open/copy log file.</li>
                <li><b>Hotkeys</b>: <b>Delete</b> = emergency stop. Fail-safe: move mouse to top-left.</li>
            </ul>
        """)
        v.addWidget(txt)
        v.addStretch()

class LogEmitter(QtCore.QObject):
    sig = QtCore.pyqtSignal(str)

class QtSignalLogHandler(logging.Handler):
    def __init__(self, emitter: LogEmitter):
        super().__init__()
        self.emitter = emitter
    def emit(self, record):
        try:
            msg = self.format(record)
            self.emitter.sig.emit(msg)
        except Exception:
            pass

class AutomationApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_TITLE} - by {APP_AUTHOR}")
        self.setGeometry(100, 100, 900, 700)
        self.setWindowIcon(self.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon))

        self.settings = QtCore.QSettings("xTheRedShirtx", "AutoRunnerPro")
        self.points = load_points(self.settings)

        self._build_ui()
        self._install_gui_logger()

        self.runner_tab.hours_input.setValue(int(self.settings.value("longwait/hours", 11)))
        self.runner_tab.minutes_input.setValue(int(self.settings.value("longwait/mins", 30)))
        self.runner_tab.step_delay.setValue(int(self.settings.value("settings/step_delay", 10)))
        self.runner_tab.max_retries.setValue(int(self.settings.value("settings/retries", 3)))
        self.runner_tab.watchdog_sec.setValue(int(self.settings.value("settings/watchdog", 15)))
        self.runner_tab.step4_wait.setValue(int(self.settings.value("settings/step4wait", 10)))
        self.runner_tab.dry_run.setChecked(bool(int(self.settings.value("settings/dry_run", 0))))
        self.runner_tab.always_on_top.setChecked(bool(int(self.settings.value("settings/ontop", 1))))
        self._apply_always_on_top(self.runner_tab.always_on_top.isChecked())

        self.thread = None

        sys.excepthook = self._handle_exception

        self._apply_styles()

        self._delete_pressed_last = False
        self.hotkey_timer = QtCore.QTimer(self)
        self.hotkey_timer.timeout.connect(self._poll_hotkeys)
        self.hotkey_timer.start(50)

    def _build_ui(self):
        self.tabs = QtWidgets.QTabWidget()
        self.runner_tab = RunnerTab()
        self.coords_tab = CoordinatesTab(self.points)
        self.debug_tab = DebugTab()
        self.help_tab = HelpTab()

        self.tabs.addTab(self.runner_tab, "Runner")
        self.tabs.addTab(self.coords_tab, "Coordinates")
        self.tabs.addTab(self.debug_tab, "Debug")
        self.tabs.addTab(self.help_tab, "Help")

        self.setCentralWidget(self.tabs)

        self.runner_tab.start_clicked.connect(self.start_automation)
        self.runner_tab.stop_clicked.connect(self.stop_automation)
        self.runner_tab.always_on_top.stateChanged.connect(
            lambda _: self._apply_always_on_top(self.runner_tab.always_on_top.isChecked())
        )

        menubar = self.menuBar()
        file_menu = menubar.addMenu("&File")
        act_open_log = file_menu.addAction("Open Log")
        act_open_log.triggered.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(LOG_FILE)))
        file_menu.addSeparator()
        act_exit = file_menu.addAction("Exit")
        act_exit.triggered.connect(self.close)

        view_menu = menubar.addMenu("&View")
        self.act_ontop = view_menu.addAction("Always on Top")
        self.act_ontop.setCheckable(True)
        self.act_ontop.setChecked(self.runner_tab.always_on_top.isChecked())
        self.act_ontop.triggered.connect(lambda checked: self.runner_tab.always_on_top.setChecked(checked))

        self.statusBar().showMessage("Ready")

    def _apply_always_on_top(self, enabled: bool):
        flags = self.windowFlags()
        if enabled:
            flags |= QtCore.Qt.WindowStaysOnTopHint
        else:
            flags &= ~QtCore.Qt.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.show()
        self.settings.setValue("settings/ontop", int(enabled))
        self.act_ontop.setChecked(enabled)

    def _apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background: #121212; color: #eaeaea; }
            QLabel, QCheckBox, QMenuBar, QMenu, QStatusBar { color: #eaeaea; }
            QMenuBar { background: #1a1a1a; }
            QMenu { background: #1a1a1a; border: 1px solid #333; }
            QStatusBar { background: #1a1a1a; }
            QTabWidget::pane { border: 1px solid #333; top: -1px; background: #121212; }
            QTabBar { font-weight: 600; }
            QTabBar::tab {
                background: #ffffff; color: #000000; padding: 8px 16px; margin-right: 2px;
                border: 1px solid #c9c9c9; border-top-left-radius: 6px; border-top-right-radius: 6px;
            }
            QTabBar::tab:selected { background: #ffffff; color: #000000; border-color: #8aa8ff; border-bottom-color: #ffffff; margin-bottom: -1px; }
            QTabBar::tab:!selected { background: #f2f2f2; color: #000000; }
            QTabBar::tab:hover { background: #fafafa; }
            QTabBar::tab:disabled { color: #888888; background: #f5f5f5; }
            QTextEdit, QSpinBox, QTableWidget, QPushButton, QProgressBar, QCheckBox {
                background: #1a1a1a; color: #eaeaea; border: 1px solid #333;
            }
            QPushButton:hover { border: 1px solid #555; }
            QProgressBar::chunk { background: #3d85c6; }
            QHeaderView::section { background: #1a1a1a; color: #eaeaea; border: 1px solid #333; }
        """)

    def _install_gui_logger(self):
        self._log_emitter = LogEmitter()
        self._log_emitter.sig.connect(self._append_logs, QtCore.Qt.QueuedConnection)
        qt_handler = QtSignalLogHandler(self._log_emitter)
        qt_handler.setLevel(logging.INFO)
        qt_handler.setFormatter(logging.Formatter("%(levelname)s - %(message)s"))
        logger.addHandler(qt_handler)
        logger.info(f"===== Session started. Author: {APP_AUTHOR} =====")

    def _append_logs(self, message: str):
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        text = f"[{stamp}] {message}"
        self.runner_tab.log_view.append(text)
        self.debug_tab.log_view.append(text)

    def _make_thread(self):
        hours = self.runner_tab.hours_input.value()
        mins = self.runner_tab.minutes_input.value()
        total_sec = hours * 3600 + mins * 60
        t = AutomationThread(
            points=[ClickPoint(p.name, p.x, p.y) for p in self.points],
            long_wait_seconds=total_sec,
            step_delay=self.runner_tab.step_delay.value(),
            max_retries=self.runner_tab.max_retries.value(),
            watchdog_seconds=self.runner_tab.watchdog_sec.value(),
            step4_wait_sec=self.runner_tab.step4_wait.value(),
            dry_run=self.runner_tab.dry_run.isChecked(),
        )
        t.log_signal.connect(self._log)
        t.status_signal.connect(self._update_status)
        t.update_timer_signal.connect(self._update_timer)
        t.stop_signal.connect(self._on_thread_stopped)
        t.error_popup_signal.connect(self._error_popup)
        return t

    def start_automation(self):
        self.settings.setValue("longwait/hours", self.runner_tab.hours_input.value())
        self.settings.setValue("longwait/mins", self.runner_tab.minutes_input.value())
        self.settings.setValue("settings/step_delay", self.runner_tab.step_delay.value())
        self.settings.setValue("settings/retries", self.runner_tab.max_retries.value())
        self.settings.setValue("settings/watchdog", self.runner_tab.watchdog_sec.value())
        self.settings.setValue("settings/step4wait", self.runner_tab.step4_wait.value())
        self.settings.setValue("settings/dry_run", int(self.runner_tab.dry_run.isChecked()))
        save_points(self.settings, self.points)

        if self.thread and self.thread.isRunning():
            QtWidgets.QMessageBox.warning(self, "Already running", "Automation is already running.")
            return

        self.thread = self._make_thread()
        self.runner_tab.start_btn.setEnabled(False)
        self.runner_tab.stop_btn.setEnabled(True)
        self.thread.start()
        self.statusBar().showMessage("Automation running")
        logger.info("Automation started")

    def stop_automation(self):
        if self.thread:
            try:
                self.thread.stop()
            except Exception:
                pass
            try:
                # Give the thread a moment to exit cooperatively
                self.thread.wait(2000)
            except Exception:
                pass
            self.thread = None
        self._on_thread_stopped()

    def _on_thread_stopped(self):
        self.runner_tab.start_btn.setEnabled(True)
        self.runner_tab.stop_btn.setEnabled(False)
        self.runner_tab.progress.setValue(0)
        self.runner_tab.timer_label.setText("Timer: Not Started")
        self.statusBar().showMessage("Stopped")
        logger.info("Automation stopped")

    def _log(self, message: str):
        logger.info(message)

    def _update_status(self, status: str):
        self.runner_tab.status_label.setText(status)

    def _update_timer(self, remaining_seconds: int):
        total = self.runner_tab.hours_input.value() * 3600 + self.runner_tab.minutes_input.value() * 60
        h, r = divmod(remaining_seconds, 3600)
        m, s = divmod(r, 60)
        self.runner_tab.timer_label.setText(f"Time Remaining: {h:02}:{m:02}:{s:02}")
        if total > 0:
            elapsed = total - remaining_seconds
            pct = max(0, min(100, int((elapsed / total) * 100)))
            self.runner_tab.progress.setValue(pct)
        else:
            self.runner_tab.progress.setValue(0)

    def _poll_hotkeys(self):
        if not is_windows():
            return
        del_now = key_pressed(VK_DELETE)
        if del_now and not self._delete_pressed_last:
            if self.thread and self.thread.isRunning():
                logger.info("Emergency stop: Delete key pressed.")
                try:
                    self.stop_automation()
                except Exception:
                    logger.exception("Emergency stop failed")
        self._delete_pressed_last = del_now

    def _handle_exception(self, etype, value, tb):
        msg = "".join(traceback.format_exception(etype, value, tb))
        logger.exception("Unhandled exception:\n" + msg)
        QtWidgets.QMessageBox.critical(self, "Unhandled Error", str(value))

    def _error_popup(self, message: str):
        QtWidgets.QMessageBox.critical(self, "Error", message)

    def closeEvent(self, event: QtGui.QCloseEvent):
        try:
            if self.thread and self.thread.isRunning():
                self.stop_automation()
        finally:
            super().closeEvent(event)

def main():
    try:
        app = QtWidgets.QApplication(sys.argv)
        app.setApplicationName("AutoRunnerPro")
        app.setOrganizationName("xTheRedShirtx")
        window = AutomationApp()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        logger.exception("Fatal error in main")
        try:
            QtWidgets.QMessageBox.critical(None, "Fatal Error", str(e))
        except Exception:
            pass
        sys.exit(1)

if __name__ == "__main__":
    main()
