import html
import os
import sys
import time
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QTextEdit,
    QLineEdit,
    QButtonGroup,
    QFileDialog,
)
from PySide6.QtCore import (
    Qt,
    QTimer,
    QThread,
    Signal,
    Slot,
    QStandardPaths,
)
from PySide6.QtGui import QFont


# --------------------------------------------------
# PROJECT PATH
# --------------------------------------------------

ROOT = Path(__file__).resolve().parent.parent

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.chdir(ROOT)


# --------------------------------------------------
# MACOS NATIVE WINDOW SUPPORT
# --------------------------------------------------

IS_MACOS = sys.platform == "darwin"

if IS_MACOS:
    try:
        import objc
        from AppKit import (
            NSApplication,
            NSWindowCollectionBehaviorCanJoinAllSpaces,
            NSWindowCollectionBehaviorFullScreenAuxiliary,
            NSFloatingWindowLevel,
        )
        MAC_NATIVE_AVAILABLE = True
    except ImportError:
        MAC_NATIVE_AVAILABLE = False
else:
    MAC_NATIVE_AVAILABLE = False


# --------------------------------------------------
# COLORS AND UI TEXT
# --------------------------------------------------

CHAT_COLOR = "#2563EB"
EDIT_COLOR = "#EA580C"
ERROR_COLOR = "#B91C1C"

MODE_HINTS = {
    "chat": (
        "Chat mode: Ask questions and analyze your workbook. "
        "No changes will be made to Excel."
    ),
    "edit": (
        "Edit mode: Give instructions to modify your Excel workbook."
    ),
}

MODE_PLACEHOLDERS = {
    "chat": "Ask about your Excel data...",
    "edit": "Tell me what to change in Excel...",
}


# --------------------------------------------------
# BACKGROUND AGENT WORKER
# --------------------------------------------------

class AgentWorker(QThread):
    """Runs backend requests without blocking the UI."""

    reply_ready = Signal(str)
    failed = Signal(str)

    def __init__(self, text: str, mode: str):
        super().__init__()
        self.text = text
        self.mode = mode

    def run(self):
        try:
            from backend.orchestrator import process_user_request

            response = process_user_request(
                self.text,
                self.mode,
            )

            if response is None:
                response = "Request completed."

            self.reply_ready.emit(str(response))

        except Exception as exc:
            self.failed.emit(
                f"{type(exc).__name__}: {exc}"
            )


# --------------------------------------------------
# MAIN DESKTOP APPLICATION CONTROLLER
# --------------------------------------------------

class ExcelAgentUI:

    def __init__(self, app: QApplication):
        self.app = app
        self.mode = "chat"
        self.worker = None
        self._last_state = {}
        self.attached_path = None

        self.create_agent_icon()
        self.create_chat_panel()

        if MAC_NATIVE_AVAILABLE:
            self.state_timer = QTimer()
            self.state_timer.timeout.connect(
                self.log_native_state
            )
            self.state_timer.start(2000)

    # --------------------------------------------------
    # NATIVE WINDOW CONFIGURATION
    # --------------------------------------------------

    def configure_mac_window(self, widget):
        if not MAC_NATIVE_AVAILABLE:
            return

        try:
            self.app.processEvents()

            native_view = objc.objc_object(
                c_void_p=int(widget.winId())
            )

            native_window = native_view.window()

            if native_window is None:
                return

            behavior = (
                NSWindowCollectionBehaviorCanJoinAllSpaces
                | NSWindowCollectionBehaviorFullScreenAuxiliary
            )

            native_window.setCollectionBehavior_(behavior)
            native_window.setLevel_(NSFloatingWindowLevel)

        except Exception as error:
            print(
                "Native window configuration error:",
                repr(error),
            )

    def describe_native_state(self, widget):
        if not MAC_NATIVE_AVAILABLE:
            return "Native state unavailable"

        try:
            native_window = objc.objc_object(
                c_void_p=int(widget.winId())
            ).window()

            ns_app = NSApplication.sharedApplication()

            parts = [
                f"class={native_window.className()}"
            ]

            for key in (
                "level",
                "collectionBehavior",
                "isVisible",
                "isOnActiveSpace",
                "occlusionState",
                "hidesOnDeactivate",
            ):
                try:
                    value = getattr(native_window, key)()
                    parts.append(f"{key}={value}")
                except Exception:
                    parts.append(f"{key}=n/a")

            parts.append(
                f"app_active={ns_app.isActive()}"
            )

            return " ".join(parts)

        except Exception as error:
            return f"STATE ERROR: {error!r}"

    def log_native_state(self):
        for label, widget in (
            ("ICON", self.agent_icon),
            ("PANEL", self.chat_panel),
        ):
            if label == "PANEL" and not widget.isVisible():
                continue

            state = self.describe_native_state(widget)

            if self._last_state.get(label) != state:
                self._last_state[label] = state
                print(
                    time.strftime("%H:%M:%S"),
                    label,
                    state,
                )

    # --------------------------------------------------
    # FLOATING AGENT ICON
    # --------------------------------------------------

    def create_agent_icon(self):
        self.agent_icon = QPushButton("✦")
        self.agent_icon.setFixedSize(58, 58)

        self.agent_icon.setWindowFlags(
            Qt.WindowType.Window
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.WindowDoesNotAcceptFocus
        )

        self.agent_icon.setAttribute(
            Qt.WidgetAttribute.WA_ShowWithoutActivating
        )

        self.agent_icon.setStyleSheet("""
            QPushButton {
                background-color: #2563EB;
                color: white;
                border-radius: 29px;
                font-size: 30px;
                border: none;
            }
            QPushButton:hover {
                background-color: #1D4ED8;
            }
        """)

        self.agent_icon.clicked.connect(self.toggle_chat)

        screen = self.app.primaryScreen().availableGeometry()
        self.agent_icon.move(
            screen.right() - 85,
            screen.center().y(),
        )
        self.agent_icon.show()
        self.configure_mac_window(self.agent_icon)

    # --------------------------------------------------
    # CHAT PANEL
    # --------------------------------------------------

    def create_chat_panel(self):
        self.chat_panel = QWidget()
        self.chat_panel.setWindowTitle("Excel Agent")
        self.chat_panel.setFixedSize(390, 560)

        self.chat_panel.setWindowFlags(
            Qt.WindowType.Window
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )

        self.chat_panel.setStyleSheet("""
            QWidget {
                background-color: #FFFFFF;
                color: #202124;
                font-family: Arial;
            }
            QLineEdit {
                border: 1px solid #D1D5DB;
                border-radius: 12px;
                padding: 12px;
                font-size: 13px;
            }
            QPushButton {
                background-color: #2563EB;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 9px;
            }
            QPushButton#modeChat,
            QPushButton#modeEdit {
                background-color: #E5E7EB;
                color: #374151;
                font-weight: bold;
            }
            QPushButton#modeChat:checked {
                background-color: #2563EB;
                color: white;
            }
            QPushButton#modeEdit:checked {
                background-color: #EA580C;
                color: white;
            }
            QPushButton#stopButton {
                background-color: #B91C1C;
            }
            QTextEdit {
                border: none;
                background-color: #F8FAFC;
                border-radius: 10px;
                padding: 10px;
            }
        """)

        layout = QVBoxLayout(self.chat_panel)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        # HEADER
        header = QHBoxLayout()
        title = QLabel("✦  Excel Agent")
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))

        close_button = QPushButton("✕")
        close_button.setFixedWidth(35)
        close_button.clicked.connect(self.hide_chat)

        header.addWidget(title)
        header.addStretch()
        header.addWidget(close_button)
        layout.addLayout(header)

        # CHAT AND EDIT MODES
        mode_layout = QHBoxLayout()
        self.chat_button = QPushButton("Chat")
        self.chat_button.setObjectName("modeChat")
        self.chat_button.setCheckable(True)
        self.chat_button.setChecked(True)

        self.edit_button = QPushButton("Edit")
        self.edit_button.setObjectName("modeEdit")
        self.edit_button.setCheckable(True)

        self.mode_group = QButtonGroup(self.chat_panel)
        self.mode_group.setExclusive(True)
        self.mode_group.addButton(self.chat_button)
        self.mode_group.addButton(self.edit_button)

        self.chat_button.clicked.connect(lambda: self.set_mode("chat"))
        self.edit_button.clicked.connect(lambda: self.set_mode("edit"))

        mode_layout.addWidget(self.chat_button)
        mode_layout.addWidget(self.edit_button)
        layout.addLayout(mode_layout)

        # MODE DESCRIPTION
        self.mode_hint = QLabel()
        self.mode_hint.setWordWrap(True)
        layout.addWidget(self.mode_hint)

        # WELCOME MESSAGE
        welcome = QLabel(
            "Hello! I'm your Excel Agent.\n\n"
            "What would you like to do with this workbook?"
        )
        welcome.setWordWrap(True)
        welcome.setFont(QFont("Arial", 13))
        layout.addWidget(welcome)

        # CHAT HISTORY
        self.chat_history = QTextEdit()
        self.chat_history.setReadOnly(True)
        layout.addWidget(self.chat_history, 1)

        # STATUS ROW
        status_layout = QHBoxLayout()
        self.status_label = QLabel("Working... Please wait.")
        self.status_label.setStyleSheet("color: #6B7280; font-size: 12px;")
        self.status_label.hide()

        self.stop_button = QPushButton("Stop")
        self.stop_button.setObjectName("stopButton")
        self.stop_button.hide()
        self.stop_button.clicked.connect(self.stop_request)

        status_layout.addWidget(self.status_label, 1)
        status_layout.addWidget(self.stop_button)
        layout.addLayout(status_layout)

        # ATTACHMENT CHIP (shown only after a file is picked)
        self.attached_path = None
        self.chip = QPushButton()
        self.chip.setObjectName("chip")
        self.chip.setStyleSheet("""
            QPushButton#chip {
                background-color: #E5E7EB;
                color: #374151;
                border-radius: 10px;
                padding: 4px 10px;
                font-size: 12px;
                text-align: left;
            }
        """)
        self.chip.setToolTip("Click to remove")
        self.chip.clicked.connect(self.clear_attachment)
        self.chip.hide()
        layout.addWidget(self.chip)

        # MESSAGE INPUT
        input_layout = QHBoxLayout()

        self.plus_button = QPushButton("+")
        self.plus_button.setFixedSize(38, 38)
        self.plus_button.setToolTip("Attach an image")
        self.plus_button.setStyleSheet("""
            QPushButton {
                background-color: #E5E7EB;
                color: #374151;
                border-radius: 19px;
                font-size: 20px;
                font-weight: bold;
                padding: 0px;
            }
            QPushButton:hover {
                background-color: #D1D5DB;
            }
        """)
        self.plus_button.clicked.connect(self.pick_file)

        self.message_input = QLineEdit()
        self.message_input.returnPressed.connect(self.send_message)

        self.send_button = QPushButton("➤")
        self.send_button.clicked.connect(self.send_message)

        input_layout.addWidget(self.plus_button)
        input_layout.addWidget(self.message_input, 1)
        input_layout.addWidget(self.send_button)
        layout.addLayout(input_layout)

        self.apply_mode_style()

    # --------------------------------------------------
    # MODE HANDLING
    # --------------------------------------------------

    def mode_color(self, mode=None):
        selected_mode = mode or self.mode
        if selected_mode == "edit":
            return EDIT_COLOR
        return CHAT_COLOR

    def apply_mode_style(self):
        color = self.mode_color()
        self.send_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                border-radius: 8px;
                padding: 9px;
            }}
        """)
        self.message_input.setStyleSheet(f"""
            QLineEdit {{
                border: 2px solid {color};
                border-radius: 12px;
                padding: 11px;
                font-size: 13px;
            }}
        """)
        self.message_input.setPlaceholderText(MODE_PLACEHOLDERS[self.mode])
        self.mode_hint.setText(MODE_HINTS[self.mode])
        self.mode_hint.setStyleSheet(f"color: {color}; font-size: 12px;")

    def set_mode(self, mode: str):
        if self.worker is not None:
            return
        if mode == self.mode:
            return
        self.mode = mode
        self.apply_mode_style()
        self.chat_history.append(
            f"""
            <p style="color:{self.mode_color()}; margin:0 0 10px 0">
            <i>— {mode.capitalize()} mode —</i>
            </p>
            """
        )

    # --------------------------------------------------
    # FILE ATTACHMENT (UI only, not sent to backend yet)
    # --------------------------------------------------

    def pick_file(self):
        if self.worker is not None:
            return
        start = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.DesktopLocation
        )
        path, _ = QFileDialog.getOpenFileName(
            self.chat_panel,
            "Select an image",
            start,
            "Images (*.png *.jpg *.jpeg *.webp);;All files (*)",
        )
        if path:
            self.attached_path = path
            self.chip.setText(f"{Path(path).name}   ✕")
            self.chip.show()

    def clear_attachment(self):
        self.attached_path = None
        self.chip.hide()

    # --------------------------------------------------
    # SHOW / HIDE CHAT PANEL
    # --------------------------------------------------

    def toggle_chat(self):
        if self.chat_panel.isVisible():
            self.hide_chat()
        else:
            self.show_chat()

    def show_chat(self):
        screen = self.app.primaryScreen().availableGeometry()
        x = self.agent_icon.x() - self.chat_panel.width() - 12
        y = (
            self.agent_icon.y()
            + self.agent_icon.height() // 2
            - self.chat_panel.height() // 2
        )
        y = max(
            screen.top() + 10,
            min(y, screen.bottom() - self.chat_panel.height() - 10),
        )

        self.chat_panel.move(x, y)
        self.chat_panel.show()
        self.configure_mac_window(self.chat_panel)
        self.chat_panel.raise_()
        self.agent_icon.raise_()
        self.chat_panel.activateWindow()

        if self.worker is None:
            self.message_input.setFocus()

    def hide_chat(self):
        self.chat_panel.hide()
        self.agent_icon.show()
        self.agent_icon.raise_()

    # --------------------------------------------------
    # CHAT MESSAGES
    # --------------------------------------------------

    def add_message(self, who: str, message: str, color: str = "#202124"):
        safe_message = html.escape(str(message)).replace("\n", "<br>")
        safe_name = html.escape(who)
        self.chat_history.append(
            f"""
            <p style="margin:0 0 10px 0">
                <b style="color:{color}">
                    {safe_name}:
                </b>
                {safe_message}
            </p>
            """
        )

    # --------------------------------------------------
    # BUSY STATE
    # --------------------------------------------------

    def set_busy(self, busy: bool):
        self.message_input.setEnabled(not busy)
        self.send_button.setEnabled(not busy)
        self.plus_button.setEnabled(not busy)
        self.chat_button.setEnabled(not busy)
        self.edit_button.setEnabled(not busy)
        self.stop_button.setVisible(busy)
        self.status_label.setVisible(busy)

        if not busy:
            self.stop_button.setEnabled(True)
            self.message_input.setFocus()

    # --------------------------------------------------
    # BACKEND REQUEST
    # --------------------------------------------------

    def start_request(self, message: str, mode: str):
        if self.worker is not None:
            return

        self.add_message("You", message)
        self.set_busy(True)

        worker = AgentWorker(message, mode)
        worker.reply_ready.connect(self.on_reply)
        worker.failed.connect(self.on_failed)
        worker.finished.connect(self.on_worker_finished)
        worker.finished.connect(worker.deleteLater)

        self.worker = worker
        worker.start()

    def send_message(self):
        if self.worker is not None:
            return

        message = self.message_input.text().strip()
        if not message:
            return

        self.message_input.clear()
        self.start_request(message, self.mode)

    # --------------------------------------------------
    # STOP REQUEST
    # --------------------------------------------------

    def stop_request(self):
        if self.worker is None:
            return

        self.status_label.setText("Stopping request...")
        self.stop_button.setEnabled(False)

        try:
            from backend.orchestrator import stop_current_request
            stop_current_request()
        except Exception as exc:
            self.add_message(
                "Excel Agent",
                f"Stop request error: {type(exc).__name__}: {exc}",
                ERROR_COLOR,
            )

    # --------------------------------------------------
    # BACKEND RESPONSE
    # --------------------------------------------------

    @Slot(str)
    def on_reply(self, reply: str):
        color = self.mode_color()
        if reply.startswith("Error:"):
            color = ERROR_COLOR
        self.add_message("Excel Agent", reply, color)

    @Slot(str)
    def on_failed(self, error: str):
        self.add_message("Excel Agent", f"Error: {error}", ERROR_COLOR)

    @Slot()
    def on_worker_finished(self):
        self.worker = None
        self.set_busy(False)


# --------------------------------------------------
# APPLICATION ENTRY POINT
# --------------------------------------------------

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    # Keep controller alive as a strong reference variable!
    controller = ExcelAgentUI(app)

    sys.exit(app.exec())