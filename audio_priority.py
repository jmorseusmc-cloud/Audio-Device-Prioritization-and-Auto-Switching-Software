import sys
import json
import time
import ctypes

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QListWidget, QPushButton, QMessageBox, QHBoxLayout
)
from PyQt5.QtGui import QClipboard

from pycaw.pycaw import AudioUtilities

import comtypes
from comtypes import GUID, HRESULT, IUnknown
from comtypes import client as cc
from comtypes import COMError
from ctypes import c_wchar_p, c_int, c_void_p

# ========= Config =========
CONFIG_FILE = "audio_priority.json"
ROLES = (0, 1, 2)  # eConsole, eMultimedia, eCommunications
POLL_MS = 5000

# ========= COM init helper =========
def _coinitialize():
    try:
        ctypes.oledll.ole32.CoInitialize(None)
    except Exception:
        pass  # S_FALSE / mode change is fine

# ========= PolicyConfig COM (try multiple IIDs; gracefully fail) =========
CLSID_PolicyConfigClient = GUID("{870AF99C-171D-4F9E-AF0D-E63DF40C2BC9}")
IID_IPolicyConfigVista   = GUID("{568B9108-44BF-40B4-9006-86AFE5B5A620}")
IID_IPolicyConfig        = GUID("{F8679F50-850A-41CF-9C72-430F290290C8}")
IID_IPolicyConfig10      = GUID("{CA286FC3-91FD-42C3-8E9B-CAAFA66242E3}")

class IPolicyConfigVista(IUnknown):
    _iid_ = IID_IPolicyConfigVista
    _methods_ = [
        comtypes.COMMETHOD([], HRESULT, "GetMixFormat",        ([], c_wchar_p), ([], c_void_p)),
        comtypes.COMMETHOD([], HRESULT, "GetDeviceFormat",     ([], c_wchar_p), ([], c_int), ([], c_void_p)),
        comtypes.COMMETHOD([], HRESULT, "SetDeviceFormat",     ([], c_wchar_p), ([], c_void_p), ([], c_void_p)),
        comtypes.COMMETHOD([], HRESULT, "GetProcessingPeriod", ([], c_wchar_p), ([], c_int), ([], c_void_p), ([], c_void_p)),
        comtypes.COMMETHOD([], HRESULT, "SetProcessingPeriod", ([], c_wchar_p), ([], c_void_p)),
        comtypes.COMMETHOD([], HRESULT, "GetShareMode",        ([], c_wchar_p), ([], c_void_p)),
        comtypes.COMMETHOD([], HRESULT, "SetShareMode",        ([], c_wchar_p), ([], c_void_p)),
        comtypes.COMMETHOD([], HRESULT, "GetPropertyValue",    ([], c_wchar_p), ([], c_void_p), ([], c_void_p)),
        comtypes.COMMETHOD([], HRESULT, "SetPropertyValue",    ([], c_wchar_p), ([], c_void_p), ([], c_void_p)),
        comtypes.COMMETHOD([], HRESULT, "SetDefaultEndpoint",  ([], c_wchar_p), ([], c_int)),
        comtypes.COMMETHOD([], HRESULT, "SetEndpointVisibility", ([], c_wchar_p), ([], c_int)),
    ]

class IPolicyConfig(IPolicyConfigVista):
    _iid_ = IID_IPolicyConfig

class IPolicyConfig10(IPolicyConfigVista):
    _iid_ = IID_IPolicyConfig10

_POLICY_IFACES = (IPolicyConfigVista, IPolicyConfig, IPolicyConfig10)

import subprocess
import os

SVV_PATH = os.path.join(os.path.dirname(__file__), "SoundVolumeView.exe")

def set_default_endpoint(device_name: str):
    """Switch default audio device using SoundVolumeView.exe"""
    if not os.path.exists(SVV_PATH):
        raise FileNotFoundError("SoundVolumeView.exe not found next to the script.")
    subprocess.run([SVV_PATH, "/SetDefault", device_name, "all"], check=True)
    print(f"Switched audio to: {device_name} (via SVV)")


# ========= App =========
class AudioPriorityApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Audio Output Priority")
        self.setGeometry(200, 200, 640, 480)

        layout = QVBoxLayout()

        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)

        # Row of action buttons
        row = QHBoxLayout()
        btn_refresh = QPushButton("Refresh")
        btn_refresh.clicked.connect(self.refresh_devices)
        row.addWidget(btn_refresh)

        btn_up = QPushButton("Move Up")
        btn_up.clicked.connect(self.move_up)
        row.addWidget(btn_up)

        btn_down = QPushButton("Move Down")
        btn_down.clicked.connect(self.move_down)
        row.addWidget(btn_down)

        btn_save = QPushButton("Save")
        btn_save.clicked.connect(self.save_priority)
        row.addWidget(btn_save)

        btn_load = QPushButton("Load")
        btn_load.clicked.connect(self.load_priority)
        row.addWidget(btn_load)

        layout.addLayout(row)

        # Manual fallback helpers
        row2 = QHBoxLayout()
        btn_copy_id = QPushButton("Copy Selected Device ID")
        btn_copy_id.clicked.connect(self.copy_selected_id)
        row2.addWidget(btn_copy_id)

        btn_open_settings = QPushButton("Open Sound Settings")
        btn_open_settings.clicked.connect(self.open_sound_settings)
        row2.addWidget(btn_open_settings)

        layout.addLayout(row2)

        self.setLayout(layout)

        _coinitialize()
        self.refresh_devices()

        self._last_set_device_id = None
        self._last_set_time = 0.0

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.enforce_priority_once)
        self.timer.start(POLL_MS)

    # ---------- Device enumeration ----------
    def refresh_devices(self):
        self.list_widget.clear()
        any_added = False
        try:
            for dev in AudioUtilities.GetAllDevices():
                if getattr(dev, "DataFlow", None) == 0 and getattr(dev, "State", None) == 1:
                    name = getattr(dev, "FriendlyName", None)
                    did = getattr(dev, "id", None) or (dev.GetId() if hasattr(dev, "GetId") else None)
                    if name and did:
                        self.list_widget.addItem(f"{name}|{did}")
                        any_added = True
        except Exception as e:
            print(f"Refresh (GetAllDevices) error: {e}")

        if not any_added:
            try:
                enumr = AudioUtilities.GetDeviceEnumerator()
                coll = enumr.EnumAudioEndpoints(0, 1)  # eRender, ACTIVE
                for i in range(coll.GetCount()):
                    imm = coll.Item(i)
                    did = imm.GetId()
                    name = None
                    try:
                        if hasattr(AudioUtilities, "CreateDevice"):
                            mm = AudioUtilities.CreateDevice(imm)
                            name = getattr(mm, "FriendlyName", None)
                    except Exception:
                        pass
                    if not name:
                        name = "Playback Device"
                    self.list_widget.addItem(f"{name}|{did}")
                    any_added = True
            except Exception as e:
                print(f"Refresh (EnumAudioEndpoints) error: {e}")

        if not any_added:
            self.list_widget.addItem("No playback devices found")

    def is_device_active(self, device_id: str) -> bool:
        try:
            enumr = AudioUtilities.GetDeviceEnumerator()
            coll = enumr.EnumAudioEndpoints(0, 1)  # eRender, ACTIVE
            for i in range(coll.GetCount()):
                if coll.Item(i).GetId() == device_id:
                    return True
        except Exception:
            pass
        return False

    # ---------- UI controls ----------
    def move_up(self):
        row = self.list_widget.currentRow()
        if row > 0:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row - 1, item)
            self.list_widget.setCurrentRow(row - 1)

    def move_down(self):
        row = self.list_widget.currentRow()
        if row < self.list_widget.count() - 1:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row + 1, item)
            self.list_widget.setCurrentRow(row + 1)

    def save_priority(self):
        items = [self.list_widget.item(i).text() for i in range(self.list_widget.count())]
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)
        QMessageBox.information(self, "Saved", "Priority list saved successfully.")

    def load_priority(self):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                items = json.load(f)
            self.list_widget.clear()
            self.list_widget.addItems(items)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not load config: {e}")

    def copy_selected_id(self):
        item = self.list_widget.currentItem()
        if not item:
            QMessageBox.information(self, "Copy Device ID", "Select a device first.")
            return
        try:
            _, device_id = item.text().split("|", 1)
            QApplication.clipboard().setText(device_id, QClipboard.Clipboard)
            QMessageBox.information(self, "Copied", "Device ID copied to clipboard.")
        except Exception:
            QMessageBox.warning(self, "Copy Failed", "Could not parse the selected list item.")

    def open_sound_settings(self):
        # Opens Windows Sound settings for quick manual switch
        ctypes.windll.shell32.ShellExecuteW(None, "open", "ms-settings:sound", None, None, 1)

    # ---------- Enforce priority ----------
    def enforce_priority_once(self):
        try:
            if self.list_widget.count() == 0:
                return

            items = [self.list_widget.item(i).text() for i in range(self.list_widget.count())]
            for entry in items:
                if "|" not in entry:
                    continue
                _, device_id = entry.split("|", 1)

                if self.is_device_active(device_id):
                    now = time.time()
                    if device_id != self._last_set_device_id or (now - self._last_set_time) > 30:
                        # Attempt COM switch; if not supported on this box, we log and stop retrying
                        ok = try_set_default_endpoint(device_id)
                        if ok:
                            self._last_set_device_id = device_id
                            self._last_set_time = now
                        else:
                            print("Auto-switch not supported on this Windows build (PolicyConfig interfaces unavailable).")
                            # Stop the timer to avoid spamming
                            self.timer.stop()
                        # Either way, break after the first active device in your priority list
                    break
        except Exception as e:
            print(f"Monitor error: {e}")

# ========= Main =========
if __name__ == "__main__":
    _coinitialize()
    app = QApplication(sys.argv)
    window = AudioPriorityApp()
    window.show()
    sys.exit(app.exec_())
