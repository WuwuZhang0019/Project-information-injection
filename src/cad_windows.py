"""Find AutoCAD windows and send a command when COM cannot reach an instance."""

import ctypes
from ctypes import wintypes
import os
import time

import win32gui
import win32process


PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
INPUT_KEYBOARD = 1
VK_RETURN = 0x0D
SW_RESTORE = 9


class _MouseInput(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_void_p)]


class _KeyboardInput(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_void_p)]


class _HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD)]


class _InputUnion(ctypes.Union):
    _fields_ = [("mi", _MouseInput), ("ki", _KeyboardInput),
                ("hi", _HardwareInput)]


class _Input(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("value", _InputUnion)]


def _process_image(pid):
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD)]
    kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return None
    try:
        size = wintypes.DWORD(32768)
        buffer = ctypes.create_unicode_buffer(size.value)
        if kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
            return buffer.value
        return None
    finally:
        kernel32.CloseHandle(handle)


def _is_autocad_process(pid):
    image = _process_image(pid)
    return bool(image and os.path.basename(image).casefold() == "acad.exe")


def enumerate_autocad_windows():
    """Return (main HWND, PID, title) for visible acad.exe instances."""
    candidates = {}
    checked_pids = {}

    def visit(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return True
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            return True
        if pid not in checked_pids:
            checked_pids[pid] = _is_autocad_process(pid)
        if checked_pids[pid]:
            try:
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                area = max(0, right - left) * max(0, bottom - top)
            except Exception:
                area = 0
            score = ("autocad" in title.casefold(), area)
            if pid not in candidates or score > candidates[pid][0]:
                candidates[pid] = (score, hwnd, title)
        return True

    win32gui.EnumWindows(visit, None)
    return [(hwnd, pid, title) for pid, (_, hwnd, title) in candidates.items()]


def _keyboard_event(vk=0, scan=0, flags=0):
    return _Input(INPUT_KEYBOARD,
                  _InputUnion(ki=_KeyboardInput(vk, scan, flags, 0, None)))


def _send_unicode_command(command):
    inputs = []
    encoded = command.encode("utf-16-le")
    for offset in range(0, len(encoded), 2):
        scan = int.from_bytes(encoded[offset:offset + 2], "little")
        inputs.append(_keyboard_event(scan=scan, flags=KEYEVENTF_UNICODE))
        inputs.append(_keyboard_event(scan=scan,
                                      flags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP))
    inputs.extend([_keyboard_event(vk=VK_RETURN),
                   _keyboard_event(vk=VK_RETURN, flags=KEYEVENTF_KEYUP)])
    events = (_Input * len(inputs))(*inputs)
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(_Input), ctypes.c_int]
    user32.SendInput.restype = wintypes.UINT
    sent = user32.SendInput(len(events), events, ctypes.sizeof(_Input))
    if sent != len(events):
        raise RuntimeError("Windows 未能向 CAD 窗口发送完整命令；请检查运行权限")


def send_cad_command(hwnd, pid, command):
    """Activate the exact CAD window and type one command into its prompt."""
    if not win32gui.IsWindow(hwnd):
        raise RuntimeError("目标 CAD 窗口已关闭，请刷新列表")
    _, current_pid = win32process.GetWindowThreadProcessId(hwnd)
    if current_pid != pid or not _is_autocad_process(pid):
        raise RuntimeError("目标窗口已变化，请刷新 CAD 实例列表")
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, SW_RESTORE)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception as exc:
        raise RuntimeError("无法激活目标 CAD 窗口，未发送命令") from exc
    time.sleep(0.1)
    if win32gui.GetForegroundWindow() != hwnd:
        raise RuntimeError("无法激活目标 CAD 窗口，未发送命令")
    _send_unicode_command(command.rstrip())


class WindowCadApplication:
    Name = "AutoCAD"

    def __init__(self, hwnd, pid, title):
        self.HWND = hwnd
        self.pid = pid
        self.Caption = title


class WindowCadDocument:
    Name = "当前活动图纸"

    def __init__(self, application):
        self.Application = application

    def SendCommand(self, command):
        send_cad_command(self.Application.HWND, self.Application.pid, command)
