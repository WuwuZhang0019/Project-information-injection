import ctypes
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from src import cad_windows


class CadWindowsTests(unittest.TestCase):
    def test_enumerates_both_autocad_process_windows(self):
        titles = {
            101: "Autodesk AutoCAD 2020 - first.dwg",
            102: "AutoCAD Properties",
            202: "second.dwg - project",
            303: "AutoCAD help in a browser",
        }
        pids = {101: 11, 102: 11, 202: 22, 303: 33}

        def enumerate_windows(callback, state):
            for hwnd in titles:
                callback(hwnd, state)

        with patch.object(cad_windows.win32gui, "EnumWindows",
                          side_effect=enumerate_windows), \
             patch.object(cad_windows.win32gui, "IsWindowVisible", return_value=True), \
             patch.object(cad_windows.win32gui, "GetWindowText",
                          side_effect=lambda hwnd: titles[hwnd]), \
             patch.object(cad_windows.win32gui, "GetWindowRect",
                          side_effect=lambda hwnd: (0, 0, 100, 100) if hwnd == 102
                          else (0, 0, 1000, 800)), \
             patch.object(cad_windows.win32process, "GetWindowThreadProcessId",
                          side_effect=lambda hwnd: (0, pids[hwnd])), \
             patch.object(cad_windows, "_process_image",
                          side_effect=lambda pid: "C:/AutoCAD/acad.exe" if pid != 33
                          else "C:/Browser/browser.exe"):
            windows = cad_windows.enumerate_autocad_windows()

        self.assertEqual([item[1] for item in windows], [11, 22])
        self.assertEqual(windows[0][0], 101)

    def test_window_command_requires_verified_foreground_target(self):
        with patch.object(cad_windows.win32gui, "IsWindow", return_value=True), \
             patch.object(cad_windows.win32gui, "IsIconic", return_value=False), \
             patch.object(cad_windows.win32gui, "SetForegroundWindow"), \
             patch.object(cad_windows.win32gui, "GetForegroundWindow", return_value=999), \
             patch.object(cad_windows.win32process, "GetWindowThreadProcessId",
                          return_value=(0, 22)), \
             patch.object(cad_windows, "_is_autocad_process", return_value=True), \
             patch.object(cad_windows, "_send_unicode_command") as send, \
             patch.object(cad_windows.time, "sleep"):
            with self.assertRaisesRegex(RuntimeError, "未发送命令"):
                cad_windows.send_cad_command(202, 22, "(load \"file.lsp\") ")
            send.assert_not_called()

    def test_window_command_reaches_selected_window(self):
        with patch.object(cad_windows.win32gui, "IsWindow", return_value=True), \
             patch.object(cad_windows.win32gui, "IsIconic", return_value=False), \
             patch.object(cad_windows.win32gui, "SetForegroundWindow"), \
             patch.object(cad_windows.win32gui, "GetForegroundWindow", return_value=202), \
             patch.object(cad_windows.win32process, "GetWindowThreadProcessId",
                          return_value=(0, 22)), \
             patch.object(cad_windows, "_is_autocad_process", return_value=True), \
             patch.object(cad_windows, "_send_unicode_command") as send, \
             patch.object(cad_windows.time, "sleep"):
            cad_windows.send_cad_command(202, 22, "(load \"file.lsp\") ")
            send.assert_called_once_with("(load \"file.lsp\")")

    def test_win64_keyboard_input_matches_windows_structure_size(self):
        if ctypes.sizeof(ctypes.c_void_p) == 8:
            self.assertEqual(ctypes.sizeof(cad_windows._Input), 40)

    def test_unicode_command_includes_full_text_and_enter(self):
        send = Mock(side_effect=lambda count, events, size: count)
        with patch.object(cad_windows.ctypes, "WinDLL",
                          return_value=SimpleNamespace(SendInput=send)):
            cad_windows._send_unicode_command("A中")

        count, events, size = send.call_args.args
        self.assertEqual(count, 6)
        self.assertEqual(size, ctypes.sizeof(cad_windows._Input))
        self.assertEqual(events[0].value.ki.wScan, ord("A"))
        self.assertEqual(events[2].value.ki.wScan, ord("中"))
        self.assertEqual(events[4].value.ki.wVk, cad_windows.VK_RETURN)


if __name__ == "__main__":
    unittest.main()
