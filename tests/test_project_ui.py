import tkinter as tk
import unittest
from tkinter import ttk
from unittest.mock import patch

from src.main import CadInfoInjector


class ProjectUiTests(unittest.TestCase):
    def setUp(self):
        projects = {
            "示例": {"test": "value"},
            "扬州": {"原项": {"value": "1.0"}},
            "上海": {"另一项": {}},
        }

        def load_data(app):
            app.projects = projects

        with (patch.object(CadInfoInjector, "load_data", load_data),
              patch.object(CadInfoInjector, "refresh_cad_list")):
            self.app = CadInfoInjector()
        self.app.withdraw()
        self.save_data = patch.object(self.app, "save_data")
        self.save_data.start()

    def tearDown(self):
        self.save_data.stop()
        self.app.destroy()

    def select(self, item_id):
        self.app.proj_tree.selection_set(item_id)
        self.app.on_project_select(None)

    def test_add_child_keeps_other_groups_open_and_shows_new_child(self):
        tree = self.app.proj_tree
        tree.item("group_扬州", open=True)
        tree.item("group_上海", open=True)
        self.select("group_扬州")

        with patch.object(self.app, "prompt_dialog", return_value="新项"):
            self.app.add_project()

        self.assertTrue(tree.item("group_扬州", "open"))
        self.assertTrue(tree.item("group_上海", "open"))
        self.assertEqual(tree.selection(), ("proj_扬州_新项",))
        self.assertEqual(self.app.current_project, ("扬州", "新项"))

    def test_copy_to_another_group_expands_destination(self):
        tree = self.app.proj_tree
        tree.item("group_扬州", open=True)
        self.select("proj_扬州_原项")

        with (patch.object(self.app, "prompt_combobox", return_value="上海"),
              patch.object(self.app, "prompt_dialog", return_value="副本")):
            self.app.copy_project()

        self.assertTrue(tree.item("group_扬州", "open"))
        self.assertTrue(tree.item("group_上海", "open"))
        self.assertEqual(tree.selection(), ("proj_上海_副本",))
        self.assertEqual(self.app.projects["上海"]["副本"], {"value": "1.0"})

    def test_rename_group_keeps_it_expanded(self):
        tree = self.app.proj_tree
        tree.item("group_扬州", open=True)
        self.select("group_扬州")

        with patch.object(self.app, "prompt_dialog", return_value="扬州新"):
            self.app.rename_item()

        self.assertTrue(tree.item("group_扬州新", "open"))
        self.assertEqual(tree.selection(), ("group_扬州新",))

    def test_refresh_keeps_a_manually_collapsed_selected_group_closed(self):
        tree = self.app.proj_tree
        tree.item("group_扬州", open=True)
        self.select("proj_扬州_原项")
        tree.item("group_扬州", open=False)

        self.app.refresh_project_list()

        self.assertFalse(tree.item("group_扬州", "open"))
        self.assertEqual(tree.selection(), ("proj_扬州_原项",))

    def test_input_dialog_keeps_chinese_punctuation(self):
        def enter_value():
            dialog = next(child for child in self.app.winfo_children()
                          if isinstance(child, tk.Toplevel))
            entry = next(child for child in dialog.winfo_children()
                         if isinstance(child, ttk.Entry))
            entry.delete(0, tk.END)
            entry.insert(0, "中文。．，1.0")
            ok = next(child for child in dialog.winfo_children()
                      if isinstance(child, tk.Frame))
            next(child for child in ok.winfo_children()
                 if isinstance(child, tk.Button) and child.cget("text") == "确定").invoke()

        self.app.after(0, enter_value)
        self.assertEqual(
            self.app.prompt_dialog("修改变量值", "输入变量值:", "旧值"),
            "中文。．，1.0")


if __name__ == "__main__":
    unittest.main()
