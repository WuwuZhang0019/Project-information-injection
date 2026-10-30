import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from src.main import CadInfoInjector


class LspGenerationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory(
            dir=os.environ.get("CAD_INFO_TEST_TMP"))
        self.addCleanup(self.temp_dir.cleanup)
        self.lsp_path = Path(self.temp_dir.name) / "current_project.lsp"
        self.doc = Mock()
        self.app = Mock()

    def inject(self, state):
        with (patch("src.main.LSP_FILE", str(self.lsp_path)),
              patch("src.main.messagebox.showerror") as showerror,
              patch("src.main.messagebox.showwarning") as showwarning):
            CadInfoInjector.inject_to_cad(state)
        return showerror, showwarning

    def make_state(self, projects, current_project, auto_regen):
        return SimpleNamespace(
            projects=projects,
            current_project=current_project,
            auto_regen_var=SimpleNamespace(get=lambda: auto_regen),
            cad_target_var=SimpleNamespace(get=lambda: "地下室.dwg"),
            cad_docs_cache=[("地下室.dwg", "地下室.dwg", self.app, self.doc)],
        )

    def test_middle_dot_in_value_and_project_names_is_preserved(self):
        state = self.make_state(
            {"介休·组": {"地下室·分项": {
                "xmmc": "富德隆中心、街里·1989",
                "zxmc": "地下室",
            }}},
            ("介休·组", "地下室·分项"),
            auto_regen=True,
        )

        showerror, showwarning = self.inject(state)

        showerror.assert_not_called()
        showwarning.assert_not_called()
        content = self.lsp_path.read_bytes().decode("gbk")
        self.assertIn('(setq xmmc "富德隆中心、街里·1989")', content)
        self.assertIn('(setq zxmc "地下室")', content)
        self.assertIn("[介休·组 / 地下室·分项]", content)
        self.assertIn('(command "REGEN")', content)
        cad_path = os.path.abspath(self.lsp_path).replace("\\", "/")
        self.doc.SendCommand.assert_called_once_with(f'(load "{cad_path}") ')

    def test_middle_dot_is_preserved_without_auto_regen(self):
        state = self.make_state(
            {"项目·根": {"xmmc": "中心·1989"}},
            (None, "项目·根"),
            auto_regen=False,
        )

        showerror, showwarning = self.inject(state)

        showerror.assert_not_called()
        showwarning.assert_not_called()
        content = self.lsp_path.read_bytes().decode("gbk")
        self.assertIn('(setq xmmc "中心·1989")', content)
        self.assertIn("[项目·根]", content)
        self.assertNotIn('(command "REGEN")', content)
        self.doc.SendCommand.assert_called_once()

    def test_unsupported_character_does_not_overwrite_previous_script(self):
        original = b"previous script"
        self.lsp_path.write_bytes(original)
        state = self.make_state(
            {"示例": {"xmmc": "项目🌟"}},
            (None, "示例"),
            auto_regen=True,
        )

        showerror, _ = self.inject(state)

        showerror.assert_called_once()
        self.assertIn("U+1F31F", showerror.call_args[0][1])
        self.assertEqual(self.lsp_path.read_bytes(), original)
        self.doc.SendCommand.assert_not_called()


if __name__ == "__main__":
    unittest.main()
