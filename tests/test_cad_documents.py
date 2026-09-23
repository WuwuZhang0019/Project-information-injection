import unittest
from types import SimpleNamespace
from unittest.mock import patch

from src.main import CadInfoInjector


class FakeDocuments:
    def __init__(self, docs):
        self.docs = docs
        self.Count = len(docs)

    def Item(self, index):
        return self.docs[index]


class FakeApplication:
    Name = "AutoCAD Application"

    def __init__(self, hwnd, docs):
        self.HWND = hwnd
        self.Documents = FakeDocuments(docs)
        self.Application = self


class FakeDocument:
    def __init__(self, name, full_name, application=None):
        self.Name = name
        self.FullName = full_name
        self.Application = application


class FakeEnumerator:
    def __init__(self, objects):
        self.objects = iter(objects)

    def Next(self, count):
        item = next(self.objects, None)
        return (item,) if item is not None else ()


class FakeRot:
    def __init__(self, objects):
        self.objects = objects

    def EnumRunning(self):
        return FakeEnumerator(self.objects)

    def GetObject(self, moniker):
        return FakeUnknown(moniker)


class FakeUnknown:
    def __init__(self, obj):
        self.obj = obj

    def QueryInterface(self, iid):
        return self.obj


class FakeVar:
    def __init__(self):
        self.value = ""

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeCombo:
    def __init__(self, variable):
        self.variable = variable
        self.values = []
        self.index = -1

    def __setitem__(self, key, value):
        if key == 'values':
            self.values = value

    def current(self, index=None):
        if index is None:
            return self.index
        self.index = index
        self.variable.set(self.values[index])

    def set(self, value):
        self.variable.set(value)


class FakeLabel:
    def config(self, **kwargs):
        self.options = kwargs


class CadDocumentDiscoveryTests(unittest.TestCase):
    def discover(self, entries, active_object=None):
        rot = FakeRot(entries)
        with patch("src.main.pythoncom.GetRunningObjectTable", return_value=rot), \
             patch("src.main.win32com.client.Dispatch", side_effect=lambda obj: obj), \
             patch("src.main.win32com.client.GetActiveObject", return_value=active_object):
            return CadInfoInjector.get_all_cad_documents(None)

    def test_document_moniker_reveals_later_instance_and_all_its_drawings(self):
        first_doc = FakeDocument("first.dwg", r"C:\first.dwg")
        second_doc = FakeDocument("second.dwg", r"D:\second.dwg")
        third_doc = FakeDocument("third.dwg", r"D:\third.dwg")
        first = FakeApplication(1001, [first_doc])
        second = FakeApplication(1002, [second_doc, third_doc])
        second_doc.Application = second
        third_doc.Application = second

        docs = self.discover([first, second_doc, third_doc, second_doc])

        self.assertEqual([item[1] for item in docs],
                         ["first.dwg", "second.dwg", "third.dwg"])
        self.assertIs(docs[1][2], second)
        self.assertIs(docs[2][3], third_doc)

    def test_same_drawing_path_in_two_instances_remains_selectable(self):
        path = r"C:\shared.dwg"
        first_doc = FakeDocument("shared.dwg", path)
        second_doc = FakeDocument("shared.dwg", path)
        first = FakeApplication(1001, [first_doc])
        second = FakeApplication(1002, [second_doc])
        second_doc.Application = second

        docs = self.discover([first, second_doc])

        self.assertEqual(len(docs), 2)
        self.assertNotEqual(docs[0][0], docs[1][0])
        self.assertIn("1001", docs[0][0])
        self.assertIn("1002", docs[1][0])

    def test_active_object_fallback_still_works(self):
        doc = FakeDocument("fallback.dwg", r"C:\fallback.dwg")
        app = FakeApplication(1001, [doc])

        docs = self.discover([], active_object=app)

        self.assertEqual(len(docs), 1)
        self.assertIs(docs[0][2], app)

    def test_selecting_instance_limits_injection_targets_to_its_drawings(self):
        first_doc = FakeDocument("first.dwg", r"C:\first.dwg")
        second_doc = FakeDocument("second.dwg", r"D:\second.dwg")
        first = FakeApplication(1001, [first_doc])
        second = FakeApplication(1002, [second_doc])
        docs = [("first.dwg", "first.dwg", first, first_doc),
                ("second.dwg", "second.dwg", second, second_doc)]
        state = SimpleNamespace(
            all_cad_docs_cache=[], cad_docs_cache=[], cad_instances=[],
            selected_cad_hwnd=None, cad_target_var=FakeVar(),
            selected_cad_title=FakeVar(), cad_status_label=FakeLabel(),
            get_all_cad_documents=lambda: docs,
            get_cad_instances=lambda items: [("CAD 1", 1001), ("CAD 2", 1002)],
        )
        state.cad_combo = FakeCombo(state.cad_target_var)
        state.cad_combobox = FakeCombo(state.selected_cad_title)
        state.show_cad_documents = lambda: CadInfoInjector.show_cad_documents(state)

        CadInfoInjector.refresh_cad_list(state)
        self.assertEqual([item[1] for item in state.cad_docs_cache], ["first.dwg"])

        state.cad_combobox.current(1)
        CadInfoInjector.on_cad_instance_select(state)
        self.assertEqual([item[1] for item in state.cad_docs_cache], ["second.dwg"])
        self.assertEqual(state.cad_target_var.get(), "second.dwg")

        CadInfoInjector.refresh_cad_list(state)
        self.assertEqual(state.selected_cad_hwnd, 1002)
        self.assertEqual(state.cad_target_var.get(), "second.dwg")


if __name__ == "__main__":
    unittest.main()
