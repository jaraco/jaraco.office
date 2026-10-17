import sys
import types

from jaraco.office import convert


def test_converter_init(monkeypatch):
    word = types.SimpleNamespace(Quit=lambda: None)
    client = types.SimpleNamespace(Dispatch=lambda name: word)
    monkeypatch.setitem(sys.modules, 'win32com', types.ModuleType('win32com'))
    monkeypatch.setitem(sys.modules, 'win32com.client', client)
    monkeypatch.setitem(sys.modules, 'pythoncom', types.SimpleNamespace())
    assert convert.Converter().word is word
