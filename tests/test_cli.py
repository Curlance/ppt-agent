import argparse

from pptd import cli


def test_cli_preserves_explicit_text_whitespace():
    assert cli._kv_pairs(["text=  内容  ", "shape=2"]) == {"text": "  内容  ", "shape": "2"}
    assert cli._kv_pairs(["text="])["text"] == ""


def test_doctor_uses_daemon_instead_of_second_com(monkeypatch, capsys):
    import pptd.com
    import pptd.mcp_server

    seen = []
    info = {"pid": 1, "port": 8791, "token": "test"}
    monkeypatch.setattr(cli, "runtime_info", lambda cfg: info)
    monkeypatch.setattr(cli, "ensure", lambda cfg, timeout: info)
    monkeypatch.setattr(cli, "observer_url", lambda cfg: "http://localhost/")
    monkeypatch.setattr(cli, "interpreter_report", lambda: {"executable": "python", "version": "3.13", "in_venv": True, "pywin32": True})
    def forbid(**kwargs):
        raise AssertionError("自检不能另开 COM 会话")
    monkeypatch.setattr(pptd.com, "ComSession", forbid)
    class Caller:
        def __init__(self, cfg, record):
            assert record == info
        def call(self, name, args):
            seen.append(name)
            return {"ok": True, "result": {"com": {"alive": True, "visible": True, "version": "16.0"}, "presentations": []}}
    monkeypatch.setattr(pptd.mcp_server, "ForwardingCaller", Caller)
    assert cli.cmd_doctor(argparse.Namespace()) == 0
    assert seen == ["ppt_status"]
    assert "经守护进程" in capsys.readouterr().out
