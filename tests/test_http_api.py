"""HTTP 面：观察台、鉴权、通用调用、每个工具一个 OpenAPI 端点。

用 TestClient 直连 FastAPI 应用，配有假 COM 的引擎，无需守护进程。
"""

from __future__ import annotations

import json

import pytest
from fastapi.testclient import TestClient

from pptd.http_api import create_app

TOKEN = "test-token"


@pytest.fixture()
def client(engine) -> TestClient:
    return TestClient(create_app(engine))


@pytest.fixture()
def auth() -> dict[str, str]:
    return {"X-PPT-Token": TOKEN}


def test_health_needs_no_token(client: TestClient) -> None:
    resp = client.get("/health")
    assert resp.status_code == 200
    body = resp.json()
    assert body["ok"] is True
    assert body["name"] == "ppt-agent"


def test_state_requires_token(client: TestClient) -> None:
    assert client.get("/state").status_code == 401
    assert client.get("/state", headers={"X-PPT-Token": "wrong"}).status_code == 401
    resp = client.get("/state", headers={"X-PPT-Token": TOKEN})
    assert resp.status_code == 200
    assert resp.json()["com"]["alive"] is True


def test_token_via_query_and_bearer(client: TestClient) -> None:
    assert client.get(f"/state?token={TOKEN}").status_code == 200
    assert client.get("/state", headers={"Authorization": f"Bearer {TOKEN}"}).status_code == 200


def test_tools_manifest(client: TestClient, auth: dict[str, str]) -> None:
    resp = client.get("/tools", headers=auth)
    assert resp.status_code == 200
    body = resp.json()
    assert body["count"] >= 4
    names = {t["name"] for t in body["tools"]}
    assert "ppt_open" in names
    # 中文摘要必须原样带出来（观察台与模型都读它）
    open_tool = next(t for t in body["tools"] if t["name"] == "ppt_open")
    assert "打开" in open_tool["summary"]
    assert open_tool["inputSchema"]["required"] == ["path"]


def test_call_endpoint(client: TestClient, auth: dict[str, str]) -> None:
    resp = client.post("/call", json={"name": "ppt_status", "args": {}}, headers=auth)
    assert resp.status_code == 200
    assert resp.json()["ok"] is True


def test_call_unknown_tool_is_400(client: TestClient, auth: dict[str, str]) -> None:
    resp = client.post("/call", json={"name": "nope", "args": {}}, headers=auth)
    assert resp.status_code == 400
    assert resp.json()["error"]["code"] == "tool/unknown"


def test_per_tool_endpoint(client: TestClient, auth: dict[str, str]) -> None:
    """每个工具一个 REST 端点——ChatGPT Actions 直接吃 OpenAPI 里这个。"""
    resp = client.post("/tools/ppt_status", json={}, headers=auth)
    assert resp.status_code == 200
    assert resp.json()["ok"] is True


def test_per_tool_endpoint_reports_tool_error(client: TestClient, auth: dict[str, str]) -> None:
    resp = client.post("/tools/ppt_open", json={"path": "C:\\missing.pptx"}, headers=auth)
    assert resp.status_code == 400
    assert resp.json()["error"]["code"] == "path/missing"


def test_per_tool_endpoint_validates_body(client: TestClient, auth: dict[str, str]) -> None:
    resp = client.post("/tools/ppt_open", json={}, headers=auth)
    assert resp.status_code == 422  # 缺必填 path，由 pydantic 拦下


def test_per_tool_endpoint_requires_token(client: TestClient) -> None:
    assert client.post("/tools/ppt_status", json={}).status_code == 401


def test_observer_page_injects_token(client: TestClient) -> None:
    resp = client.get(f"/?token={TOKEN}")
    assert resp.status_code == 200
    assert "ppt-agent 观察台" in resp.text
    assert TOKEN in resp.text
    assert "EventSource" in resp.text


def test_observer_requires_token(client: TestClient) -> None:
    assert client.get("/").status_code == 401


def test_openapi_describes_every_tool(client: TestClient) -> None:
    """给 ChatGPT 侧用的接口契约：每个工具都要有 operationId 与入参 schema。"""
    spec = client.get("/openapi.json").json()
    paths = spec["paths"]
    for name in ("ppt_status", "ppt_open", "ppt_activate", "ppt_close"):
        assert f"/tools/{name}" in paths, f"{name} 没进 OpenAPI"
        operation = paths[f"/tools/{name}"]["post"]
        assert operation["operationId"] == name
        assert operation["summary"]
        assert "requestBody" in operation
    schema = paths["/tools/ppt_open"]["post"]["requestBody"]["content"]["application/json"]["schema"]
    assert "$ref" in schema or "properties" in schema


def test_events_route_is_registered(client: TestClient) -> None:
    spec = client.get("/openapi.json").json()
    assert "/events" in spec["paths"]


def test_shutdown_without_server_object_is_safe(client: TestClient, auth: dict[str, str]) -> None:
    """未通过 daemon.serve() 启动时（比如测试里），急停不能炸。"""
    resp = client.post("/shutdown", headers=auth)
    assert resp.status_code == 200
    assert resp.json()["stopping"] is True


# --------------------------------------------------------------------------
# 截图取图：观察台与任务窗格靠它显示画面
# --------------------------------------------------------------------------


def _make_png(engine, name: str = "demo-p001.png") -> str:
    """在截图目录里放一张真的小 PNG。"""
    from PIL import Image

    engine.settings.shot_dir.mkdir(parents=True, exist_ok=True)
    path = engine.settings.shot_dir / name
    Image.new("RGB", (32, 18), (10, 20, 40)).save(path)
    return name


def test_shot_file_requires_token(client: TestClient) -> None:
    assert client.get("/shots/whatever.png").status_code == 401


def test_shot_file_serves_image(client: TestClient, auth: dict[str, str], engine) -> None:
    name = _make_png(engine)
    resp = client.get(f"/shots/{name}", headers=auth)
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("image/png")
    assert resp.content.startswith(b"\x89PNG"), "取回来的不是 PNG 内容"


def test_shot_file_rejects_traversal(client: TestClient, auth: dict[str, str]) -> None:
    """路径穿越必须挡住——这个端点会把文件内容吐出去。"""
    assert client.get("/shots/..%2F..%2Fruntime.json", headers=auth).status_code == 404
    assert client.get("/shots/....//runtime.json", headers=auth).status_code == 404


def test_shot_file_missing_is_404(client: TestClient, auth: dict[str, str]) -> None:
    assert client.get("/shots/nope.png", headers=auth).status_code == 404


# --------------------------------------------------------------------------
# 实时取景：观察台当实时监视器用
# --------------------------------------------------------------------------


def test_live_requires_token(client: TestClient) -> None:
    assert client.get("/live").status_code == 401


def test_live_without_open_deck_is_409(client: TestClient, auth: dict[str, str]) -> None:
    """没有打开的演示时要给出明确的状态码，而不是 500。"""
    resp = client.get("/live", headers=auth)
    assert resp.status_code == 409


def test_pace_endpoint_updates_state(client: TestClient, auth: dict[str, str], engine) -> None:
    resp = client.post("/pace", json={"pace_ms": 1500}, headers=auth)
    assert resp.status_code == 200
    assert resp.json()["pace_ms"] == 1500
    assert engine.state["pace_ms"] == 1500
    assert client.get("/state", headers=auth).json()["pace_ms"] == 1500


def test_pace_endpoint_validates_range(client: TestClient, auth: dict[str, str]) -> None:
    assert client.post("/pace", json={"pace_ms": -1}, headers=auth).status_code == 422
    assert client.post("/pace", json={"pace_ms": 99999}, headers=auth).status_code == 422


def test_pace_requires_token(client: TestClient) -> None:
    assert client.post("/pace", json={"pace_ms": 100}).status_code == 401


# --------------------------------------------------------------------------
# 面板数据源：VBA / PowerShell 这类不方便读 SSE 的客户端靠它
# --------------------------------------------------------------------------


def test_feed_requires_token(client: TestClient) -> None:
    assert client.get("/feed").status_code == 401


def test_feed_returns_state_and_events(client: TestClient, auth: dict[str, str], engine) -> None:
    engine.invoke("ppt_status", {})  # 造一条事件
    resp = client.get("/feed", headers=auth)
    assert resp.status_code == 200
    body = resp.json()
    assert body["ok"] is True
    assert body["state"]["com"]["alive"] is True
    assert "pace_ms" in body["state"]
    assert body["count"] >= 2  # op.start + op.end
    types = [e["type"] for e in body["events"]]
    assert "op.start" in types and "op.end" in types
    # 事件要按时间正序，面板直接 append 即可
    assert [e["seq"] for e in body["events"]] == sorted(e["seq"] for e in body["events"])


def test_feed_limit_is_clamped(client: TestClient, auth: dict[str, str]) -> None:
    assert len(client.get("/feed?limit=1", headers=auth).json()["events"]) <= 1
    assert client.get("/feed?limit=99999", headers=auth).status_code == 200


def test_feed_includes_human_summary(client: TestClient, auth: dict[str, str], engine) -> None:
    """面板给用户看的是中文人话，不是 API 调用堆。"""
    engine.invoke("ppt_status", {})
    events = client.get("/feed", headers=auth).json()["events"]
    ended = [e for e in events if e["type"] == "op.end"]
    assert ended and ended[-1]["data"]["human"]


# --------------------------------------------------------------------------
# 纯文本面板：VBA 里解析 JSON 要手写解析器，所以另给一份 Split 就能用的格式
# --------------------------------------------------------------------------


def test_feed_text_requires_token(client: TestClient) -> None:
    assert client.get("/feed.txt").status_code == 401


def _parse_feed_text(text: str) -> tuple[dict[str, str], list[list[str]]]:
    head: dict[str, str] = {}
    rows: list[list[str]] = []
    in_rows = False
    for line in text.splitlines():
        if line == "---":
            in_rows = True
            continue
        if not in_rows:
            if "=" in line:
                key, _, value = line.partition("=")
                head[key] = value
            continue
        if line.strip():
            rows.append([part.strip() for part in line.split("|")])
    return head, rows


def test_feed_text_shape(client: TestClient, auth: dict[str, str], engine) -> None:
    engine.invoke("ppt_status", {})
    resp = client.get("/feed.txt", headers=auth)
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("text/plain")
    assert resp.text.startswith("PPT-AGENT-FEED v1")

    head, rows = _parse_feed_text(resp.text)
    assert head["alive"] == "true"
    assert head["version"] == "16.0"
    assert head["visible"] == "true"
    assert head["pace_ms"].isdigit()
    assert head["tools"].isdigit()
    assert int(head["count"]) == len(rows)

    # 每行三字段：时间 | 类型 | 中文人话 —— VBA 一句 Split 就能用
    assert all(len(row) == 3 for row in rows), rows
    types = {row[1] for row in rows}
    assert "op.start" in types and "op.end" in types
    ended = [row for row in rows if row[1] == "op.end"]
    assert ended[-1][2].startswith("OK ")
    assert "PowerPoint" in ended[-1][2]


def test_feed_text_marks_failures(client: TestClient, auth: dict[str, str], engine) -> None:
    engine.invoke("ppt_open", {"path": "C:\\definitely\\missing.pptx"})
    _, rows = _parse_feed_text(client.get("/feed.txt", headers=auth).text)
    ended = [row for row in rows if row[1] == "op.end"]
    assert ended[-1][2].startswith("ERR "), ended[-1]


def test_feed_text_has_no_paths_in_deck_field(client: TestClient, auth: dict[str, str], engine) -> None:
    """面板上只显示文件名——一整条路径会把窄面板撑爆。"""
    _, rows = _parse_feed_text(client.get("/feed.txt", headers=auth).text)
    assert rows is not None  # 结构可解析即可
    head, _ = _parse_feed_text(client.get("/feed.txt", headers=auth).text)
    assert "\\" not in head.get("deck", "")
    assert "/" not in head.get("deck", "")


# --------------------------------------------------------------------------
# 状态条：宿主界面（例如 DSH 输入框下方）那一行窄条的数据源
# --------------------------------------------------------------------------


@pytest.fixture()
def local_client(engine) -> TestClient:
    """Host=localhost 的客户端——模拟宿主界面/本机浏览器直接访问。"""
    return TestClient(create_app(engine), base_url="http://localhost")


def test_strip_serves_one_line_to_localhost(local_client: TestClient) -> None:
    resp = local_client.get("/strip.txt")
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("text/plain")
    assert resp.headers["cache-control"] == "no-store"
    assert resp.text.strip().count("\n") == 0, "状态条就该是一行"
    assert "PowerPoint" in resp.text


def test_strip_mentions_last_operation(local_client: TestClient, engine) -> None:
    engine.invoke("ppt_status", {})
    body = local_client.get("/strip.txt").text
    assert "PowerPoint" in body
    assert "正常" in body or "打开着" in body, f"应带上最近一步的中文摘要：{body}"


def test_strip_reports_disconnected_state(settings) -> None:
    """COM 不在时要给出可执行的提示，而不是让窄条空着或崩掉。"""
    import pptd.ops  # noqa: F401
    from pptd.com import ComInfo
    from pptd.engine import Engine
    from pptd.events import EventBus

    class DeadSession:
        def info(self):
            return ComInfo(alive=False, last_error="测试：COM 未启动")

        def run(self, *a, **k):
            raise AssertionError("不该调用 COM")

        def stop(self):
            pass

    dead = Engine.create(settings, bus=EventBus(), token="t", session=DeadSession())
    client = TestClient(create_app(dead), base_url="http://localhost")
    body = client.get("/strip.txt").text
    assert "未连接" in body
    assert "pptctl serve" in body


def test_strip_refuses_remote_host(client: TestClient) -> None:
    """隧道之后 Host 变成公网域名，就必须带令牌。"""
    assert client.get("/strip.txt").status_code == 401
    assert client.get("/strip.txt", headers={"X-PPT-Token": "test-token"}).status_code == 200


def test_strip_has_no_secrets(local_client: TestClient) -> None:
    """状态条不该泄露令牌——它是不带鉴权就能拿到的。"""
    body = local_client.get("/strip.txt").text
    assert "test-token" not in body
    for word in ("token", "Token", "secret", "密钥"):
        assert word not in body


def test_observer_page_served_to_localhost_without_token(local_client: TestClient) -> None:
    """用户直接敲地址、或从宿主界面点进来时带不了令牌——本机放行。"""
    resp = local_client.get("/")
    assert resp.status_code == 200
    assert "观察台" in resp.text
    assert "test-token" in resp.text, "页面里必须注入令牌，否则它调不动 API"


# --------------------------------------------------------------------------
# 状态条的跨源读取（宿主界面在别的端口上）
#
# 真实故障：插件在 DSH 里加载成功之后，状态条却永远显示"未连接"。
# 原因是 DSH 界面跑在 19387，从这里 fetch 8791 的 /strip.txt 属于跨源请求，
# 响应缺少 Access-Control-Allow-Origin 会被浏览器整个拦下（实测 TypeError: Failed to fetch）。
# --------------------------------------------------------------------------


def test_strip_allows_local_origin(local_client: TestClient) -> None:
    """宿主界面（127.0.0.1:19387）必须能读到。"""
    resp = local_client.get("/strip.txt", headers={"Origin": "http://127.0.0.1:19387"})
    assert resp.status_code == 200
    assert resp.headers.get("access-control-allow-origin") == "http://127.0.0.1:19387"
    assert resp.headers.get("vary") == "Origin", "不带 Vary 会让缓存把 CORS 头串台"


def test_strip_allows_any_localhost_port(local_client: TestClient) -> None:
    """本机任意端口都算本机——开发期的 Vite 端口（5173）也要能用。"""
    resp = local_client.get("/strip.txt", headers={"Origin": "http://localhost:5173"})
    assert resp.headers.get("access-control-allow-origin") == "http://localhost:5173"


def test_strip_refuses_foreign_origin(local_client: TestClient) -> None:
    """**这条是安全边界**：绝不能放行任意站点。

    否则你访问的任何一个网页都能读到你的演示名与当前页码。
    """
    for origin in ("https://evil.example.com", "http://192.168.1.9:8080", "null", ""):
        resp = local_client.get("/strip.txt", headers={"Origin": origin} if origin else {})
        assert resp.status_code == 200
        assert "access-control-allow-origin" not in resp.headers, f"{origin} 不该被放行"


def test_strip_preflight_is_handled(local_client: TestClient) -> None:
    """宿主将来若给 fetch 加自定义头，浏览器会先发 OPTIONS；不处理会撞成 405。"""
    resp = local_client.options(
        "/strip.txt",
        headers={"Origin": "http://127.0.0.1:19387", "Access-Control-Request-Method": "GET"},
    )
    assert resp.status_code == 204
    assert resp.headers.get("access-control-allow-origin") == "http://127.0.0.1:19387"
    assert "GET" in resp.headers.get("access-control-allow-methods", "")


def test_other_endpoints_do_not_get_cors(local_client: TestClient) -> None:
    """只给状态条开口子，别的端点不给——面包屑越少越好。"""
    resp = local_client.get("/state", headers={"Origin": "http://127.0.0.1:19387"})
    assert "access-control-allow-origin" not in resp.headers
