"""HTTP 面：观察台 + 通用 API + 每个工具一个 REST 端点。

为什么要动态生成每个工具的端点
------------------------------
ChatGPT 侧要接的是 OpenAPI。一个笼统的 ``POST /call {name, args}`` 在
GPT Actions 里很难用（模型得靠字符串拼工具名），而**每个工具一个端点**
能让 OpenAPI 直接给出强类型 schema，接起来最省事。两个都提供。

本面同时服务观察台页面（可见面②）与 PowerShell / 任务窗格 / ChatGPT。
"""

from __future__ import annotations

import asyncio
import json
import os
import queue
import re
import threading
from pathlib import Path
from typing import Any, Optional

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from pydantic import BaseModel, ConfigDict, Field, create_model
from sse_starlette.sse import EventSourceResponse
from starlette.concurrency import run_in_threadpool

from . import __version__
from .addin_https import manifest_xml
from .addin_page import icon_png, render_taskpane
from .config import ensure_dirs
from .engine import Engine
from .observer import render_observer
from .ops.deck import LIVE_NO_PRESENTATION, LIVE_NO_SLIDES, capture_live_frame
from .registry import Param, Tool, all_tools

__all__ = ["create_app"]

#: 观察台页面本身也要令牌，否则同机的任意网页都能驱动 PowerPoint。
_OPEN_PATHS = {"/health"}

#: 任务窗格由 WebView 直接打开，没法带自定义请求头，所以那一组端点不能强制令牌。
#: 代价是：**一旦端口被隧道到公网**，任何人都能拿到注入在页面里的令牌。
#: 所以只用"Host 是不是本机"来放行——隧道之后 Host 会变成公网域名，必须带令牌。
_LOCAL_HOSTNAMES = {"localhost", "127.0.0.1", "::1", "[::1]"}

#: 允许跨源读取 ``/strip.txt`` 的**来源**白名单。
#:
#: 为什么需要：状态条挂在宿主界面里，而宿主界面跑在**另一个端口**（DSH 默认 19387）。
#: 从 ``http://127.0.0.1:19387`` 去 fetch ``http://127.0.0.1:8791/strip.txt`` 是跨源请求，
#: 浏览器会因为响应缺少 ``Access-Control-Allow-Origin`` 直接拦掉——
#: 表现成状态条**永远显示"未连接"，而守护进程其实好好地在跑**。
#: （实测确认：``TypeError: Failed to fetch``。）
#:
#: 为什么不用 ``*``：那等于让**任何**网站都能读到你的演示名和页码。
#: 只放行本机来源，既够用，又不把信息送给互联网上的任意页面。
_LOCAL_ORIGIN = re.compile(r"^https?://(localhost|127\.0\.0\.1|\[::1\])(:\d+)?$")


def _is_local_request(request: Request) -> bool:
    host = (request.headers.get("host") or "").rsplit(":", 1)[0].strip().lower()
    return host in _LOCAL_HOSTNAMES


def _local_origin(request: Request) -> str | None:
    """请求来自本机页面时返回它的 Origin，否则 ``None``。"""
    origin = (request.headers.get("origin") or "").strip()
    return origin if _LOCAL_ORIGIN.match(origin) else None


class CallBody(BaseModel):
    """``POST /call`` 的请求体。

    必须定义在**模块级**：本模块用了 ``from __future__ import annotations``，
    若把它定义在 ``create_app`` 内部，FastAPI 解析注解时会去模块全局找而找不到，
    于是把 ``body`` 误判成查询参数。
    """

    name: str = Field(..., description="工具名，例如 ppt_open")
    args: dict[str, Any] = Field(default_factory=dict, description="工具参数")


class PaceBody(BaseModel):
    """``POST /pace`` 的请求体。同样必须定义在模块级，理由见 ``CallBody``。"""

    pace_ms: int = Field(..., ge=0, le=10000, description="写操作之间的停顿毫秒数；0=极速，400≈正常，1500=慢速")


def _extract_token(request: Request) -> str | None:
    header = request.headers.get("x-ppt-token")
    if header:
        return header.strip()
    auth = request.headers.get("authorization") or ""
    if auth.lower().startswith("bearer "):
        return auth[7:].strip()
    query = request.query_params.get("token")
    return query.strip() if query else None


def _guard(request: Request, engine: Engine) -> None:
    """除白名单外一律校验令牌。"""
    if request.url.path in _OPEN_PATHS:
        return
    if _extract_token(request) != engine.token:
        raise HTTPException(status_code=401, detail="缺少或错误的访问令牌（用 pptctl url 取正确地址）")


def _body_model(t: Tool) -> type[BaseModel]:
    """把工具的 Param 列表变成 pydantic 模型，OpenAPI 里就是强类型 body。"""
    fields: dict[str, Any] = {}
    for p in t.params:
        py = p.to_annotation()
        if p.required:
            fields[p.name] = (py, Field(..., description=p.description))
        else:
            fields[p.name] = (
                py,
                Field(p.default, description=p.description),
            )
    return create_model(f"{t.name}_body", __config__=ConfigDict(extra="forbid"), **fields)


def _make_endpoint(engine: Engine, t: Tool, model: type[BaseModel]):
    async def endpoint(request, body):  # type: ignore[no-untyped-def]
        _guard(request, engine)   # 这些端点同样是"能驱动 PowerPoint"的入口，必须鉴权
        payload = body.model_dump()
        result = await run_in_threadpool(engine.invoke, t.name, payload, source="http")
        return JSONResponse(result, status_code=200 if result.get("ok") else 400)

    endpoint.__name__ = t.name
    endpoint.__doc__ = t.summary
    # 关键在于把注解直接设成类对象：FastAPI 才能据此生成 body schema，
    # 并把 request 认成 Request 而不是查询参数。
    endpoint.__annotations__ = {"request": Request, "body": model}
    return endpoint


def _plain(text: str, request: Request | None = None) -> Response:
    """一行纯文本响应，且不许缓存——状态条每几秒就要刷一次。

    本机来源的跨源请求会拿到 CORS 头：状态条挂在**宿主界面（另一个端口）**上，
    少了这个头浏览器会直接拦掉响应，表现成"永远连不上"。
    """
    headers = {"Cache-Control": "no-store"}
    if request is not None:
        origin = _local_origin(request)
        if origin:
            headers["Access-Control-Allow-Origin"] = origin
            headers["Vary"] = "Origin"
    return Response(content=text + "\n", media_type="text/plain; charset=utf-8", headers=headers)


def _deck_label(path: Any) -> str:
    """面板上只显示文件名，不铺一整条路径。"""
    if not path:
        return ""
    return str(path).replace("\\", "/").rsplit("/", 1)[-1]


def _event_payloads(event: Any) -> list[dict[str, Any]]:
    """一条事件里可能藏信息的几层：data、data.result，以及结果里的演示对象。"""
    data = event.data or {}
    out = [data]
    result = data.get("result")
    if isinstance(result, dict):
        out.append(result)
        for key in ("presentation", "decks", "presentations"):
            nested = result.get(key)
            if isinstance(nested, dict):
                out.append(nested)
            elif isinstance(nested, list) and nested and isinstance(nested[0], dict):
                out.append(nested[0])
    return out


def _derive_from_events(events: list[Any]) -> tuple[str, int | None]:
    """从事件流里回溯"最后已知的演示与页码"。

    ``state`` 事件只在打开/关闭演示时带 deck，页码更是只出现在工具结果里；
    面板要显示"现在在哪一页"，就得自己往回翻。
    """
    deck = ""
    slide: int | None = None
    for event in reversed(events):
        for payload in _event_payloads(event):
            if not deck:
                for key in ("path", "name", "deck"):
                    value = payload.get(key)
                    if isinstance(value, str) and value:
                        deck = _deck_label(value)
                        break
            if slide is None:
                for key in ("active_slide", "current_slide", "slide"):
                    value = payload.get(key)
                    if isinstance(value, int):
                        slide = value
                        break
        if deck and slide is not None:
            break
    return deck, slide


def _event_text(event: Any) -> str:
    """每条事件压成一行**中文人话**。"""
    data = event.data or {}
    if event.type == "op.end":
        mark = "OK" if data.get("ok") else "ERR"
        return f"{mark} {data.get('human') or data.get('tool') or ''}".strip()
    if event.type == "op.start":
        return str(data.get("summary") or data.get("title") or data.get("tool") or "")
    if event.type == "op.shot":
        shots = data.get("shots") or []
        first = shots[0].get("slide") if shots else ""
        return f"已截取第 {first} 页" if first != "" else "已截图"
    if event.type == "log":
        return str(data.get("message") or "")
    if event.type == "state":
        return f"状态更新 deck={_deck_label(data.get('active_deck'))}"
    return ""


def create_app(engine: Engine) -> FastAPI:
    app = FastAPI(
        title="ppt-agent",
        version=__version__,
        description=(
            "让 agent 实时控制正在运行的 PowerPoint。\n\n"
            "同一套工具同时暴露为 MCP 与 HTTP/OpenAPI；所有操作都会广播到观察台，"
            "用户随时看得见 agent 在做什么。\n\n"
            "**鉴权**：除 `/health` 外都需要令牌，用 `X-PPT-Token` 头或 `?token=` 传入；"
            "令牌在守护进程的 runtime.json 里，用 `pptctl url` 可打印观察台地址。"
        ),
        docs_url="/docs",
        openapi_url="/openapi.json",
    )
    app.state.engine = engine
    app.state.uvicorn_server = None  # 由 daemon.serve() 注入，供 /shutdown 使用
    #: 实时取景的闸门：上一帧没拍完就丢掉这一拍，避免请求堆在 STA 队列上
    app.state.live_lock = threading.Lock()

    # -- 观察台 ----------------------------------------------------------

    @app.get("/", response_class=HTMLResponse, include_in_schema=False)
    async def observer(request: Request) -> HTMLResponse:
        """观察台页面。

        与 ``/addin/`` 同一条规矩：**本机 Host 免令牌**（用户直接敲地址、或从宿主界面
        点进来时带不了令牌），非本机 Host 必须带令牌——隧道之后 Host 会变成公网域名。
        """
        if not _is_local_request(request):
            _guard(request, engine)
        return HTMLResponse(render_observer(engine.token, engine.settings.host, engine.settings.port))

    # -- 健康与状态 ------------------------------------------------------

    @app.get("/health", summary="存活探针（唯一无需令牌的端点）", tags=["meta"])
    async def health() -> dict[str, Any]:
        info = engine.session.info()
        return {
            "ok": True,
            "name": "ppt-agent",
            "version": __version__,
            "pid": os.getpid(),
            "port": engine.settings.port,
            "com": {"alive": info.alive, "attached": info.attached, "version": info.version},
        }

    @app.get("/state", summary="守护进程与 PowerPoint 的当前状态", tags=["meta"])
    async def state(request: Request) -> dict[str, Any]:
        _guard(request, engine)
        return engine.session_state()

    @app.get("/tools", summary="列出全部工具及其 JSON Schema", tags=["meta"])
    async def tools(request: Request) -> dict[str, Any]:
        _guard(request, engine)
        return engine.manifest()

    # -- 截图文件 --------------------------------------------------------

    @app.get("/shots/{name}", summary="取一张截图文件（观察台与任务窗格用）", tags=["meta"])
    async def shot_file(request: Request, name: str) -> FileResponse:
        _guard(request, engine)
        # 只允许取截图目录里的**文件名**，杜绝路径穿越。
        if not name or name != Path(name).name or name.startswith("."):
            raise HTTPException(status_code=404, detail="非法的文件名")
        directory = engine.settings.shot_dir.resolve()
        target = (directory / name).resolve()
        if directory != target.parent or not target.is_file():
            raise HTTPException(status_code=404, detail="截图不存在")
        return FileResponse(target)

    # -- Office 加载项（任务窗格）----------------------------------------

    @app.get("/addin/", response_class=HTMLResponse, include_in_schema=False)
    async def addin_page(request: Request) -> HTMLResponse:
        """任务窗格页面。

        WebView 直接打开这个地址，**没法带自定义请求头**，所以这里不能强制令牌。
        代价用"Host 是不是本机"来兜：一旦端口被隧道到公网，Host 会变成公网域名，
        就必须带令牌——否则等于把注入在页面里的令牌泄露给全世界。
        """
        if not _is_local_request(request):
            _guard(request, engine)
        return HTMLResponse(render_taskpane(engine.token, engine.settings.host, engine.settings.port))

    @app.get("/addin/manifest.xml", include_in_schema=False)
    async def addin_manifest(request: Request) -> Response:
        """加载项清单。地址与端口是**按当前配置生成**的，不会和实际监听不一致。"""
        if not _is_local_request(request):
            _guard(request, engine)
        return Response(
            content=manifest_xml(engine.settings),
            media_type="application/xml; charset=utf-8",
        )

    @app.get("/addin/icon-{size}.png", include_in_schema=False)
    async def addin_icon(request: Request, size: int) -> Response:
        if size not in {16, 32, 80}:
            raise HTTPException(status_code=404, detail="只提供 16 / 32 / 80 三种尺寸")
        if not _is_local_request(request):
            _guard(request, engine)
        return Response(
            content=icon_png(size),
            media_type="image/png",
            headers={"Cache-Control": "max-age=86400"},
        )

    # -- 实时取景 --------------------------------------------------------

    @app.get("/live", summary="实时取景：把此刻这一页现拍一张 JPEG", tags=["meta"])
    async def live(request: Request, width: int = 960, slide: int | None = None) -> Response:
        """给观察台当实时监视器用。

        ``ppt_shot`` 是留档快照，这个是**取景**：覆盖同一个文件、不产缩略图、不进事件流，
        所以可以每隔一两秒调一次。它读的是 PowerPoint 的当前状态，因此**用户在
        PowerPoint 里自己翻页或改内容，观察台也看得见**。
        """
        _guard(request, engine)
        if not engine.session.info().alive:
            raise HTTPException(status_code=503, detail="COM 会话不可用")
        # 上一帧还没拍完就直接放弃这一拍：宁可少一帧，也不要把请求堆在 STA 队列上
        if not app.state.live_lock.acquire(blocking=False):
            return Response(status_code=204, headers={"Cache-Control": "no-store"})
        try:
            target = engine.settings.shot_dir / "live.jpg"
            ensure_dirs(engine.settings)
            def capture(app):
                index = capture_live_frame(app, target, width=width, slide=slide)
                engine.observe_view(app)
                return index
            index = await run_in_threadpool(
                engine.session.run, capture, timeout=25.0
            )
            if index == LIVE_NO_PRESENTATION:
                raise HTTPException(status_code=409, detail="当前没有打开任何演示")
            if index == LIVE_NO_SLIDES:
                raise HTTPException(status_code=409, detail="当前演示还没有任何幻灯片")
            return FileResponse(
                target,
                media_type="image/jpeg",
                headers={"Cache-Control": "no-store", "X-PPT-Slide": str(index)},
            )
        except HTTPException:
            raise
        except Exception as exc:  # noqa: BLE001 - 取景失败不该打断观察台
            raise HTTPException(status_code=503, detail=f"实时取景失败：{type(exc).__name__}: {exc}") from exc
        finally:
            app.state.live_lock.release()

    # -- 面板数据源（轮询式客户端用）-------------------------------------

    @app.get("/feed", summary="面板快照：状态 + 最近的操作流水", tags=["meta"])
    async def feed(request: Request, limit: int = 40) -> dict[str, Any]:
        """一次请求拿全面板要显示的东西。

        SSE（``/events``）适合网页；但 VBA、PowerShell 这类客户端不方便消费长连接，
        所以同样一份数据再给一个轮询端点——PowerPoint 里的加载项面板就靠它。
        """
        _guard(request, engine)
        count = max(1, min(int(limit), 200))
        recent = engine.bus.recent(count)
        return {
            "ok": True,
            "state": engine.session_state(),
            "count": len(recent),
            "events": [e.to_dict() for e in recent],
        }

    @app.get("/feed.txt", summary="面板快照（纯文本，给 VBA/PowerShell 用）", tags=["meta"])
    async def feed_text(request: Request, limit: int = 40) -> Response:
        """与 ``/feed`` 同一份数据，但用**纯文本**。

        为什么要有这个：VBA 里解析 JSON 要手写解析器，脆弱又长。纯文本格式
        一行一条、用 `` | `` 分隔，VBA 一句 ``Split`` 就能用；PowerShell 里
        ``curl`` 直接看得懂，调试也方便。
        """
        _guard(request, engine)
        count = max(1, min(int(limit), 200))
        state = engine.session_state()
        com = state.get("com") or {}
        events = engine.bus.recent(count)
        # state 里没有的就从事件流回溯——面板要显示"现在在哪一页"
        derived_deck, derived_slide = _derive_from_events(events)
        deck = _deck_label(state.get("active_deck")) if state.get("view_known") else (_deck_label(state.get("active_deck")) or derived_deck)
        if state.get("view_known"):
            derived_slide = state.get("active_slide")

        lines = [
            "PPT-AGENT-FEED v1",
            f"alive={str(bool(com.get('alive'))).lower()}",
            f"version={com.get('version') or ''}",
            f"visible={str(bool(com.get('visible'))).lower()}",
            f"degraded={str(bool(com.get('degraded'))).lower()}",
            f"pace_ms={state.get('pace_ms') or 0}",
            f"tools={state.get('tools') or 0}",
            f"deck={deck}",
            f"slide={derived_slide if derived_slide is not None else ''}",
            f"count={len(events)}",
            "---",
        ]
        for event in events:
            lines.append(f"{event.at} | {event.type} | {_event_text(event)}")
        body = "\n".join(lines) + "\n"
        return Response(
            content=body,
            media_type="text/plain; charset=utf-8",
            headers={"Cache-Control": "no-store"},
        )

    # -- 状态条：给宿主界面里的一条窄条用 ---------------------------------

    @app.get("/strip.txt", summary="一行状态（给宿主 UI 的窄条用）", tags=["meta"])
    async def strip(request: Request) -> Response:
        """一行纯文本，服务端已经拼好。

        为什么单独开一个端点而不是让客户端读 ``/feed.txt``：

        - 宿主界面里的组件拿不到令牌（它跑在渲染进程里，读不到 runtime.json），
          而让宿主去实现"取令牌"的握手会引入一堆我验证不了的契约；
        - 这一行**不含任何密钥**，只有版本、演示名、页码、跟速和最近一步的中文摘要。

        所以它按 ``/addin/`` 同样的规矩放行：只认本机 Host。**端口一旦被隧道到公网，
        Host 会变成公网域名，这里就会要求令牌。**

        另外还要放行**本机来源的跨源读取**（见 :data:`_LOCAL_ORIGIN`）——
        宿主界面在另一个端口上，不放行的话浏览器会把这个响应整个拦掉。
        """
        if not _is_local_request(request):
            _guard(request, engine)

        state = engine.session_state()
        com = state.get("com") or {}
        if not com.get("alive"):
            return _plain("○ ppt-agent 未连接 —— 先运行 pptctl serve", request)

        events = engine.bus.recent(20)
        deck, slide = _derive_from_events(events)
        if state.get("view_known"):
            deck, slide = _deck_label(state.get("active_deck")), state.get("active_slide")
        parts = [f"● PowerPoint {com.get('version') or '?'}"]
        if not com.get("visible"):
            parts.append("窗口不可见")
        parts.append(deck or "没有打开的演示")
        if slide is not None:
            parts.append(f"第 {slide} 页")
        if state.get("pace_ms"):
            parts.append(f"跟速 {state['pace_ms']}ms")

        latest = next((e for e in reversed(events) if e.type == "op.end"), None)
        if latest is not None and latest.data.get("human"):
            mark = "" if latest.data.get("ok") else "失败 "
            parts.append(f"{mark}{latest.data['human']}")
        else:
            running = next((e for e in reversed(events) if e.type == "op.start"), None)
            if running is not None:
                parts.append(f"正在进行：{running.data.get('summary') or running.data.get('tool')}")

        text = " · ".join(str(p) for p in parts if p)
        if com.get("degraded"):
            text = "⚠ " + text + "（COM 已降级，建议 pptctl restart）"
        return _plain(text, request)

    @app.options("/strip.txt", include_in_schema=False)
    async def strip_preflight(request: Request) -> Response:
        """预检。

        状态条的 fetch 是"简单请求"，本来不需要预检；但宿主将来若加了自定义头
        （比如带个客户端标识），浏览器就会先发 OPTIONS——不处理的话会撞成 405。
        """
        headers = {"Access-Control-Allow-Methods": "GET, OPTIONS", "Vary": "Origin"}
        origin = _local_origin(request)
        if origin:
            headers["Access-Control-Allow-Origin"] = origin
            headers["Access-Control-Allow-Headers"] = request.headers.get(
                "access-control-request-headers", "*"
            )
        return Response(status_code=204, headers=headers)

    # -- 通用调用 --------------------------------------------------------

    @app.post("/call", summary="按名字调用任意工具", tags=["invoke"])
    async def call(request: Request, body: CallBody) -> JSONResponse:
        _guard(request, engine)
        result = await run_in_threadpool(engine.invoke, body.name, body.args, source="http")
        return JSONResponse(result, status_code=200 if result.get("ok") else 400)

    # -- 事件流（观察台的数据源）-----------------------------------------

    @app.get("/events", summary="SSE 事件流：agent 每一步操作", tags=["meta"])
    async def events(request: Request) -> EventSourceResponse:
        _guard(request, engine)
        bus = engine.bus

        async def generator():
            q = bus.subscribe()
            try:
                for event in bus.recent(40):
                    yield {"event": event.type, "data": json.dumps(event.to_dict(), ensure_ascii=False)}
                while True:
                    if await request.is_disconnected():
                        break
                    try:
                        event = await asyncio.to_thread(q.get, True, 15.0)
                    except queue.Empty:
                        yield {"event": "ping", "data": "{}"}
                        continue
                    except Exception:  # noqa: BLE001 - 客户端断开
                        break
                    yield {"event": event.type, "data": json.dumps(event.to_dict(), ensure_ascii=False)}
            finally:
                bus.unsubscribe(q)

        return EventSourceResponse(generator())

    # -- 跟速 ------------------------------------------------------------

    @app.post("/pace", summary="设置跟速（观察台与任务窗格用）", tags=["meta"])
    async def set_pace(request: Request, body: PaceBody) -> dict[str, Any]:
        _guard(request, engine)
        engine.state["pace_ms"] = int(body.pace_ms)
        engine.bus.publish("state", pace_ms=int(body.pace_ms))
        return {"ok": True, "pace_ms": int(body.pace_ms)}

    # -- 急停 ------------------------------------------------------------

    @app.post("/shutdown", summary="急停：停止守护进程（不会关闭 PowerPoint）", tags=["meta"])
    async def shutdown(request: Request) -> dict[str, Any]:
        _guard(request, engine)
        engine.bus.publish("log", level="warn", message="收到急停请求，守护进程即将退出")

        def _stop() -> None:
            server = app.state.uvicorn_server
            if server is not None:
                server.should_exit = True

        asyncio.get_running_loop().call_later(0.4, _stop)
        return {"ok": True, "stopping": True}

    # -- 每个工具一个端点 -------------------------------------------------

    for t in all_tools():
        app.add_api_route(
            f"/tools/{t.name}",
            _make_endpoint(engine, t, _body_model(t)),
            methods=["POST"],
            name=t.title,
            operation_id=t.name,
            summary=t.summary,
            description=(
                f"{t.summary}\n\n"
                f"- 行为：`{t.behavior}`（read=只读，write=写入，destroy=破坏性）\n"
                f"- 返回：{t.returns}"
            ),
            tags=[t.behavior],
            response_model=None,
            responses={
                200: {"description": "成功"},
                400: {"description": "工具报错（响应体里带 error.code / error.message）"},
            },
        )

    return app
