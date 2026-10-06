"""引擎：把工具调用包成"可见的一步操作"。

职责
----
1. 校验参数 → 在 COM 线程上执行 → 归一化结果；
2. 全程广播事件：``op.start`` / ``op.end`` / ``state``，这是观察台与任务窗格的数据来源；
3. 为每一步生成**中文人话摘要**——这是"用户能看见 agent 在做什么"的文字层，
   比堆 API 调用有用得多。
"""

from __future__ import annotations

import itertools
import threading
import time
from dataclasses import dataclass, field
from typing import Any

from .config import DEFAULT_PACE_MS, Settings, settings as default_settings
from .errors import PptAgentError, to_error
from .events import EventBus, utc_now
from .registry import Context, Tool, all_tools, get_tool, validate_args

__all__ = ["Engine", "describe_result"]


# --------------------------------------------------------------------------
# 人话摘要
# --------------------------------------------------------------------------


def describe_result(t: Tool, result: Any) -> str:
    """把结构化结果翻成一句中文，供事件流与观察台显示。"""
    if not isinstance(result, dict):
        return f"{t.title}完成"

    if t.name == "ppt_open":
        pres = result.get("presentation") or {}
        name = pres.get("name") or "演示"
        if result.get("already_open"):
            return f"《{name}》本来就开着，已切到它"
        slides = pres.get("slides")
        suffix = f"，共 {slides} 页" if slides is not None else ""
        return f"已打开《{name}》{suffix}"

    if t.name == "ppt_close":
        pres = result.get("presentation") or {}
        name = pres.get("name") or "演示"
        return f"已关闭《{name}》" + ("（已保存）" if result.get("saved") else "（未保存的改动已丢弃）")

    if t.name == "ppt_new":
        deck = (result.get("presentation") or {}).get("name") or "新演示"
        return f"已新建《{deck}》" + ("（套用了模板）" if result.get("templated") else "")

    if t.name == "ppt_activate":
        return "PowerPoint 已切到前台"

    if t.name == "ppt_status":
        com = result.get("com") or {}
        decks = result.get("presentations") or []
        if result.get("com_error"):
            return f"PowerPoint 不可用：{result['com_error']}"
        version = com.get("version") or "?"
        slide = result.get("active_slide")
        where = f"，当前第 {slide} 页" if slide else ""
        return f"PowerPoint {version} 正常，打开着 {len(decks)} 个演示{where}"

    if t.name == "ppt_decks":
        decks = result.get("decks") or []
        names = "、".join(str(d.get("name")) for d in decks[:3])
        more = f" 等 {len(decks)} 个" if len(decks) > 3 else ""
        return f"打开着 {len(decks)} 个演示：{names}{more}" if decks else "没有打开任何演示"

    if t.name == "ppt_slides":
        deck = (result.get("deck") or {}).get("name") or "演示"
        return f"《{deck}》共 {result.get('total')} 页，本次列出 {result.get('returned')} 页"

    if t.name == "ppt_read_slide":
        return f"读完第 {result.get('slide')} 页：{result.get('returned')} 个形状"

    if t.name == "ppt_goto":
        return f"已切到第 {result.get('current_slide')} 页"

    if t.name == "ppt_shot":
        shots = result.get("shots") or []
        first = shots[0]["slide"] if shots else "?"
        if len(shots) == 1:
            return f"已截取第 {first} 页"
        return f"已截取第 {first} 页起共 {len(shots)} 页"

    # -- M2 编辑类 ---------------------------------------------------------

    if t.name == "ppt_save":
        if result.get("copied"):
            return f"已备份到 {result.get('copy_path')}"
        deck = (result.get("presentation") or {}).get("name") or "演示"
        return f"已保存《{deck}》" + (f"（另存为 {result['saved_as']}）" if result.get("saved_as") else "")

    if t.name == "ppt_add_slide":
        return f"已新增第 {result.get('slide')} 页（版式：{result.get('layout')}），现在共 {result.get('total')} 页"

    if t.name == "ppt_delete_slide":
        count = result.get("deleted") or 1
        where = f"第 {result.get('from')} 页" if count == 1 else f"第 {result.get('from')} 页起 {count} 页"
        return f"已删除{where}，现在共 {result.get('total')} 页"

    if t.name == "ppt_move_slide":
        return f"已把第 {result.get('from')} 页移到第 {result.get('to')} 页"

    if t.name == "ppt_set_text":
        return f"已改写「{result.get('shape')}」的文字（{result.get('chars')} 字）"

    if t.name == "ppt_add_textbox":
        return f"已加文本框「{result.get('shape')}」"

    if t.name == "ppt_set_geometry":
        return f"已调整「{result.get('shape')}」的{'、'.join(result.get('changed') or [])}"

    if t.name == "ppt_delete_shape":
        return f"已删除形状「{result.get('shape')}」"

    if t.name == "ppt_add_picture":
        return f"已插入图片「{result.get('shape')}」"

    if t.name == "ppt_set_notes":
        chars = len(result.get("after") or "")
        return f"已写入第 {result.get('slide')} 页备注（{chars} 字）"

    if t.name == "ppt_set_table_cell":
        return f"已修改第 {result.get('slide')} 页「{result.get('shape')}」第 {result.get('row')} 行第 {result.get('column')} 列"

    if t.name == "ppt_select":
        target = result.get("selected")
        return f"已切到第 {result.get('current_slide')} 页" + (f"并选中 {target}" if target else "")

    if t.name == "ppt_undo":
        done = result.get("undone") or 0
        text = f"已撤销 {done} 步"
        if result.get("stopped"):
            text += f"（{result['stopped']}）"
        return text

    if t.name == "ppt_set_pace":
        return f"跟速已设为 {result.get('pace_ms')} 毫秒（{result.get('label')}）"

    return f"{t.title}完成"


def describe_target(t: Tool, args: dict[str, Any]) -> dict[str, Any] | None:
    """从参数里挑出"这一步在动什么"，方便观察台高亮。"""
    if not args:
        return None
    keys = ("path", "name", "slide", "shape", "index", "deck")
    picked = {k: args[k] for k in keys if args.get(k) not in (None, "")}
    return picked or None


# --------------------------------------------------------------------------
# 引擎
# --------------------------------------------------------------------------


@dataclass
class Engine:
    """守护进程的内核。一个进程一个 Engine。"""

    settings: Settings
    bus: EventBus
    session: Any  # ComSession —— 用 Any 避免 registry 反向依赖 com
    token: str = ""
    state: dict[str, Any] = field(default_factory=dict)
    started_at: str = field(default_factory=utc_now)

    _op_seq: itertools.count = field(default_factory=lambda: itertools.count(1), init=False)
    _lock: threading.Lock = field(default_factory=threading.Lock, init=False)
    _invoke_lock: threading.RLock = field(default_factory=threading.RLock, init=False)

    # -- 构造 ---------------------------------------------------------------

    @classmethod
    def create(
        cls,
        settings: Settings | None = None,
        *,
        bus: EventBus | None = None,
        token: str = "",
        session: Any | None = None,
    ) -> "Engine":
        from .com import ComSession
        from .config import new_token

        cfg = settings or default_settings
        if session is None:
            session = ComSession(
                visible=True,           # 可见性铁律
                alerts_off=False,       # 保留用户的全局警告设置，工具显式避免另存为/关闭弹窗
                attach_preferred=cfg.attach_preferred,
                call_timeout=cfg.call_timeout,
            )
        return cls(
            settings=cfg,
            bus=bus or EventBus(),
            session=session,
            token=token or new_token(),
            state={"active_deck": None, "attached": None, "pace_ms": DEFAULT_PACE_MS},
        )

    def context(self) -> Context:
        return Context(session=self.session, bus=self.bus, settings=self.settings, state=self.state)

    # -- 生命周期 -----------------------------------------------------------

    def start(self) -> dict[str, Any]:
        """启动 COM 会话并广播初始状态。"""
        info = self.session.start()
        self.state["attached"] = info.attached
        self.bus.publish(
            "state",
            com=info.to_dict(),
            active_deck=self.state.get("active_deck"),
            started_at=self.started_at,
        )
        return info.to_dict()

    def stop(self) -> None:
        self.session.stop()
        self.bus.publish("log", level="info", message="守护进程正在停止")

    # -- 调用 ---------------------------------------------------------------

    def invoke(self, name: str, args: dict[str, Any] | None = None, *, source: str = "api") -> dict[str, Any]:
        """执行一个工具，返回 ``{ok, tool, result|error, ...}``（不抛异常）。"""
        # 事件、操作和跟速属于同一步；只串行 COM 不能防止事件交错或并发绕过停顿。
        deadline = time.monotonic() + self.settings.call_timeout
        if not self._invoke_lock.acquire(timeout=self.settings.call_timeout):
            return {"ok": False, "tool": name, "error": {
                "code": "operation/queue-timeout", "message": "等待前一步操作超时，本次操作未执行。",
            }}
        try:
            return self._invoke(name, args, source=source, deadline=deadline)
        finally:
            self._invoke_lock.release()

    def _invoke(self, name: str, args: dict[str, Any] | None, *, source: str, deadline: float) -> dict[str, Any]:
        try:
            t = get_tool(name)
            checked = validate_args(t, args)
        except PptAgentError as exc:
            return {"ok": False, "tool": name, "error": exc.to_dict()}
        if time.monotonic() >= deadline:
            return {"ok": False, "tool": name, "error": {
                "code": "operation/queue-timeout", "message": "等待前一步操作超时，本次操作未执行。",
            }}

        op_id = f"{next(self._op_seq):04d}"
        ctx = self.context()
        started = time.perf_counter()
        self.bus.publish(
            "op.start",
            id=op_id,
            tool=t.name,
            title=t.title,
            summary=t.summary,      # 中文人话：这一步打算做什么
            behavior=t.behavior,
            source=source,
            target=describe_target(t, checked),
            args=checked,
        )

        try:
            if t.needs_com:
                if not self.session.info().alive:
                    raise PptAgentError(
                        self.session.info().last_error or "COM 会话未启动",
                        code="com/unavailable",
                    )
                result = self.session.run(
                    self._run_tool, t, checked, ctx, timeout=max(0.001, deadline - time.monotonic()),
                    retry_safe=t.behavior in {"read", "idempotent"},
                )
            else:
                result = t.handler(None, checked, ctx)
        except BaseException as exc:  # noqa: BLE001 - 引擎边界，统一归一化
            error = to_error(exc)
            elapsed = round((time.perf_counter() - started) * 1000, 1)
            self.bus.publish(
                "op.end",
                id=op_id,
                tool=t.name,
                title=t.title,
                ok=False,
                ms=elapsed,
                error=error,
                human=f"{t.title}失败：{error['message']}",
            )
            return {"ok": False, "tool": t.name, "error": error, "op_id": op_id, "ms": elapsed}

        elapsed = round((time.perf_counter() - started) * 1000, 1)
        human = describe_result(t, result)
        self.bus.publish(
            "op.end",
            id=op_id,
            tool=t.name,
            title=t.title,
            ok=True,
            ms=elapsed,
            human=human,
            result=result if _small(result) else None,
        )
        # 截图单独发一条事件：观察台据此把画面挂到这一步上，并更新大图。
        if isinstance(result, dict) and result.get("shots"):
            self.bus.publish("op.shot", id=op_id, tool=t.name, shots=result["shots"])
        if t.needs_com:
            self.bus.publish("state", **self.session_state())

        self._pace(t)
        return {"ok": True, "tool": t.name, "result": result, "op_id": op_id, "ms": elapsed, "human": human}

    def observe_view(self, app: Any) -> None:
        """只在 STA 上调用：缓存真实活动演示和页码，面板不能把旧事件当现状。"""
        from .ops.common import active_slide_index, presentation_info, safe

        pres = safe(lambda: app.ActivePresentation)
        info = presentation_info(pres) if pres is not None else {}
        with self._lock:
            self.state["active_deck"] = info.get("path") or info.get("name")
            self.state["active_slide"] = active_slide_index(app) if pres is not None else None
            self.state["view_known"] = True

    def _run_tool(self, app: Any, t: Tool, args: dict[str, Any], ctx: Context) -> Any:
        try:
            return t.handler(app, args, ctx)
        finally:
            self.observe_view(app)

    def _pace(self, t: Tool) -> None:
        """写操作之后停一下，让人跟得上。

        放在事件全部广播之后：用户先看到这一步的结果，下一步才迟迟到来——
        否则连续操作会在观察台上糊成一团。
        """
        if t.behavior not in {"write", "destroy"}:
            return
        pace_ms = int(self.state.get("pace_ms") or 0)
        if pace_ms > 0:
            time.sleep(pace_ms / 1000.0)

    # -- 清单 ---------------------------------------------------------------

    def manifest(self) -> dict[str, Any]:
        return {
            "tools": [t.to_manifest() for t in all_tools()],
            "count": len(all_tools()),
        }

    def session_state(self) -> dict[str, Any]:
        return {
            "com": self.session.info().to_dict(),
            "active_deck": self.state.get("active_deck"),
            "active_slide": self.state.get("active_slide"),
            "view_known": self.state.get("view_known", False),
            "pace_ms": self.state.get("pace_ms"),
            "started_at": self.started_at,
            "tools": len(all_tools()),
            "subscribers": self.bus.subscriber_count,
        }


def _small(value: Any, limit: int = 4000) -> bool:
    """结果太大就不塞进事件流（观察台只做摘要，细节走工具返回值）。"""
    try:
        import json

        return len(json.dumps(value, ensure_ascii=False, default=str)) <= limit
    except Exception:  # noqa: BLE001
        return False
