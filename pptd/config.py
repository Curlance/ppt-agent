"""全局配置与运行时发现。

设计要点
--------
1. **解释器被显式固定**：使用安装了项目依赖和 pywin32 的 Python 环境。
   所有入口复用该解释器；``interpreter_report()`` 负责自检。
2. **运行时状态落在用户目录**，不污染仓库：``%LOCALAPPDATA%\\ppt-agent\\runtime.json``
   记录 pid / 端口 / 令牌，供 CLI、MCP 转发器与加载项发现已经活着的守护进程。
3. **默认只绑定回环地址**，并给写操作加令牌，避免同机其它程序静默驱动 PowerPoint。
"""

from __future__ import annotations

import json
import os
import secrets
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

APP_NAME = "ppt-agent"

#: 写操作之间的默认停顿（毫秒）。0=极速；400≈人跟得上；1500=慢速看得清。
#: 这是"用户能看见 agent 在做什么"的节奏控制，不是一个可有可无的装饰。
DEFAULT_PACE_MS = 400

__all__ = [
    "APP_NAME",
    "DEFAULT_PACE_MS",
    "Settings",
    "settings",
    "ensure_dirs",
    "read_runtime",
    "write_runtime",
    "clear_runtime",
    "new_token",
    "interpreter_report",
    "host_interpreter",
]


def _env_path(name: str, default: Path) -> Path:
    raw = os.environ.get(name)
    return Path(raw).expanduser() if raw else default


def _default_home() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return Path(base) / APP_NAME


@dataclass(frozen=True)
class Settings:
    """进程级配置快照。构造后不变，便于测试替换。"""

    host: str
    port: int
    home: Path
    call_timeout: float
    attach_preferred: bool
    #: MCP over HTTP（streamable-http）的端口。0 表示用 ``port + 1``。
    #: 刻意用**独立端口**而不是子路径：挂子路径会踩到「子应用 lifespan 不执行」
    #: 和「/mcp 被重定向到 /mcp/」两个坑，而 MCP 客户端 POST 通常不跟重定向。
    mcp_port: int = 0
    #: Office Web Add-in 任务窗格的 HTTPS 端口。0 表示用 ``port + 2``。
    #: 必须单独一个 HTTPS 端口：任务窗格页面与它要调的 API 必须**同源**，
    #: 否则 HTTPS 页面去 fetch HTTP 会被混合内容拦掉。
    addin_port: int = 0

    # 派生路径
    runtime_file: Path = field(init=False)
    log_dir: Path = field(init=False)
    shot_dir: Path = field(init=False)
    cache_dir: Path = field(init=False)
    pid_file: Path = field(init=False)

    def __post_init__(self) -> None:
        object.__setattr__(self, "runtime_file", self.home / "runtime.json")
        object.__setattr__(self, "log_dir", self.home / "logs")
        object.__setattr__(self, "shot_dir", self.home / "shots")
        object.__setattr__(self, "cache_dir", self.home / "cache")
        object.__setattr__(self, "pid_file", self.home / "daemon.pid")

    @property
    def base_url(self) -> str:
        return f"http://{self.host}:{self.port}"

    @property
    def mcp_bind_port(self) -> int:
        """MCP over HTTP 实际监听的端口。"""
        return self.mcp_port or (self.port + 1)

    @property
    def mcp_url(self) -> str:
        return f"http://{self.host}:{self.mcp_bind_port}/mcp"

    @property
    def addin_bind_port(self) -> int:
        """Office 加载项任务窗格实际监听的 HTTPS 端口。"""
        return self.addin_port or (self.port + 2)

    @property
    def addin_url(self) -> str:
        """任务窗格地址。**必须是 localhost** —— Office 只对 localhost 用本地证书。"""
        return f"https://localhost:{self.addin_bind_port}/addin/"


def _load() -> Settings:
    home = _env_path("PPT_AGENT_HOME", _default_home())
    port = int(os.environ.get("PPT_AGENT_PORT", "8791"))
    return Settings(
        host=os.environ.get("PPT_AGENT_HOST", "127.0.0.1"),
        port=port,
        home=home,
        call_timeout=float(os.environ.get("PPT_AGENT_CALL_TIMEOUT", "120")),
        attach_preferred=os.environ.get("PPT_AGENT_ATTACH", "1") not in {"0", "false", "False"},
        mcp_port=int(os.environ.get("PPT_AGENT_MCP_PORT", "0")) or (port + 1),
        addin_port=int(os.environ.get("PPT_AGENT_ADDIN_PORT", "0")) or (port + 2),
    )


settings = _load()
"""模块级默认配置。CLI 可用 ``dataclasses.replace`` 派生变体。"""


def ensure_dirs(s: Settings | None = None) -> None:
    """确保运行时目录都存在。"""
    cfg = s or settings
    for directory in (cfg.home, cfg.log_dir, cfg.shot_dir, cfg.cache_dir):
        directory.mkdir(parents=True, exist_ok=True)


def new_token() -> str:
    """生成一次性本地访问令牌。"""
    return secrets.token_urlsafe(24)


def pinned_token() -> str:
    """``PPT_AGENT_TOKEN`` 指定的固定令牌；没设就返回空串。

    默认每次启动换一个令牌（更安全），代价是**长驻客户端要自己重读**
    （MCP stdio 进程就踩过这个坑：重启一次就全变"未知错误"）。
    如果你希望令牌跨重启稳定——比如把观察台地址存成书签、或给外部工具配了固定凭据——
    设这个环境变量即可。**注意它会被运行时记录读出来，等同于把凭据落盘。**
    """
    return os.environ.get("PPT_AGENT_TOKEN", "").strip()


def write_runtime(payload: dict[str, Any], s: Settings | None = None) -> Path:
    """原子写入 runtime.json（先写临时文件再替换，避免读到半截）。"""
    cfg = s or settings
    ensure_dirs(cfg)
    tmp = cfg.runtime_file.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, cfg.runtime_file)
    return cfg.runtime_file


def read_runtime(s: Settings | None = None) -> dict[str, Any] | None:
    """读取 runtime.json；不存在或损坏时返回 None（不抛异常）。"""
    cfg = s or settings
    try:
        raw = cfg.runtime_file.read_text(encoding="utf-8")
        data = json.loads(raw)
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) else None


def clear_runtime(s: Settings | None = None) -> None:
    """删除 runtime.json（守护进程退出时调用）。"""
    cfg = s or settings
    for path in (cfg.runtime_file, cfg.pid_file):
        try:
            path.unlink()
        except OSError:
            pass


def interpreter_report() -> dict[str, Any]:
    """自检解释器：路径、版本，以及 pywin32 是否可用。

    这是 M0 最容易踩的坑——用错解释器就会在守护进程启动时才炸。
    """
    exe, extra_env = host_interpreter()
    report: dict[str, Any] = {
        "executable": sys.executable,
        "version": sys.version.split()[0],
        "in_venv": sys.prefix != getattr(sys, "base_prefix", sys.prefix),
        "host_executable": exe,
        "host_env": extra_env,
    }
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401

        report["pywin32"] = True
    except Exception as exc:  # pragma: no cover - 环境相关
        report["pywin32"] = False
        report["pywin32_error"] = f"{type(exc).__name__}: {exc}"
    try:
        import pptx  # noqa: F401

        report["python_pptx"] = True
    except Exception:  # pragma: no cover - 环境相关
        report["python_pptx"] = False
    return report


def _children_of_shim() -> tuple[Path | None, Path | None]:
    """解析 pyvenv.cfg，拿到基础解释器与 venv 的 site-packages。"""
    if sys.prefix == getattr(sys, "base_prefix", sys.prefix):
        return None, None
    site = Path(sys.prefix) / "Lib" / "site-packages"
    config_file = Path(sys.prefix) / "pyvenv.cfg"
    home: Path | None = None
    try:
        for line in config_file.read_text(encoding="utf-8").splitlines():
            key, _, value = line.partition("=")
            if key.strip().lower() == "home":
                home = Path(value.strip())
                break
    except OSError:
        home = None
    if home is None:
        # 退一步：基础解释器通常在 base_prefix 下
        home = Path(getattr(sys, "base_prefix", sys.prefix))
    candidate = home / ("python.exe" if os.name == "nt" else "python3")
    return (candidate if candidate.is_file() else None, site if site.is_dir() else None)


def host_interpreter() -> tuple[str, dict[str, str]]:
    """给宿主编排器用的真实解释器，以及需要补的环境变量。

    **为什么不能用 venv 里的 python.exe**：uv 建的 venv 在该 exe 上套了一层
    再执行的壳（shim）。宿主杀掉壳时，真正的子进程会变成孤儿并继续占着
    stdio 管道，导致宿主关闭时挂起、并残留僵尸进程。这里直接返回基础解释器，
    再用 ``PYTHONPATH`` 把 venv 的包目录挂上去——单进程，且依赖齐全。
    """
    base, site = _children_of_shim()
    if base is None or str(base) == sys.executable:
        return sys.executable, {}
    env: dict[str, str] = {}
    parts = [str(site)] if site is not None else []
    existing = os.environ.get("PYTHONPATH")
    if existing:
        parts.append(existing)
    if parts:
        env["PYTHONPATH"] = os.pathsep.join(parts)
    return str(base), env
