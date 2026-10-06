"""给 Office Web Add-in 准备 HTTPS 证书。

为什么必须 HTTPS
----------------
Office 加载项**即使在开发期也强制 HTTPS**，localhost 并不豁免——这是官方文档的原话，
也是 `office-addin-dev-certs` 这个包存在的唯一理由。所以任务窗格要跑起来，
就得有一张 WebView 认的证书。

分工
----
**生成证书是程序的事；把 CA 装进受信任根是用户的决定。** 所以这里分成两步：

- :func:`generate_certs` —— 纯文件操作，生成一份本地 CA 和一张签给 localhost 的证书；
- :func:`is_ca_trusted` / :func:`trust_hint` —— 只**检查并告诉用户**该跑哪条命令，
  绝不自己去改受信任根存储。

自签的叶子证书直接塞进受信任根也能用，但那等于把"这张证书就是信任锚"和
"它签了谁"混成一件事。这里用两级（CA + 叶子），跟正经开发证书一致，
用户想撤销时删掉 CA 一处即可。
"""

from __future__ import annotations

import datetime
import ipaddress
import json
import subprocess
from pathlib import Path
from typing import Any

from . import __version__
from .config import Settings, ensure_dirs, settings as default_settings

__all__ = [
    "CA_CN",
    "HOSTS",
    "certs_dir",
    "ca_path",
    "leaf_cert_path",
    "leaf_key_path",
    "generate_certs",
    "cert_status",
    "is_ca_trusted",
    "trust_hint",
    "trust_command",
    "manifest_xml",
    "write_manifest",
]

CA_CN = "ppt-agent Dev CA"
LEAF_CN = "localhost"
#: 证书要覆盖的主机名。Office 用 https://localhost，但 127.0.0.1 也一并签上，
#: 免得将来有人把它写成 IP 时又要重签。
HOSTS = ("localhost", "127.0.0.1")
#: 叶子证书有效期。825 天是各种浏览器/WebView 对单张证书的常见上限。
LEAF_DAYS = 825
CA_DAYS = 3650


def _home(cfg: Settings | None) -> Path:
    return (cfg or default_settings).home


def certs_dir(cfg: Settings | None = None) -> Path:
    return _home(cfg) / "certs"


def ca_path(cfg: Settings | None = None) -> Path:
    return certs_dir(cfg) / "ca.pem"


def leaf_cert_path(cfg: Settings | None = None) -> Path:
    return certs_dir(cfg) / "localhost.pem"


def leaf_key_path(cfg: Settings | None = None) -> Path:
    return certs_dir(cfg) / "localhost.key"


def _meta_path(cfg: Settings | None = None) -> Path:
    return certs_dir(cfg) / "meta.json"


# --------------------------------------------------------------------------
# 生成
# --------------------------------------------------------------------------


def generate_certs(cfg: Settings | None = None, *, force: bool = False) -> dict[str, Any]:
    """生成 CA + localhost 证书。已存在且不 ``force`` 时直接复用。

    返回证书路径、指纹与有效期。**不碰受信任根存储**——那是用户的事。
    """
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

    ensure_dirs(cfg)
    directory = certs_dir(cfg)
    directory.mkdir(parents=True, exist_ok=True)

    if leaf_cert_path(cfg).is_file() and ca_path(cfg).is_file() and not force:
        return {"ok": True, "generated": False, **cert_status(cfg)}

    now = datetime.datetime.now(datetime.timezone.utc)

    # ---- CA ----
    ca_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    ca_name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, CA_CN)])
    ca_cert = (
        x509.CertificateBuilder()
        .subject_name(ca_name)
        .issuer_name(ca_name)
        .public_key(ca_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(now + datetime.timedelta(days=CA_DAYS))
        .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                content_commitment=False,
                key_encipherment=False,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=True,
                crl_sign=True,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(x509.SubjectKeyIdentifier.from_public_key(ca_key.public_key()), critical=False)
        .sign(ca_key, hashes.SHA256())
    )

    # ---- 叶子证书（签给 localhost）----
    leaf_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    san = [x509.DNSName("localhost")]
    for host in HOSTS:
        try:
            san.append(x509.IPAddress(ipaddress.ip_address(host)))
        except ValueError:
            if host != "localhost":
                san.append(x509.DNSName(host))

    leaf_cert = (
        x509.CertificateBuilder()
        .subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, LEAF_CN)]))
        .issuer_name(ca_name)
        .public_key(leaf_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(now + datetime.timedelta(days=LEAF_DAYS))
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(x509.SubjectAlternativeName(san), critical=False)
        .add_extension(
            x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH, ExtendedKeyUsageOID.CLIENT_AUTH]),
            critical=False,
        )
        .add_extension(x509.SubjectKeyIdentifier.from_public_key(leaf_key.public_key()), critical=False)
        .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_key.public_key()), critical=False)
        .sign(ca_key, hashes.SHA256())
    )

    # ---- 落盘 ----
    # uvicorn/ssl 要的是"证书链"与"私钥"两个分开的 PEM
    ca_path(cfg).write_bytes(ca_cert.public_bytes(serialization.Encoding.PEM))
    leaf_cert_path(cfg).write_bytes(
        leaf_cert.public_bytes(serialization.Encoding.PEM) + ca_cert.public_bytes(serialization.Encoding.PEM)
    )
    leaf_key_path(cfg).write_bytes(
        leaf_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )
    _meta_path(cfg).write_text(
        json.dumps(
            {
                "ca_thumbprint": ca_cert.fingerprint(hashes.SHA1()).hex().upper(),
                "ca_subject": CA_CN,
                "leaf_subject": LEAF_CN,
                "hosts": list(HOSTS),
                "not_after": leaf_cert.not_valid_after_utc.isoformat(),
                "generated_at": now.isoformat(),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return {"ok": True, "generated": True, **cert_status(cfg)}


# --------------------------------------------------------------------------
# 状态与信任
# --------------------------------------------------------------------------


def _read_meta(cfg: Settings | None = None) -> dict[str, Any]:
    try:
        return json.loads(_meta_path(cfg).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def cert_status(cfg: Settings | None = None) -> dict[str, Any]:
    """证书文件在不在、还有多久过期、CA 有没有被信任。"""
    directory = certs_dir(cfg)
    files = {
        "ca": ca_path(cfg).is_file(),
        "leaf_cert": leaf_cert_path(cfg).is_file(),
        "leaf_key": leaf_key_path(cfg).is_file(),
    }
    ready = all(files.values())
    meta = _read_meta(cfg)
    days_left: int | None = None
    if meta.get("not_after"):
        try:
            expires = datetime.datetime.fromisoformat(meta["not_after"])
            days_left = (expires - datetime.datetime.now(datetime.timezone.utc)).days
        except ValueError:
            days_left = None

    return {
        "dir": str(directory),
        "files": files,
        "ready": ready,
        "trusted": is_ca_trusted(cfg) if ready else False,
        "ca_thumbprint": meta.get("ca_thumbprint"),
        "hosts": meta.get("hosts"),
        "days_left": days_left,
    }


def is_ca_trusted(cfg: Settings | None = None) -> bool:
    """CA 是否已在**当前用户的**受信任根里。

    用 ``certutil -user -store Root <指纹>`` 判定：找得到就退出码 0。
    刻意用当前用户作用域——不需要管理员权限，撤销也只是删一条。
    """
    import os

    if os.name != "nt" or not ca_path(cfg).is_file():
        return False
    meta = _read_meta(cfg)
    thumbprint = str(meta.get("ca_thumbprint") or "").replace(" ", "")
    if not thumbprint:
        return False
    try:
        result = subprocess.run(
            ["certutil", "-user", "-store", "Root", thumbprint],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return result.returncode == 0


def trust_command(cfg: Settings | None = None) -> str:
    """把 CA 装进受信任根的命令——**由用户执行**。"""
    return f'certutil -user -addstore Root "{ca_path(cfg)}"'


def trust_hint(cfg: Settings | None = None) -> str:
    return (
        "Office 加载项强制 HTTPS，而 WebView 只认受信任的证书。所以要把我们生成的这张\n"
        "本地 CA 装进**当前用户**的受信任根存储（不需要管理员，撤销也只是删一条）：\n\n"
        f"  {trust_command(cfg)}\n\n"
        f"撤销：certutil -user -delstore Root {_read_meta(cfg).get('ca_thumbprint') or '<指纹>'}\n"
        "或者：certmgr.msc → 受信任的根证书颁发机构 → 证书 → 删除 “ppt-agent Dev CA”。\n"
        "不装也能用其它所有功能；只有 PowerPoint 任务窗格需要它。"
    )


# --------------------------------------------------------------------------
# manifest
# --------------------------------------------------------------------------


def addin_origin(cfg: Settings | None = None) -> str:
    """任务窗格的源。必须是 localhost（Office 只对 localhost 放宽到本机证书）。"""
    s = cfg or default_settings
    return f"https://localhost:{s.addin_bind_port}"


def manifest_xml(cfg: Settings | None = None, *, addin_id: str | None = None) -> str:
    """生成 Office 加载项清单。

    清单里的地址必须与守护进程实际监听的端口一致，所以**生成**而不是手写死。
    """
    s = cfg or default_settings
    origin = addin_origin(s)
    app_id = addin_id or "7c9b2b1e-6a44-4c2f-9d3a-1f0b6f2c8d51"
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>{app_id}</Id>
  <Version>{__version__}.0</Version>
  <ProviderName>ppt-agent</ProviderName>
  <DefaultLocale>zh-CN</DefaultLocale>
  <DisplayName DefaultValue="ppt-agent 监视台"/>
  <Description DefaultValue="实时显示 agent 对这份演示做了什么：操作流水、当前页画面、跟速与急停。"/>
  <IconUrl DefaultValue="{origin}/addin/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{origin}/addin/icon-80.png"/>
  <SupportUrl DefaultValue="{origin}/addin/"/>
  <AppDomains>
    <AppDomain>{origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Default"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="{origin}/addin/"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="PptAgentTab">
              <Group id="PptAgentGroup">
                <Label resid="GroupLabel"/>
                <Icon>
                  <bt:Image size="16" resid="Icon16"/>
                  <bt:Image size="32" resid="Icon32"/>
                  <bt:Image size="80" resid="Icon80"/>
                </Icon>
                <Control xsi:type="Button" id="PptAgentShow">
                  <Label resid="ShowLabel"/>
                  <Supertip>
                    <Title resid="ShowLabel"/>
                    <Description resid="ShowTip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon16"/>
                    <bt:Image size="32" resid="Icon32"/>
                    <bt:Image size="80" resid="Icon80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>PptAgentPane</TaskpaneId>
                    <SourceLocation resid="PaneUrl"/>
                  </Action>
                </Control>
              </Group>
              <Label resid="TabLabel"/>
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon16" DefaultValue="{origin}/addin/icon-32.png"/>
        <bt:Image id="Icon32" DefaultValue="{origin}/addin/icon-32.png"/>
        <bt:Image id="Icon80" DefaultValue="{origin}/addin/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="PaneUrl" DefaultValue="{origin}/addin/"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="TabLabel" DefaultValue="ppt-agent"/>
        <bt:String id="GroupLabel" DefaultValue="监视台"/>
        <bt:String id="ShowLabel" DefaultValue="打开监视台"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="ShowTip" DefaultValue="实时显示 agent 对 PowerPoint 的每一步操作。"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"""


def write_manifest(out: Path | None = None, cfg: Settings | None = None) -> Path:
    target = out or (Path(__file__).resolve().parent.parent / "addin" / "web" / "manifest.xml")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(manifest_xml(cfg), encoding="utf-8")
    return target
