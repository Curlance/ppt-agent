"""Office Web Add-in：证书链、清单、任务窗格页与端点。

证书是这条路的命门——Office 加载项强制 HTTPS，WebView 只认受信任的证书。
所以这里不测"生成了文件"，而是**真的解析证书、验证签发关系与 SAN**。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from pptd import addin_https, addin_page
from pptd.http_api import create_app

NS = {
    "o": "http://schemas.microsoft.com/office/appforoffice/1.1",
    "ov": "http://schemas.microsoft.com/office/taskpaneappversionoverrides",
    "bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0",
}


# --------------------------------------------------------------------------
# 证书
# --------------------------------------------------------------------------


def test_generate_creates_ca_and_leaf(settings) -> None:
    result = addin_https.generate_certs(settings)
    assert result["ok"] is True and result["generated"] is True
    assert result["ready"] is True
    for path in (addin_https.ca_path(settings), addin_https.leaf_cert_path(settings), addin_https.leaf_key_path(settings)):
        assert path.is_file() and path.stat().st_size > 0


def test_certificate_chain_is_real(settings) -> None:
    """叶子必须真的由这张 CA 签发，且 SAN 覆盖 localhost。"""
    from cryptography import x509

    addin_https.generate_certs(settings)
    ca = x509.load_pem_x509_certificate(addin_https.ca_path(settings).read_bytes())
    leaf = x509.load_pem_x509_certificate(addin_https.leaf_cert_path(settings).read_bytes())

    assert ca.extensions.get_extension_for_class(x509.BasicConstraints).value.ca is True
    assert leaf.extensions.get_extension_for_class(x509.BasicConstraints).value.ca is False
    assert leaf.issuer == ca.subject, "叶子的签发者不是我们生成的 CA"
    # 真正的验证：用 CA 的公钥验签叶子的签名
    ca.public_key().verify(
        leaf.signature,
        leaf.tbs_certificate_bytes,
        __import__("cryptography.hazmat.primitives.asymmetric.padding", fromlist=["PKCS1v15"]).PKCS1v15(),
        leaf.signature_hash_algorithm,
    )

    san = leaf.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
    names = {str(g.value) for g in san}
    assert "localhost" in names
    assert "127.0.0.1" in names, "127.0.0.1 也要签上，否则写成 IP 的客户端会失败"


def test_generate_is_idempotent_without_force(settings) -> None:
    first = addin_https.generate_certs(settings)
    thumbprint = first["ca_thumbprint"]
    second = addin_https.generate_certs(settings)
    assert second["generated"] is False
    assert second["ca_thumbprint"] == thumbprint, "不 force 时不该换证书（换了就得重装信任）"


def test_generate_force_rotates(settings) -> None:
    first = addin_https.generate_certs(settings)
    second = addin_https.generate_certs(settings, force=True)
    assert second["ca_thumbprint"] != first["ca_thumbprint"]


def test_status_reports_days_left(settings) -> None:
    addin_https.generate_certs(settings)
    status = addin_https.cert_status(settings)
    assert status["ready"] is True
    assert status["days_left"] and status["days_left"] > 700
    assert status["hosts"] == list(addin_https.HOSTS)


def test_status_before_generation(settings) -> None:
    status = addin_https.cert_status(settings)
    assert status["ready"] is False and status["trusted"] is False


def test_trust_is_reported_not_assumed(settings) -> None:
    """没装进受信任根时必须如实说没有——装它是用户的决定，不是我们偷偷做。"""
    addin_https.generate_certs(settings)
    assert addin_https.is_ca_trusted(settings) is False
    hint = addin_https.trust_hint(settings)
    assert "certutil" in hint and "addstore" in hint
    assert "受信任的根证书颁发机构" in hint
    assert "delstore" in hint or "删除" in hint, "必须告诉用户怎么撤销"


def test_trust_command_uses_user_scope(settings) -> None:
    """用当前用户作用域：不需要管理员，撤销也只是删一条。"""
    command = addin_https.trust_command(settings)
    assert "-user" in command
    assert str(addin_https.ca_path(settings)) in command


# --------------------------------------------------------------------------
# 清单
# --------------------------------------------------------------------------


def test_manifest_parses_and_has_required_parts(settings) -> None:
    root = ET.fromstring(addin_https.manifest_xml(settings))
    assert root.tag.endswith("OfficeApp")
    assert root.get("{http://www.w3.org/2001/XMLSchema-instance}type") == "TaskPaneApp"
    assert root.find("o:Id", NS) is not None
    assert root.find("o:Hosts/o:Host", NS).get("Name") == "Presentation"
    assert root.find("o:Permissions", NS).text == "ReadWriteDocument"


def test_manifest_urls_are_https_and_match_the_port(settings) -> None:
    root = ET.fromstring(addin_https.manifest_xml(settings))
    origin = addin_https.addin_origin(settings)
    assert origin.startswith("https://localhost:")
    assert str(settings.addin_bind_port) in origin

    source = root.find("o:DefaultSettings/o:SourceLocation", NS).get("DefaultValue")
    assert source == f"{origin}/addin/"
    assert root.find("o:AppDomains/o:AppDomain", NS).text == origin

    # 清单里的**地址值**必须全是 https。
    # （注意：XML 命名空间本身就写着 http://schemas.microsoft.com，那是标识符不是地址，
    #  所以这里检查具体元素，不做全文匹配。另外功能区里的 <bt:Image> 只是**引用**
    #  Resources 里的资源 id、不带地址，真正的地址在 ov:Resources 下。）
    url_values = [
        root.find("o:IconUrl", NS).get("DefaultValue"),
        root.find("o:HighResolutionIconUrl", NS).get("DefaultValue"),
        root.find("o:SupportUrl", NS).get("DefaultValue"),
        source,
        root.find("o:AppDomains/o:AppDomain", NS).text,
    ]
    resources = root.find("ov:VersionOverrides/ov:Resources", NS)
    assert resources is not None, "VersionOverrides 里缺少 Resources（图标与地址的取值处）"
    url_values += [node.get("DefaultValue") for node in resources.findall("bt:Urls/bt:Url", NS)]
    url_values += [node.get("DefaultValue") for node in resources.findall("bt:Images/bt:Image", NS)]

    assert all(v and v.startswith("https://") for v in url_values), url_values


def test_manifest_version_override_gives_a_ribbon_button(settings) -> None:
    root = ET.fromstring(addin_https.manifest_xml(settings))
    override = root.find("ov:VersionOverrides", NS)
    assert override is not None, "没有 VersionOverrides 就没有功能区按钮"

    action = override.find(".//ov:Action", NS)
    assert action is not None
    assert action.get("{http://www.w3.org/2001/XMLSchema-instance}type") == "ShowTaskpane"

    pane_url = override.find(".//bt:Url", NS)
    assert pane_url is not None, "缺少 bt:Url 资源"
    assert pane_url.get("DefaultValue").endswith("/addin/")


def test_manifest_uses_configured_id(settings) -> None:
    fixed = "11111111-2222-3333-4444-555555555555"
    root = ET.fromstring(addin_https.manifest_xml(settings, addin_id=fixed))
    assert root.find("o:Id", NS).text == fixed


# --------------------------------------------------------------------------
# 页面与图标
# --------------------------------------------------------------------------


def test_taskpane_injects_token_and_leaves_no_placeholder() -> None:
    html = addin_page.render_taskpane("SECRET-TOKEN")
    assert "SECRET-TOKEN" in html
    assert "__TOKEN__" not in html
    assert "监视台" in html


def test_taskpane_links_observer_same_origin() -> None:
    """HTTPS 页面跳 HTTP 会被混合内容拦掉，所以观察台要走同源。"""
    html = addin_page.render_taskpane("t")
    assert "OBSERVER = '/?token='" in html
    assert "http://127.0.0.1" not in html


def test_icons_are_valid_png_of_requested_size() -> None:
    import io

    from PIL import Image

    for size in (16, 32, 80):
        data = addin_page.icon_png(size)
        assert data[:4] == b"\x89PNG"
        with Image.open(io.BytesIO(data)) as image:
            assert image.size == (size, size)


# --------------------------------------------------------------------------
# 端点
# --------------------------------------------------------------------------


@pytest.fixture()
def local_client(engine) -> TestClient:
    """Host=localhost 的客户端——模拟 WebView 直接打开任务窗格。"""
    return TestClient(create_app(engine), base_url="http://localhost")


@pytest.fixture()
def remote_client(engine) -> TestClient:
    """Host 不是本机的客户端——模拟端口被隧道到公网之后。"""
    return TestClient(create_app(engine), base_url="http://ppt.example.com")


def test_addin_page_serves_token_to_localhost(local_client: TestClient) -> None:
    resp = local_client.get("/addin/")
    assert resp.status_code == 200
    assert "监视台" in resp.text
    assert "test-token" in resp.text


def test_addin_page_refuses_remote_host_without_token(remote_client: TestClient) -> None:
    """这一条是安全边界：隧道之后 Host 会变成公网域名，绝不能白送令牌。"""
    assert remote_client.get("/addin/").status_code == 401
    ok = remote_client.get("/addin/", headers={"X-PPT-Token": "test-token"})
    assert ok.status_code == 200


def test_addin_manifest_endpoint(local_client: TestClient) -> None:
    resp = local_client.get("/addin/manifest.xml")
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("application/xml")
    assert "<OfficeApp" in resp.text


def test_addin_icon_endpoint(local_client: TestClient) -> None:
    resp = local_client.get("/addin/icon-32.png")
    assert resp.status_code == 200
    assert resp.headers["content-type"] == "image/png"
    assert resp.content[:4] == b"\x89PNG"


def test_addin_icon_rejects_other_sizes(local_client: TestClient) -> None:
    assert local_client.get("/addin/icon-64.png").status_code == 404


def test_addin_endpoints_are_not_in_openapi(local_client: TestClient) -> None:
    """加载项页面是给人/WebView 的，不是给模型的工具，别混进 API 契约。"""
    paths = local_client.get("/openapi.json").json()["paths"]
    assert not any(p.startswith("/addin") for p in paths)
