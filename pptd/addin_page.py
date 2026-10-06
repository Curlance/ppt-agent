"""Office Web Add-in 的任務窗格页面与图标。

任务窗格只有 300~400 px 宽，所以版面与观察台不同：状态一行、画面一块、流水一列、
按钮一排。但它读的是**同一份数据**——`/feed.txt` 与 `/live`，跟 VBA 面板、观察台完全一致，
不另造一套。

为什么用轮询而不是 SSE：任务窗格跑在 WebView 里，长连接在切换页面/休眠时容易断，
而面板要显示的只是"最近发生了什么"，两秒一次轮询更稳、更好排查。
"""

from __future__ import annotations

from typing import Any

__all__ = ["render_taskpane", "icon_png"]

_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ppt-agent 监视台</title>
<style>
  :root { --bg:#171a21; --line:#262b36; --fg:#e6e9ef; --dim:#8b93a7;
          --accent:#4d6bfe; --ok:#3fb950; --err:#f85149; --warn:#d29922; }
  * { box-sizing: border-box; }
  html, body { margin:0; height:100%; background:var(--bg); color:var(--fg);
               font:13px/1.55 "Segoe UI","Microsoft YaHei",system-ui,sans-serif; }
  body { display:flex; flex-direction:column; }
  header { padding:8px 10px 6px; border-bottom:1px solid var(--line); }
  #status { font-size:13px; }
  #sub { color:var(--dim); font-size:12px; margin-top:2px; word-break:break-all; }
  .dot { display:inline-block; width:7px; height:7px; border-radius:50%;
         background:var(--dim); margin-right:6px; vertical-align:middle; }
  .dot.on { background:var(--ok); } .dot.off { background:var(--err); }
  #stage { background:#0a0c10; border-bottom:1px solid var(--line); text-align:center; padding:6px; }
  #stage img { max-width:100%; border-radius:4px; display:none; }
  #stage .empty { color:var(--dim); font-size:12px; padding:22px 0; }
  #feedWrap { flex:1; overflow:auto; min-height:80px; }
  #feed { list-style:none; margin:0; padding:4px 0; }
  #feed li { padding:5px 10px; border-bottom:1px solid #1e222b; }
  #feed li:last-child { border-bottom:0; }
  #feed .t { color:var(--dim); font-size:11px; margin-right:6px; font-variant-numeric:tabular-nums; }
  #feed .err { color:var(--err); }
  footer { border-top:1px solid var(--line); padding:8px 10px; display:flex; gap:6px;
           align-items:center; flex-wrap:wrap;
           /* 关键：不许被压缩。第一版没写这句，窄窗格下按钮换行到第二行，
              而第二行落到视口之外——表现成"急停按钮不见了"。 */
           flex:0 0 auto; }
  select, button { background:#232833; color:var(--fg); border:1px solid var(--line);
                   border-radius:6px; padding:4px 9px; font-size:12px; cursor:pointer;
                   flex:0 0 auto; white-space:nowrap; }
  button:hover { border-color:var(--accent); }
  button.danger:hover { border-color:var(--err); color:var(--err); }
  #btnObserver { margin-left:auto; }
  #feed li { overflow-wrap:anywhere; word-break:break-word; }
  #sub { overflow-wrap:anywhere; }
  #error { color:var(--err); padding:6px 10px; font-size:12px; display:none; }
</style>
</head>
<body>
<header>
  <div id="status"><span class="dot" id="dot"></span><span id="statusText">正在连接守护进程…</span></div>
  <div id="sub"></div>
</header>
<div id="error"></div>
<section id="stage"><div class="empty" id="stageEmpty">正在取景…</div><img id="stageImg" alt=""></section>
<div id="feedWrap"><ul id="feed"></ul></div>
<footer>
  <label style="color:var(--dim);font-size:12px">跟速</label>
  <select id="pace">
    <option value="0">极速</option>
    <option value="400">正常</option>
    <option value="1500">慢速</option>
  </select>
  <button id="btnObserver">观察台</button>
  <button id="btnStop" class="danger">急停</button>
</footer>
<script id="paneScript">
const TOKEN = "__TOKEN__";
// 刻意用**同源**地址：任务窗格是 HTTPS，而观察台默认跑在 HTTP 端口上，
// 从 HTTPS 页面跳 HTTP 可能被混合内容规则拦掉。这个 HTTPS 端口服务的是同一个应用，
// 所以观察台在这里也拿得到，直接同源打开最稳。
const OBSERVER = '/?token=' + encodeURIComponent(TOKEN);

const dot = document.getElementById('dot');
const statusText = document.getElementById('statusText');
const subEl = document.getElementById('sub');
const feedEl = document.getElementById('feed');
const errorEl = document.getElementById('error');
const stageImg = document.getElementById('stageImg');
const stageEmpty = document.getElementById('stageEmpty');
const paceSel = document.getElementById('pace');

let lastEventAt = Date.now();
let busy = false;

function showError(msg) {
  errorEl.textContent = msg || '';
  errorEl.style.display = msg ? 'block' : 'none';
}

function authHeaders(extra) {
  return Object.assign({ 'X-PPT-Token': TOKEN }, extra || {});
}

function parseFeed(text) {
  const head = {};
  const rows = [];
  let inRows = false;
  for (const raw of text.split('\\n')) {
    const line = raw.replace(/\\r/g, '').trim();
    if (line === '---') { inRows = true; continue; }
    if (!line) continue;
    if (!inRows) {
      const i = line.indexOf('=');
      if (i > 0) head[line.slice(0, i)] = line.slice(i + 1);
    } else {
      rows.push(line);
    }
  }
  return { head, rows };
}

async function refreshFeed() {
  if (busy) return;
  busy = true;
  try {
    const resp = await fetch('/feed.txt?limit=40', { headers: authHeaders(), cache: 'no-store' });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const { head, rows } = parseFeed(await resp.text());

    const alive = head.alive === 'true';
    dot.className = 'dot ' + (alive ? 'on' : 'off');
    if (!alive) {
      statusText.textContent = '守护进程未连接';
      subEl.textContent = '先运行：pptctl serve';
    } else {
      statusText.textContent = 'PowerPoint ' + (head.version || '?')
        + (head.visible === 'true' ? ' · 窗口可见' : ' · 窗口不可见');
      const bits = [];
      if (head.deck) bits.push(head.deck);
      if (head.slide) bits.push('第 ' + head.slide + ' 页');
      if (head.pace_ms) bits.push('跟速 ' + head.pace_ms + 'ms');
      subEl.textContent = bits.join(' · ');
      if ([...paceSel.options].some(o => o.value === String(head.pace_ms))) paceSel.value = String(head.pace_ms);
    }

    feedEl.innerHTML = '';
    for (const row of rows.slice(-40)) {
      const parts = row.split('|').map(s => s.trim());
      const li = document.createElement('li');
      const time = document.createElement('span');
      time.className = 't';
      time.textContent = (parts[0] || '').slice(11, 19);
      li.appendChild(time);
      const body = document.createElement('span');
      const text = parts.slice(2).join(' | ') || parts[1] || '';
      body.textContent = text;
      if (text.startsWith('ERR')) body.className = 'err';
      li.appendChild(body);
      feedEl.appendChild(li);
    }
    feedEl.parentElement.scrollTop = feedEl.parentElement.scrollHeight;
    lastEventAt = Date.now();
    showError('');
  } catch (e) {
    showError('读取面板数据失败：' + e.message);
  } finally {
    busy = false;
  }
}

async function refreshLive() {
  // 上一帧还没拍完（204）静默跳过；没有演示（409）要说清楚原因，
  // 别让画面区一直空着让人猜（第一版就是这么含混过去的）。
  try {
    const resp = await fetch('/live?width=420&t=' + Date.now(), { headers: authHeaders(), cache: 'no-store' });
    if (resp.status === 204) return;
    if (resp.status === 409) {
      if (!stageImg.dataset.blob) stageEmpty.textContent = '当前没有打开的演示';
      return;
    }
    if (!resp.ok) return;
    const url = URL.createObjectURL(await resp.blob());
    if (stageImg.dataset.blob) URL.revokeObjectURL(stageImg.dataset.blob);
    stageImg.dataset.blob = url;
    stageImg.src = url;
    stageImg.style.display = 'block';
    stageEmpty.style.display = 'none';
  } catch (e) { /* 取景失败不影响面板其余部分 */ }
}

paceSel.onchange = async () => {
  try {
    await fetch('/pace', {
      method: 'POST',
      headers: authHeaders({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ pace_ms: Number(paceSel.value) })
    });
    refreshFeed();
  } catch (e) { showError('设置跟速失败：' + e.message); }
};

document.getElementById('btnObserver').onclick = () => window.open(OBSERVER, '_blank');

document.getElementById('btnStop').onclick = async () => {
  if (!confirm('确定要停止 ppt-agent 守护进程吗？\\nPowerPoint 不会被关闭，已打开的文件不受影响。')) return;
  try {
    await fetch('/shutdown', { method: 'POST', headers: authHeaders() });
    showError('已请求停止守护进程（PowerPoint 未受影响）');
  } catch (e) { showError('发送失败：' + e.message); }
};

refreshFeed();
refreshLive();
setInterval(refreshFeed, 2500);
setInterval(refreshLive, 4000);
document.addEventListener('visibilitychange', () => {
  if (!document.hidden) { refreshFeed(); refreshLive(); }
});
</script>
</body>
</html>
"""


def render_taskpane(token: str, host: str = "", port: int = 0) -> str:
    """渲染任务窗格页面。

    令牌由服务端注入——页面本身不携带任何凭据。``host``/``port`` 只为兼容调用方签名，
    页面里的链接一律走同源，不写死外部地址。
    """
    return _TEMPLATE.replace("__TOKEN__", token)


def icon_png(size: int) -> bytes:
    """生成加载项图标。

    刻意不写字：字体渲染在不同机器上不一致，几何图形才是可复现的。
    两个错位的圆角矩形 = 两张叠起来的幻灯片。
    """
    import io

    from PIL import Image, ImageDraw

    scale = 4  # 超采样后再缩，边缘才干净
    canvas = Image.new("RGBA", (size * scale, size * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(canvas)

    unit = size * scale
    radius = max(2, unit // 8)
    back = (int(unit * 0.10), int(unit * 0.10), int(unit * 0.62), int(unit * 0.76))
    front = (int(unit * 0.30), int(unit * 0.26), int(unit * 0.92), int(unit * 0.92))

    draw.rounded_rectangle(back, radius=radius, fill=(77, 107, 254, 255))          # #4D6BFE
    draw.rounded_rectangle(front, radius=radius, fill=(107, 132, 255, 255))        # #6B84FF
    draw.rounded_rectangle(front, radius=radius, outline=(255, 255, 255, 220), width=max(1, unit // 32))

    out = canvas.resize((size, size), Image.LANCZOS)
    buffer = io.BytesIO()
    out.save(buffer, "PNG")
    return buffer.getvalue()
