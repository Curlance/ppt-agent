"""观察台页面：用户看得见 agent 在做什么的那一屏。

刻意做成单文件、零外部依赖：守护进程直接吐这一页，浏览器打开就能看，
不需要装加载项、不需要额外前端工程。

页面读的是同一条 SSE 事件流，所以它和 PowerPoint 任务窗格、DSH 内嵌面板
看到的是同一份事实，不各自造数据。
"""

from __future__ import annotations

__all__ = ["render_observer"]

_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ppt-agent 观察台</title>
<style>
  :root {
    --bg: #0f1115; --panel: #171a21; --line: #262b36; --fg: #e6e9ef; --dim: #8b93a7;
    --accent: #4d6bfe; --ok: #3fb950; --warn: #d29922; --err: #f85149;
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--fg);
         font: 14px/1.6 "Segoe UI", "Microsoft YaHei", system-ui, sans-serif; }
  header { display: flex; align-items: center; gap: 12px; padding: 11px 18px;
           border-bottom: 1px solid var(--line); background: var(--panel); position: sticky; top: 0; z-index: 5; }
  header h1 { font-size: 15px; margin: 0; font-weight: 600; letter-spacing: .3px; white-space: nowrap; }
  .badge { padding: 2px 9px; border-radius: 999px; font-size: 12px;
           border: 1px solid var(--line); color: var(--dim); white-space: nowrap; }
  .badge.on { color: var(--ok); border-color: #1f6f34; }
  .badge.off { color: var(--err); border-color: #7d2b26; }
  .grow { flex: 1; }
  button { background: #232833; color: var(--fg); border: 1px solid var(--line);
           border-radius: 7px; padding: 5px 12px; cursor: pointer; font-size: 13px; white-space: nowrap; }
  button:hover { border-color: var(--accent); }
  button.danger:hover { border-color: var(--err); color: var(--err); }
  .pace { display: flex; align-items: center; gap: 6px; color: var(--dim); font-size: 12px; }
  .pace select { background: #232833; color: var(--fg); border: 1px solid var(--line);
                 border-radius: 7px; padding: 4px 8px; font-size: 13px; }
  .pace input { accent-color: var(--accent); }
  #error { padding: 9px 18px; color: var(--err); display: none; background: #1d1416; }
  main { padding: 14px 18px 22px; display: grid; gap: 14px; }
  section { background: var(--panel); border: 1px solid var(--line); border-radius: 10px; overflow: hidden; }
  section > h2 { margin: 0; padding: 9px 14px; font-size: 13px; font-weight: 600;
                 color: var(--dim); border-bottom: 1px solid var(--line);
                 display: flex; align-items: center; gap: 10px; }
  section > h2 .cap { color: var(--fg); font-weight: 500; }
  .stage { background: #0a0c10; display: flex; align-items: center; justify-content: center;
           min-height: 220px; max-height: 52vh; padding: 10px; }
  .stage img { max-width: 100%; max-height: 48vh; border-radius: 6px; box-shadow: 0 6px 24px #0008; }
  .stage .empty { color: var(--dim); font-size: 13px; padding: 40px 0; }
  .cols { display: grid; grid-template-columns: 1fr 330px; gap: 14px; align-items: start; }
  @media (max-width: 900px) { .cols { grid-template-columns: 1fr; } }
  #feed { list-style: none; margin: 0; padding: 6px 0; max-height: 46vh; overflow: auto; }
  #feed li { display: flex; gap: 10px; padding: 7px 14px; border-bottom: 1px solid #1e222b; }
  #feed li:last-child { border-bottom: 0; }
  #feed .t { color: var(--dim); font-variant-numeric: tabular-nums; font-size: 12px; white-space: nowrap; padding-top: 2px; }
  #feed .m { flex: 1; min-width: 0; }
  #feed .m .sub { color: var(--dim); font-size: 12px; }
  #feed .shots { display: flex; gap: 6px; margin-top: 6px; flex-wrap: wrap; }
  #feed .shots img { width: 132px; border-radius: 4px; border: 1px solid var(--line);
                     cursor: pointer; display: block; }
  #feed .shots img:hover { border-color: var(--accent); }
  .dot { width: 8px; height: 8px; border-radius: 50%; margin-top: 7px; flex: none; background: var(--dim); }
  .dot.run { background: var(--accent); animation: pulse 1s ease-in-out infinite; }
  .dot.ok { background: var(--ok); }
  .dot.err { background: var(--err); }
  .dot.state { background: var(--warn); }
  @keyframes pulse { 50% { opacity: .25; } }
  dl { margin: 0; padding: 10px 14px; display: grid; grid-template-columns: auto 1fr; gap: 5px 12px; font-size: 13px; }
  dt { color: var(--dim); font-size: 12px; white-space: nowrap; }
  dd { margin: 0; word-break: break-all; }
  .tools { margin: 0; padding: 8px 14px 12px; list-style: none; font-size: 13px; }
  .tools li { padding: 5px 0; border-bottom: 1px solid #1e222b; }
  .tools li:last-child { border-bottom: 0; }
  .tools code { color: var(--accent); }
  .tools span { display: block; color: var(--dim); font-size: 12px; }
</style>
</head>
<body>
<header>
  <h1>ppt-agent 观察台</h1>
  <span id="b-com" class="badge">COM ?</span>
  <span id="b-vis" class="badge">可见性 ?</span>
  <span id="b-deck" class="badge">未打开演示</span>
  <span class="grow"></span>
  <label class="pace">跟速
    <select id="pace">
      <option value="0">极速</option>
      <option value="400">正常</option>
      <option value="1500">慢速</option>
    </select>
  </label>
  <label class="pace"><input type="checkbox" id="live" checked> 实时</label>
  <button id="btn-clear">清空视图</button>
  <button id="btn-stop" class="danger">急停（停止守护进程）</button>
</header>
<div id="error"></div>
<main>
  <section>
    <h2>当前画面 <span class="cap" id="stage-cap">还没有截图</span></h2>
    <div class="stage"><div class="empty" id="stage-empty">正在取景…</div><img id="stage-img" alt="" style="display:none"></div>
  </section>
  <div class="cols">
    <section>
      <h2>操作流 — agent 正在做什么</h2>
      <ul id="feed"></ul>
    </section>
    <div style="display:grid; gap:14px">
      <section>
        <h2>状态</h2>
        <dl id="state"><dt>载入中…</dt><dd></dd></dl>
      </section>
      <section>
        <h2>可用工具</h2>
        <ul class="tools" id="tools"></ul>
      </section>
    </div>
  </div>
</main>
<script>
const TOKEN = "__TOKEN__";
const feed = document.getElementById('feed');
const stateEl = document.getElementById('state');
const toolsEl = document.getElementById('tools');
const errorEl = document.getElementById('error');
const stageImg = document.getElementById('stage-img');
const stageEmpty = document.getElementById('stage-empty');
const stageCap = document.getElementById('stage-cap');
const liveBox = document.getElementById('live');
const rows = new Map();
let liveTimer = null;
let lastEventAt = Date.now();

function showError(msg) {
  errorEl.textContent = msg;
  errorEl.style.display = msg ? 'block' : 'none';
}

function el(tag, cls, text) {
  const node = document.createElement(tag);
  if (cls) node.className = cls;
  if (text !== undefined) node.textContent = text;
  return node;
}

function clock(iso) {
  try { return new Date(iso).toLocaleTimeString('zh-CN', { hour12: false }); }
  catch (e) { return ''; }
}

function withToken(url) {
  return url + (url.includes('?') ? '&' : '?') + 'token=' + encodeURIComponent(TOKEN);
}

function setStage(shot, fromClick) {
  // 用户主动点缩略图 = 想细看这一帧高清图，于是暂停实时，免得被下一拍覆盖
  if (fromClick) {
    liveBox.checked = false;
    scheduleLive();
  }
  const image = shot.image || shot.thumb || {};
  if (!image.url) return;
  clearStageBlob();
  stageImg.src = withToken(image.url);
  stageImg.style.display = 'block';
  stageEmpty.style.display = 'none';
  const title = shot.title ? '· ' + shot.title : '';
  stageCap.textContent = '快照 · 第 ' + shot.slide + ' 页 ' + title;
}

function clearStageBlob() {
  if (stageImg.dataset.blob) {
    URL.revokeObjectURL(stageImg.dataset.blob);
    delete stageImg.dataset.blob;
  }
}

async function refreshLive() {
  if (!liveBox.checked || document.hidden) return;
  try {
    const resp = await fetch('/live?width=960&token=' + encodeURIComponent(TOKEN) + '&t=' + Date.now());
    if (resp.status === 204) return;          // 上一帧还没拍完，跳过这一拍
    if (resp.status === 409) {
      // 演示没打开或没有幻灯片——把原因说清楚，别让大屏一直空着让人猜
      if (!stageImg.dataset.blob) {
        stageEmpty.textContent = '当前没有打开的演示（先用 ppt_open 打开一个）';
        stageEmpty.style.display = 'block';
        stageCap.textContent = '取不到画面';
      }
      return;
    }
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const slide = resp.headers.get('X-PPT-Slide') || '';
    const blob = await resp.blob();
    clearStageBlob();
    const url = URL.createObjectURL(blob);
    stageImg.dataset.blob = url;
    stageImg.src = url;
    stageImg.style.display = 'block';
    stageEmpty.style.display = 'none';
    stageCap.textContent = '实时 · 第 ' + slide + ' 页';
    showError('');
  } catch (e) {
    showError('实时取景失败：' + e.message);
    if (!stageImg.dataset.blob) stageEmpty.textContent = '取景失败：' + e.message;
  }
}

function scheduleLive() {
  clearTimeout(liveTimer);
  if (!liveBox.checked || document.hidden) return;
  // 有动静时勤一点（2 秒），闲着时省一点（8 秒），标签页不可见就完全停
  const idle = Date.now() - lastEventAt > 15000;
  liveTimer = setTimeout(async () => { await refreshLive(); scheduleLive(); }, idle ? 8000 : 2000);
}

function addRow(id, dotClass, text, sub) {
  const li = el('li');
  li.appendChild(el('span', 'dot ' + dotClass));
  li.appendChild(el('span', 't', clock(new Date().toISOString())));
  const m = el('div', 'm');
  m.appendChild(el('div', null, text));
  if (sub) m.appendChild(el('div', 'sub', sub));
  li.appendChild(m);
  feed.appendChild(li);
  if (id) rows.set(id, li);
  while (feed.children.length > 300) feed.removeChild(feed.firstChild);
  feed.scrollTop = feed.scrollHeight;
  return li;
}

function rowFor(d) {
  let li = rows.get(d.id);
  if (!li) li = addRow(d.id, 'run', d.title || d.tool || '操作', '');
  return li;
}

function attachShots(d) {
  const li = rowFor(d);
  const box = el('div', 'shots');
  for (const shot of d.shots || []) {
    const thumb = shot.thumb || shot.image;
    if (!thumb || !thumb.url) continue;
    const img = el('img');
    img.src = withToken(thumb.url);
    img.alt = '第 ' + shot.slide + ' 页';
    img.title = '点击看这一帧的高清图（会暂停实时）';
    img.onclick = () => setStage(shot, true);
    box.appendChild(img);
  }
  if (box.children.length) li.querySelector('.m').appendChild(box);
  // 实时开着时大屏归实时管，别让快照把它顶掉；关着时才自动上最新快照
  if (!liveBox.checked) {
    const latest = (d.shots || [])[0];
    if (latest) setStage(latest, false);
  }
}

function applyState(s) {
  const com = s.com || {};
  const bCom = document.getElementById('b-com');
  bCom.textContent = 'COM ' + (com.alive ? (com.version || '已连接') : '未连接');
  bCom.className = 'badge ' + (com.alive ? 'on' : 'off');
  const bVis = document.getElementById('b-vis');
  bVis.textContent = '可见性 ' + (com.visible ? '窗口可见' : '窗口不可见');
  bVis.className = 'badge ' + (com.visible ? 'on' : 'off');
  const deck = s.active_deck || '';
  const bDeck = document.getElementById('b-deck');
  // 徽章只放文件名：一整条路径会把标题栏撑爆（截图里看到的）
  bDeck.textContent = deck ? String(deck).split(/[\\\\/]/).pop() : '未打开演示';
  bDeck.className = 'badge ' + (deck ? 'on' : '');

  const rowsData = [
    ['PowerPoint', com.version || '—'],
    ['附着方式', com.attached === null || com.attached === undefined ? '—' : (com.attached ? '接管已运行的实例' : '由 pptd 启动')],
    ['窗口可见', com.visible ? '是' : '否'],
    ['跟速', (s.pace_ms ?? 0) + ' ms'],
    ['已调用次数', com.calls ?? 0],
    ['超时次数', com.timeouts ?? 0],
    ['降级', com.degraded ? '是（建议急停后重启）' : '否'],
    ['当前演示', deck ? String(deck).split(/[\\\\/]/).pop() : '—'],
    ['启动时间', s.started_at || '—'],
    ['最近错误', com.last_error || '无'],
  ];
  stateEl.innerHTML = '';
  for (const [k, v] of rowsData) {
    stateEl.appendChild(el('dt', null, k));
    stateEl.appendChild(el('dd', null, String(v)));
  }
  syncPace(s.pace_ms);
}

function handle(ev) {
  lastEventAt = Date.now();
  let e;
  try { e = JSON.parse(ev.data); } catch (err) { return; }
  const d = e.data || {};
  if (e.type === 'op.start') {
    addRow(d.id, 'run', d.summary || d.title || d.tool, '开始 · ' + d.tool + ' #' + d.id);
  } else if (e.type === 'op.end') {
    const li = rowFor(d);
    li.querySelector('.dot').className = 'dot ' + (d.ok ? 'ok' : 'err');
    li.querySelector('.m').firstChild.textContent = d.human || (d.ok ? '完成' : '失败');
    li.querySelector('.m').appendChild(el('div', 'sub',
      (d.ok ? '完成' : '失败') + ' · ' + (d.ms ?? '?') + ' ms #' + d.id));
  } else if (e.type === 'op.shot') {
    attachShots(d);
  } else if (e.type === 'state') {
    window.__state = Object.assign({}, window.__state || {}, d);
    applyState(window.__state);
  } else if (e.type === 'log') {
    // 只在警告/错误时显示级别——每条都挂个 "info" 只是噪音（截图里看到的）
    const level = String(d.level || '').toLowerCase();
    const noisy = level && level !== 'info';
    // 注意 addRow 的签名是 (id, dotClass, text, sub)，别把 message 传到 dotClass 上
    addRow(null, noisy ? 'err' : '', d.message || '日志', noisy ? level : '');
  }
}

async function refreshState() {
  try {
    const r = await fetch('/state', { headers: { 'X-PPT-Token': TOKEN } });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const s = await r.json();
    window.__state = Object.assign({}, window.__state || {}, s);
    applyState(window.__state);
    showError('');
  } catch (e) { showError('无法读取守护进程状态：' + e.message); }
}

async function loadTools() {
  try {
    // 必须带令牌：/tools 不在本机免令牌的白名单里。
    // （这个 bug 是"真的截图看一眼"才发现的——只读页面的工具列表一直是空的，
    //  而 HTML 合法、JS 语法正确、接口也通，前三项加起来仍然不等于页面对。）
    const r = await fetch('/tools', { headers: { 'X-PPT-Token': TOKEN } });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const data = await r.json();
    toolsEl.innerHTML = '';
    for (const t of data.tools) {
      const li = el('li');
      li.appendChild(el('code', null, t.name));
      li.appendChild(el('span', null, t.summary));
      toolsEl.appendChild(li);
    }
  } catch (e) {
    toolsEl.innerHTML = '';
    toolsEl.appendChild(el('li', null, '读取工具列表失败：' + e.message));
  }
}

document.getElementById('btn-clear').onclick = () => { feed.innerHTML = ''; rows.clear(); };

const paceSel = document.getElementById('pace');
paceSel.onchange = async () => {
  try {
    await fetch('/pace', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-PPT-Token': TOKEN },
      body: JSON.stringify({ pace_ms: Number(paceSel.value) })
    });
  } catch (e) { showError('设置跟速失败：' + e.message); }
};

function syncPace(ms) {
  if (ms === null || ms === undefined) return;
  const value = String(ms);
  if ([...paceSel.options].some(o => o.value === value)) paceSel.value = value;
}

document.getElementById('btn-stop').onclick = async () => {
  if (!confirm('确定要停止守护进程吗？PowerPoint 不会被关闭。')) return;
  await fetch('/shutdown', { method: 'POST', headers: { 'X-PPT-Token': TOKEN } });
  showError('守护进程已请求停止。');
};

const es = new EventSource('/events?token=' + encodeURIComponent(TOKEN));
for (const type of ['op.start', 'op.end', 'op.shot', 'state', 'log']) {
  es.addEventListener(type, handle);
}
es.onerror = () => showError('与守护进程的事件流断开，正在重连…');
es.onopen = () => showError('');

liveBox.onchange = () => {
  if (liveBox.checked) refreshLive();
  scheduleLive();
};
document.addEventListener('visibilitychange', scheduleLive);

refreshState();
loadTools();
setInterval(refreshState, 5000);
refreshLive();
scheduleLive();
</script>
</body>
</html>
"""


def render_observer(token: str, host: str, port: int) -> str:
    """渲染观察台页面。令牌直接注入——只有回环客户端能取到这一页。"""
    return (
        _TEMPLATE.replace("__TOKEN__", token)
        .replace("__HOST__", host)
        .replace("__PORT__", str(port))
    )
