/**
 * 用桩模拟 DSH 的浏览器模块表，检查 client.js 是否符合插件契约。
 *
 * 为什么值得写这个：界面我无法在真机上验证（DSH 要重启才会加载），但**契约**可以验：
 * 注册 id 是否等于包名、是否注入 slots、是否注册到那个 slot、工厂有没有副作用、
 * 组件卸载后定时器有没有被清掉。这些恰恰是最容易写错、也最容易被宿主拒绝的地方。
 *
 * 用法：node dsh_client_probe.mjs <client.js 路径>
 * 输出：一行 JSON，交给 pytest 断言。
 */

import { readFileSync } from 'node:fs';

const target = process.argv[2];
const result = { ok: false, checks: [], errors: [] };

function check(name, condition, detail) {
  result.checks.push({ name, pass: Boolean(condition), detail: detail === undefined ? null : String(detail) });
  return Boolean(condition);
}

// ---- 桩：DSH 的浏览器模块表 ----
const registrations = [];
const windowStub = {
  __ModuleLoader__: {
    load(spec) {
      registrations.push(spec);
    },
  },
  addEventListener() {},
};

// ---- 桩：React ----
// useState 的 setter 要真的记下来：否则组件里"连不上时改成提示文案"这段逻辑
// 在探针里根本观察不到（第一版就是这么误报的）。
const rendered = [];
const effects = [];
const stateUpdates = [];
const reactStub = {
  createElement(type, props, ...children) {
    const element = { __element: true, type, props, children };
    rendered.push(element);
    return element;
  },
  useState(initial) {
    return [initial, (value) => stateUpdates.push(value)];
  },
  useEffect(fn) {
    effects.push(fn);
  },
};

const required = [];
const requireStub = (name) => {
  required.push(name);
  if (name === 'react') return reactStub;
  throw new Error('未声明的外部依赖：' + name);
};

// ---- 桩：定时器与 fetch ----
const timers = [];
const originalSetInterval = globalThis.setInterval;
const originalClearInterval = globalThis.clearInterval;
globalThis.setInterval = (fn, ms) => {
  timers.push({ fn, ms, cleared: false });
  return timers.length;
};
globalThis.clearInterval = (id) => {
  if (timers[id - 1]) timers[id - 1].cleared = true;
};

let fetchCalls = 0;
let fetchShouldFail = false;
let fetchBody = '● PowerPoint 16.0 · source.pptx · 第 2 页 · 已改写「标题 1」的文字';
globalThis.fetch = async () => {
  fetchCalls += 1;
  if (fetchShouldFail) throw new Error('连接被拒绝');
  return { ok: true, status: 200, text: async () => fetchBody };
};

// ---- 加载被测文件 ----
let code;
try {
  code = readFileSync(target, 'utf8');
} catch (err) {
  result.errors.push('读不到 client.js：' + err.message);
  console.log(JSON.stringify(result));
  process.exit(0);
}

try {
  new Function('window', code)(windowStub);
} catch (err) {
  result.errors.push('执行 client.js 抛异常：' + err.message);
  console.log(JSON.stringify(result));
  process.exit(0);
}

check('只注册了一个模块', registrations.length === 1, `实际 ${registrations.length}`);

const spec = registrations[0] || {};
check('注册 id 等于包名', spec.id === '@local/ppt-agent-mcp', spec.id);
check('factory 是函数', typeof spec.factory === 'function');

// ---- 工厂必须无副作用（DSH 明文要求）----
const renderedBefore = rendered.length;
const requiredBefore = required.length;
let module_ = null;
try {
  module_ = spec.factory(requireStub);
} catch (err) {
  result.errors.push('factory 抛异常：' + err.message);
}
check('factory 调用本身不渲染任何东西', rendered.length === renderedBefore, `${rendered.length - renderedBefore} 个元素`);
check(
  'factory 只声明了 react 这一个依赖',
  required.slice(requiredBefore).every((n) => n === 'react'),
  required.slice(requiredBefore).join(',')
);

// ---- 插件对象契约 ----
check('声明注入 slots', Array.isArray(module_?.inject) && module_.inject.includes('slots'), JSON.stringify(module_?.inject));
check('apply 是函数', typeof module_?.apply === 'function');

// ---- 注册到哪个 slot ----
const injected = [];
const registered = [];
const ctxStub = {
  slots: {
    inject(name, factory) {
      injected.push(name);
      factory();
      return () => {};
    },
    register(options, component) {
      registered.push({ options, component });
      return () => {};
    },
  },
};

try {
  module_.apply(ctxStub);
} catch (err) {
  result.errors.push('apply 抛异常：' + err.message);
}

check(
  '注入的 slot 是 conversation.composer.dock',
  injected.length === 1 && injected[0] === 'conversation.composer.dock',
  injected.join(',')
);
check('注册了一个组件', registered.length === 1, `实际 ${registered.length}`);

const entry = registered[0] || { options: {}, component: null };
check('注册项带 name', entry.options.name === 'conversation.composer.dock', entry.options.name);
check('注册项带稳定 id', entry.options.id === 'ppt-agent', entry.options.id);
check('注册项带 order（不与宿主项抢位）', typeof entry.options.order === 'number', entry.options.order);
check('组件是函数', typeof entry.component === 'function');

// ---- 组件能渲染，且换行被收起（窄条不能被撑破）----
let element = null;
try {
  element = entry.component();
} catch (err) {
  result.errors.push('组件渲染抛异常：' + err.message);
}
if (element) {
  check('渲染出的是一个链接', element.type === 'a', element.type);
  check('链接指向本地观察台', String(element.props?.href || '').includes('127.0.0.1'), element.props?.href);
  check('在新标签页打开', element.props?.target === '_blank', element.props?.target);
  const style = element.props?.style || {};
  check('窄条不换行', style.whiteSpace === 'nowrap', style.whiteSpace);
  check('超长省略而不是撑破', style.textOverflow === 'ellipsis', style.textOverflow);
  check('继承宿主主题色', style.color === 'inherit' || String(style.color).includes('dsw-'), style.color);
  check('文字是状态行', typeof element.children?.[0] === 'string', element.children?.[0]);
}

// ---- 副作用：轮询与卸载清理 ----
check('组件注册了一个 effect', effects.length >= 1, `实际 ${effects.length}`);

let cleanup = null;
if (effects.length) {
  try {
    cleanup = effects[0]();
  } catch (err) {
    result.errors.push('effect 抛异常：' + err.message);
  }
}
check('起了定时器做轮询', timers.length === 1, `实际 ${timers.length}`);
check('轮询周期合理（2~10 秒）', timers[0] && timers[0].ms >= 2000 && timers[0].ms <= 10000, timers[0]?.ms);
check('effect 返回了清理函数', typeof cleanup === 'function');

await new Promise((r) => setTimeout(r, 30));
check('真的去取了状态', fetchCalls >= 1, `fetch 调用 ${fetchCalls} 次`);
check(
  '拿到状态后更新了显示文本',
  stateUpdates.some((v) => typeof v === 'string' && v.includes('PowerPoint')),
  JSON.stringify(stateUpdates.slice(0, 3))
);

if (typeof cleanup === 'function') {
  cleanup();
  check('卸载时清掉了定时器', timers[0]?.cleared === true);
}

// ---- 守护进程不在时不能把宿主界面弄崩，而且要给出可执行的提示 ----
fetchShouldFail = true;
const updatesBefore = stateUpdates.length;
let failureElement = null;
try {
  failureElement = entry.component();
  if (effects.length > 1) effects[effects.length - 1]();
  await new Promise((r) => setTimeout(r, 30));
  check('连不上时仍然渲染出内容', Boolean(failureElement));
  const failureUpdates = stateUpdates.slice(updatesBefore).map(String);
  check(
    '连不上时给出可执行的提示',
    failureUpdates.some((v) => v.includes('pptctl serve')),
    JSON.stringify(failureUpdates)
  );
  check(
    '连不上时标记为断开（显示会变淡）',
    stateUpdates.slice(updatesBefore).includes(true),
    JSON.stringify(stateUpdates.slice(updatesBefore))
  );
} catch (err) {
  result.errors.push('守护进程不可用时组件抛异常：' + err.message);
}

globalThis.setInterval = originalSetInterval;
globalThis.clearInterval = originalClearInterval;

result.ok = result.checks.every((c) => c.pass) && result.errors.length === 0;
console.log(JSON.stringify(result));
