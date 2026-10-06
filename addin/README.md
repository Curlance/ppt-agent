# PowerPoint 加载项

这个目录下有**两条独立的路**，给 PowerPoint 里加一块由 agent 驱动的面板。
两条实现独立；需要分别完成安装与真机验收。当前已验证浏览器中的页面渲染，
尚未完成 PowerPoint 内加载后的端到端验收，详见根目录的交付报告。

| | **① `.ppam` VBA 监视台** | **② Web Add-in 任务窗格** |
|---|---|---|
| 形态 | 无模式**浮窗** | **停靠式任务窗格**（功能区按钮打开） |
| 代价 | 导入一次 VBA 源码（或开一次 VBA 工程信任） | 要往**受信任根**装一张本地 CA + 旁加载清单 |
| 依赖 | 无 | HTTPS + 证书（Office 加载项强制 HTTPS，localhost 不豁免） |
| 文件 | [`vba/`](vba/) | [`web/manifest.xml`](web/manifest.xml) + 服务端页面 |
| 数据源 | `/feed.txt` | `/feed.txt` + `/live`（**和 VBA 面板同一份**） |

---

# ① `.ppam` VBA 监视台

实时显示守护进程状态与 agent 的操作流水，带「极速 / 正常 / 慢速 / 急停」按钮。

**最快的装法**（需要先关掉 PowerPoint 开关一次信任，见下一节）：

```powershell
pptctl addin-trust-vba --enable     # 关掉 PowerPoint 后执行
pptctl addin-build                  # 自动编译出 .ppam
```

```
┌─ ppt-agent 监视台 ───────────────────────────────┐
│ ● PowerPoint 16.0 · 窗口可见 · demo.pptx · 第 2 页 · 跟速 400ms · 24 个工具 │
├──────────────────────────────────────────────────┤
│ 10:44:12 | op.start | 把某一页导出成图片：给人看高清 PNG，给模型看…          │
│ 10:44:15 | op.end   | OK 已截取第 2 页                                        │
│ 10:44:20 | op.start | 改写形状里的文字                                        │
│ 10:44:21 | op.end   | OK 已改写「标题 1」的文字（8 字）                       │
├──────────────────────────────────────────────────┤
│ ☑ 自动刷新   [极速] [正常] [慢速] [刷新] [急停]   │
└──────────────────────────────────────────────────┘
```

## 先说清楚三件事

1. **这不是停靠式任务窗格。** VBA 没有 `CustomTaskPane` API，所以 Windows 版 PowerPoint
   里 VBA 加载项只能给**无模式浮窗**。真正的停靠窗格要用 Office Web Add-in 或 COM 加载项，
   而 Office 加载项**即使在开发期也强制 HTTPS**（localhost 不豁免），
   意味着要往你的受信任根存储里装一张开发证书。这里选了不碰证书的那条路。
2. **面板不做 COM 自动化。** 它只读 `http://127.0.0.1:8791/feed.txt`、只发 `/pace` 与 `/shutdown`。
   真正操作 PowerPoint 的是 pptd 守护进程——两个进程同时驱动同一个 PowerPoint 会互相
   把对方的调用打成 `RPC_E_CALL_REJECTED`。
3. **它读的是纯文本而不是 JSON。** VBA 里没有 JSON 解析器，手写一个又长又脆；
   `/feed.txt` 一行一条、` | ` 分隔，`Split` 一句就够用。

## 方式一：自动构建（推荐，需要一次性开启一个开发开关）

```powershell
pptctl addin-build
```

前提是开启「信任对 VBA 工程对象模型的访问」：

> PowerPoint → 文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置
> → 勾选「**信任对 VBA 工程对象模型的访问**」→ 确定 → 重启 PowerPoint

这是微软为 VBA 开发提供的标准开关，**随时可以取消勾选**。开启后 `pptctl addin-build`
会用 PowerPoint 自己把源码编译成 `addin/ppt-agent.ppam`。

## 方式二：手动导入（不需要开那个开关）

```powershell
pptctl addin-export          # 导出成 VBE 能读的 GBK 编码
```

然后在 PowerPoint 里：

1. **先新建一个空白演示**（VBA 导入会进到"当前演示"的工程里，别污染你正在做的文件）
2. 按 `Alt+F11` 打开 VBA 编辑器
3. 菜单 **文件 → 导入文件** → 选 `addin/export/PPTAgent.bas`
4. 菜单 **文件 → 导入文件** → 选 `addin/export/PPTAgentPanel.frm`
5. 回到 PowerPoint → **文件 → 另存为** → 类型选 **PowerPoint 加载项 (\*.ppam)**
   → 存到 `%APPDATA%\Microsoft\AddIns\ppt-agent.ppam`
6. 关掉那份临时空白演示（选"不保存"）
7. **文件 → 选项 → 加载项** → 底下"管理"选 **PowerPoint 加载项** → **转到** → 勾选 **ppt-agent**
8. 再用 **文件 → 选项 → 快速访问工具栏**，把宏 `PPTAgentShow` 加进去，以后一键打开面板

> 第 3、4 步的顺序无所谓，但**必须先新建空白演示**——否则 `.bas` 会被导进你正在编辑的
> 那份演示里。

## 用之前

守护进程要先跑起来：

```powershell
pptctl serve          # 或 pptctl status 让它自动拉起
pptctl doctor         # 自检
```

面板上的「连不上守护进程」提示会告诉你要跑哪条命令。令牌是自动从
`%LOCALAPPDATA%\ppt-agent\runtime.json` 读的，不用手填。

## 想自动构建 `.ppam`？先开关一次 VBA 工程信任

`.ppam` 是 PowerPoint 自己用 VBE 编译出来的，所以需要一个开关：
**「信任对 VBA 工程对象模型的访问」**。不开的话 `Application.VBE` 会直接抛
`Programmatic access to Visual Basic Project is not trusted`。

```powershell
pptctl addin-trust-vba              # 看当前状态与手动路径
pptctl addin-trust-vba --enable     # 一条命令代劳（要求先关掉 PowerPoint）
pptctl addin-trust-vba --disable    # 构建完建议关掉
```

**这个开关有真实的安全代价，所以它必须由你显式开、也随时可以关**：开启后任何程序都能
读写你文档里的 VBA 代码——宏病毒正是靠它扩散的。微软 [官方文档](https://learn.microsoft.com/en-in/archive/blogs/cristib/vba-how-to-programmatically-enable-access-to-the-vba-object-model-using-macros)
里的原话是别背着用户偷偷打开它。我们只提供命令，不自动执行。

### 为什么必须"先关掉 PowerPoint"

同一份文档写明：这个设置**只在应用程序启动时读取**；运行期间用任何方式（宏、脚本、
手改注册表）改它，**退出时都会被丢弃**。

所以 `--enable` 会先检查 PowerPoint 在不在跑，在跑就**拒绝执行**并告诉你正确顺序：

```
✗ PowerPoint 正在运行，这次改动不会生效。
  1) 关掉 PowerPoint（先存好你的文件）
  2) pptctl addin-trust-vba --enable
  3) 重新打开 PowerPoint，再跑 pptctl addin-build
```

静默写一个会被丢弃的值，比直接报错更糟——你会以为开好了，其实白开。

不想开这个开关也完全没关系：按下面的步骤手动导入源码，两条路等价。

## 排查

```powershell
pptctl addin-status   # 源码在不在、.frm 有没有问题、VBA 信任开没开、.ppam 建没建
```

`addin-status` 会**校验 `.frm` 里每个控件的 CLSID 是否真是本机 FM20.DLL 注册的控件**——
CLSID 写错一个字节，导入就会失败（实测：凭印象写的 `DFD0A2C2-…` 系列根本不在
FM20 的注册表里，正确的是 `D7053240-…` 按钮、`8BD21D20-…` 列表框）。

| 现象 | 原因 |
|---|---|
| 导入 `.frm` 报错 | 用了 `.frx` 引用但没分发 `.frx`；我们的 `.frm` 刻意不含任何 `.frx` 引用 |
| 面板中文变乱码 | 用了 UTF-8 的源文件；必须走 `pptctl addin-export`（会转 GBK） |
| 打不开面板 | 窗体没导进当前工程；`PPTAgentShow` 的提示里会说明 |
| 自动刷新不动 | 某些 PowerPoint 状态下 `Application.OnTime` 不可用；点「刷新」即可 |

---

# ② Web Add-in 任务窗格

**真正的停靠式任务窗格**：功能区多一个「ppt-agent」选项卡，点一下在右侧打开，
里面是当前页实时画面 + agent 的操作流水 + 跟速/急停。

## 为什么它要装证书

Office 加载项**强制 HTTPS，localhost 也不豁免**——这是官方文档的原话
（[PowerPoint add-in tutorial](https://learn.microsoft.com/ru-ru/office/dev/add-ins/tutorials/powerpoint-tutorial-yo?WT.mc_id=sitertzn_officedevcntr&tabs=jsonmanifest#3)），
也是 `office-addin-dev-certs` 这个包存在的唯一理由。WebView 只认受信任的证书，
所以要么装一张本地 CA，要么这条就用不了。

**这是这条路的全部代价**，装一次、825 天有效，撤销只是删一条。

## 三步

```powershell
# 1) 生成本地 CA + localhost 证书 + 加载项清单
pptctl addin-https

# 2) 把 CA 装进【当前用户】的受信任根（不需要管理员）
#    这一步必须由你显式执行——程序不会自己动你的受信任根存储
pptctl addin-https --trust

# 3) 重启守护进程，让 HTTPS 服务起来
pptctl restart
```

然后**旁加载清单**（一次性）：

> PowerPoint → 插入 → 我的加载项 → **上传我的加载项** → 选
> `addin/web/manifest.xml` → 确定
>
> 或者：文件 → 选项 → 信任中心 → 信任中心设置 → **受信任的加载项目录**，
> 把 `addin/web/` 加进去，再到 插入 → 我的加载项 → 共享文件夹 里选它（长期方案）

打开：**插入 → 我的加载项 → ppt-agent 监视台**，或功能区多出来的「ppt-agent」选项卡。

## 撤销

```powershell
# 删掉受信任的 CA
certutil -user -delstore Root <上面输出的指纹>
# 或者：certmgr.msc → 受信任的根证书颁发机构 → 证书 → 删除 "ppt-agent Dev CA"

# 换一张证书（比如过期了）
pptctl addin-https --force --trust
```

## 排查

```powershell
pptctl addin-https          # 看证书状态、指纹、还剩几天、CA 有没有被信任
pptctl doctor               # 看 HTTPS 服务起没起
```

| 现象 | 原因 |
|---|---|
| 面板一片空白 / 转圈 | CA 没装进受信任根（`--trust` 那步没做或没生效） |
| 上传清单报错 | 端口不是 8793 却没重新生成清单；跑 `pptctl addin-https --write-manifest` |
| 面板说「守护进程未连接」 | 守护进程没跑；`pptctl serve` |
| 想换端口 | 设 `PPT_AGENT_ADDIN_PORT` 后重新生成清单并重启 |

## 安全边界

- `/addin/` 这一组端点**不能强制令牌**——WebView 打开页面时带不了自定义请求头，
  令牌是服务端注入进页面的。所以用"Host 是不是本机"来兜：**一旦端口被隧道到公网，
  Host 会变成公网域名，就必须带令牌**，否则等于把令牌送给全世界。
- CA 装在**当前用户**作用域，不碰机器级信任，也不需要管理员。
- 证书只签 `localhost` 与 `127.0.0.1`，签不出别的域名。
