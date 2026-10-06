---
name: ppt-agent
description: >
  Control the PowerPoint application that is actually running on this Windows
  machine, in real time, over COM — open/read/edit/save decks, add and delete
  slides, write text boxes, move and resize shapes, insert pictures, write
  speaker notes, take slide screenshots, and let the user watch every step.
  Use whenever the user asks you to operate their live PowerPoint, work with a
  .pptx they currently have open, "帮我把这份 PPT 改一下", "在 PowerPoint 里加一页",
  "看看第 N 页", "把某页截图给我", "演示 / 幻灯片 / PPT / deck / 汇报材料" combined with
  requests to inspect or modify it, or when ppt-agent / pptctl / mcp__ppt__ tools
  are available. Also use when the user asks to see what you are doing to their
  presentation — this skill's whole design is "the user can watch".
metadata:
  version: "0.1.0"
---

# ppt-agent：实时控制 PowerPoint

你操控的是**用户机器上真正开着的那个 PowerPoint**，不是文件读写库。用户看着屏幕，
你每动一下他都看得见。整个技能的设计都围绕这一点。

## 铁律（MUST）

1. **先看再动。** 任何写操作之前先 `ppt_decks` / `ppt_read_slide` 摸清现状。
   凭想象形状名去改，必然报 `shape/not-found`。
2. **外部文件一律绝对路径。** PowerPoint 是独立进程，相对路径会按**它自己的**工作目录
   解析，报出的错还很有误导性（实测 `AddPicture` 会报 `'str' object has no attribute 'Name'`）。
   工具会拦下相对路径，但别浪费这一轮。
3. **删除类操作必须显式给页码。** `ppt_delete_slide` 不会默认删"当前正在看的那一页"，
   这是刻意的。你也别去猜用户想删哪页。
4. **`ppt_undo` 的粒度不由我们决定。** 它用的是 PowerPoint 自己的撤销栈，一次 Undo
   **可能回退得比上一步更多**——实测会把刚新增的幻灯片一起回退掉。
   动一批破坏性操作前先 `ppt_save(mode="copy")` 留备份，并把这个限制告诉用户。
5. **关掉演示前先问。** `ppt_close` 默认不保存，未保存的改动会丢。用户没说关就别关。
6. **超时或响应丢失后先确认现状。** `com/timeout` 或 `daemon/result-unknown` 表示正在执行的操作
   可能仍然完成；先读回演示确认，再处理对话框并重启守护进程。已排队但未执行的操作会取消，
   写操作不会自动重放。`operation/queue-timeout` 表示等待前一步超时，本次未执行。
7. **组合形状用 ref。** `ppt_read_slide` 会递归返回组合的 `children` 与 `ref`（例如 `3/2`）；
   把 ref 原样交给形状编辑工具。增删或移动形状之后重新读取，序号可能已经改变。
8. **新演示保存要给 path。** 缺少保存路径时返回 `save/path-required`，避免另存为弹窗。

## 标准流程

```
ppt_status / ppt_decks        ← 现在开着什么、停在第几页
ppt_slides                    ← 这份演示讲了什么（逐页标题）
ppt_read_slide                ← 这一页有哪些形状（拿准确的形状名/序号）
   ↓ 需要改的话
ppt_goto / ppt_select         ← 先把视图切过去，让用户看得见你要动哪里
ppt_set_text / ppt_add_textbox / ppt_set_geometry / ppt_add_picture /
ppt_add_slide / ppt_move_slide / ppt_delete_slide / ppt_delete_shape / ppt_set_notes / ppt_set_table_cell
   ↓ 改完必须验证
ppt_shot                      ← 截图确认结果（你自己也能看到图）
ppt_read_slide                ← 读回来核对文字确实写进去了
```

**改完不验证就等于没做完。** `ppt_shot` 返回的图片会作为图片内容块回到你手里，
你能真的看见那一页长什么样。别只看 `ok: true` 就宣布成功。

## 工具地图（24 个）

| 想干什么 | 用哪个 |
|---|---|
| 看状态 / 有哪些演示 | `ppt_status`、`ppt_decks` |
| 看每页标题 | `ppt_slides` |
| 看某页形状与文字 | `ppt_read_slide` |
| 截图（自己看 + 给用户看） | `ppt_shot` |
| 新建演示 | `ppt_new` |
| 打开 / 关闭 / 保存 | `ppt_open`、`ppt_close`、`ppt_save` |
| 增 / 删 / 移页 | `ppt_add_slide`、`ppt_delete_slide`、`ppt_move_slide` |
| 改文字 / 加文本框 | `ppt_set_text`、`ppt_add_textbox` |
| 读表格 / 改单元格 | `ppt_read_slide`（含 cells 与截断标记）、`ppt_set_table_cell` |
| 移动缩放 / 删形状 / 插图 | `ppt_set_geometry`、`ppt_delete_shape`、`ppt_add_picture` |
| 演讲者备注 | `ppt_set_notes` |
| 让用户看得见 | `ppt_goto`、`ppt_select`、`ppt_activate` |
| 控制节奏 | `ppt_set_pace`（0=极速，400≈正常，1500=慢速） |
| 撤销 | `ppt_undo` |

## 三种"看得见"

编辑会自动显示目标页并选中目标形状。位置或格式写入失败会报错；空字符串可以清空文字、备注和表格单元格。

用户可能从三个地方看你干活，都是同一份数据、同一份中文人话：

1. **观察台网页** `http://127.0.0.1:8791/` —— 实时画面 + 每步操作流水
2. **PowerPoint 里的监视台**（`.ppam` 加载项）—— 同一份流水的浮窗
3. **你返回的截图** —— `ppt_shot` 的图片直接显示在对话里

所以：**每一步都要有明确目的**。用户看得见你在做什么，也看得见你做错了什么。

## 跟速

默认每步之间停 400ms，让用户跟得上。用户说"太慢了"就 `ppt_set_pace(pace_ms=0)`；
说"看不清"就调到 1500。批量生成时用 0，演示给人看时用 400~1500。

## 出问题怎么办

| 错误码 | 含义 | 怎么办 |
|---|---|---|
| `com/timeout` | PowerPoint 弹了模态框，COM 调用被卡住 | 让用户去处理那个对话框；必要时 `pptctl restart` |
| `com/unavailable` | 拿不到 PowerPoint 实例 | 查 `pptctl doctor`；PowerPoint 是否已安装/被策略禁用 |
| `deck/none` / `deck/not-found` | 没有活动演示 / 名字对不上 | 错误详情里会列出**已打开的演示名单**，照着改 |
| `shape/not-found` | 形状名对不上 | 错误详情里会列出**这一页所有形状名**，照着改 |
| `shape/ambiguous` | 名字匹配到多个 | 用完整名称或序号 |
| `slide/out-of-range` | 页码越界 | 错误详情里带 `total` |
| `path/relative` / `path/missing` | 路径问题 | 给绝对路径；先确认文件在 |

## 不要做的事

- **不要自己再起一个 PowerPoint**。守护进程独占 COM 会话；第二个客户端
  `Dispatch` 会与它抢，互相把调用打成 `RPC_E_CALL_REJECTED`。
- **不要在用户没要求时保存**。`ppt_save` 会覆盖用户的文件。
- **不要把守护进程当文件库用**。要批量生成 .pptx 而不需要实时可见时，
  直接用 python-pptx 更快；这个技能的价值在于**用户看得见的实时操作**。
