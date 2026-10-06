# ppt-agent 工具参考

> 本文件由 `pptctl export-api` 生成，**请勿手改**。
> 改了代码就重跑生成，`pptctl export-api --check` 会在文档过期时报错。

共 **24 个工具**。同一份注册表同时暴露为 MCP 与 HTTP/OpenAPI，
所以这里的 schema 与 `openapi.json`、以及 MCP `tools/list` 给出的**完全一致**。

每个工具在 HTTP 上都有一个同名端点：`POST /tools/<工具名>`（需要令牌），
也可以用 `POST /call` 带 `{"name": ..., "args": {...}}` 统一调用。

## 看（只读）

### `ppt_decks` — 列出打开的演示

列出 PowerPoint 里打开着的演示，以及各自停在第几页

- 返回：每份演示的名称、路径、页数、页面尺寸，以及哪一份是活动演示
- 依赖：需要 PowerPoint（走 COM）
- 标签：read, deck

*无参数。*

### `ppt_slides` — 列出每一页

逐页列出标题与形状数量，快速了解这份演示讲了什么

- 返回：页码、标题、形状数；可指定范围并对结果做截断
- 依赖：需要 PowerPoint（走 COM）
- 标签：read, deck

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `from_slide` | integer |  | `1` | 起始页，从 1 开始 |
| `to_slide` | integer |  |  | 结束页（含）；留空表示到最后一页 |
| `limit` | integer |  | `60` | 本次最多返回多少页（范围 1..300） |

### `ppt_read_slide` — 读取一页的形状

读取某一页里所有形状的文字、位置、尺寸与类型

- 返回：该页版式、备注与形状明细；包含表格单元格、组合子形状及可用于编辑的 ref 序号路径
- 依赖：需要 PowerPoint（走 COM）
- 标签：read, slide

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 页码，从 1 开始；留空用当前正在看的那页 |
| `include_text` | boolean |  | `true` | 是否读取文字内容 |
| `include_geometry` | boolean |  | `true` | 是否读取位置与尺寸 |
| `max_shapes` | integer |  | `80` | 最多返回多少个形状（含组合子形状）（范围 1..400） |
| `max_table_cells` | integer |  | `200` | 每个表格最多读取多少个单元格；截断会明确标记（范围 1..2000） |

### `ppt_shot` — 截图查看某一页

把某一页导出成图片：给人看高清 PNG，给模型看压缩缩略图

- 返回：每页的高清图片与缩略图（含路径与可访问 URL）
- 依赖：需要 PowerPoint（走 COM）
- 标签：read, slide, vision

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 起始页码；留空用当前正在看的那页 |
| `count` | integer |  | `1` | 连续导出几页（范围 1..10） |
| `width` | integer |  | `1280` | 高清图宽度（像素）（范围 320..3840） |
| `format` | string |  | `"png"` | 高清图格式：png 清晰、jpg 更小（可选：png / jpg） |
| `thumb_width` | integer |  | `1024` | 给模型看的缩略图宽度（像素）（范围 320..1600） |
| `inline` | boolean |  | `true` | 是否生成给模型内联查看的缩略图 |
| `goto` | boolean |  | `true` | 截图后把视图切到第一张，便于用户跟随 |

### `ppt_status` — 查看 PowerPoint 状态

查看 PowerPoint 是否在运行、有哪些演示打开着、当前停在第几页

- 返回：会话状态对象：解释器、COM 会话、已打开的演示列表
- 依赖：不需要 PowerPoint（COM 挂掉也能用）
- 标签：read, diagnose

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `list_slides` | boolean |  | `false` | 是否列出每个演示的每一页标题（大文件会慢） |

## 切（幂等，重复执行结果相同）

### `ppt_goto` — 跳到某一页

把 PowerPoint 切到指定页，让用户看得见你在看哪一页

- 返回：切换后的页码
- 依赖：需要 PowerPoint（走 COM）
- 标签：visibility, slide

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer | 是 |  | 目标页码，从 1 开始 |

### `ppt_select` — 选中并跳页

跳到某一页并选中某个形状，让用户看得见 agent 正在动哪里

- 返回：当前页码与选中的形状
- 依赖：需要 PowerPoint（走 COM）
- 标签：visibility, slide

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 页码；留空表示不跳页 |
| `shape` | string |  |  | 要选中的形状名或序号；留空只跳页不选中 |

### `ppt_set_pace` — 设置跟速

设置每一步之间的停顿，好让人跟得上 agent 的操作

- 返回：生效后的停顿毫秒数
- 依赖：不需要 PowerPoint（COM 挂掉也能用）
- 标签：visibility

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `pace_ms` | integer |  | `400` | 写操作之间停顿的毫秒数；0=极速，400≈正常，1500=慢速看得清（范围 0..10000） |

### `ppt_activate` — 切到 PowerPoint 前台

把 PowerPoint 窗口带到最前面

- 返回：是否成功
- 依赖：需要 PowerPoint（走 COM）
- 标签：visibility

*无参数。*

## 做（写入）

### `ppt_save` — 保存演示

保存这份演示；也可以先另存一份备份再改

- 返回：保存结果；mode=copy 时给出备份文件路径
- 依赖：需要 PowerPoint（走 COM）
- 标签：deck

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `mode` | string |  | `"save"` | save=原地保存；copy=另存一份备份（不动当前文件）（可选：save / copy） |
| `path` | string |  |  | mode=copy 时的备份路径；留空自动放到 ppt-agent 的 backups 目录 |

### `ppt_add_slide` — 新增一页

在指定位置插入一页新幻灯片

- 返回：新页的页码、所用版式与当前总页数
- 依赖：需要 PowerPoint（走 COM）
- 标签：slide

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `layout` | string |  |  | 版式名称或序号；留空沿用最后一页的版式 |
| `index` | integer |  |  | 插到第几页；留空表示追加到末尾（范围 1..） |

### `ppt_move_slide` — 移动一页

把某一页挪到另一个位置

- 返回：移动后的页码
- 依赖：需要 PowerPoint（走 COM）
- 标签：slide

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer | 是 |  | 要移动的页码（必填） |
| `to` | integer | 是 |  | 移动到第几页（必填）（范围 1..） |

### `ppt_set_text` — 改写形状里的文字

把某一页某个形状里的文字替换掉（可同时设字号、加粗、颜色、对齐）

- 返回：被改形状的名称与序号，以及实际生效的格式
- 依赖：需要 PowerPoint（走 COM）
- 标签：shape, text

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 页码，从 1 开始；留空用当前正在看的那页 |
| `shape` | string | 是 |  | 形状名、序号或组合 ref 路径（如 3/2；可用 title 指标题占位符）；先从 ppt_read_slide 里看 |
| `text` | string | 是 |  | 新的文字内容；空字符串清空文字 |
| `font_size` | number |  |  | 字号（磅）（范围 1..4000） |
| `bold` | boolean |  |  | 是否加粗 |
| `color` | string |  |  | 文字颜色，形如 #4D6BFE |
| `align` | string |  |  | 段落对齐（可选：left / center / right / justify） |

### `ppt_add_textbox` — 加一个文本框

在某一页加一个文本框并写入文字

- 返回：新形状的名称与序号，后续可用它来改文字或位置
- 依赖：需要 PowerPoint（走 COM）
- 标签：shape, text

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 页码，从 1 开始；留空用当前正在看的那页 |
| `text` | string | 是 |  | 文本框里的文字 |
| `left` | number |  | `60.0` | 左边距（磅） |
| `top` | number |  | `60.0` | 上边距（磅） |
| `width` | number |  | `400.0` | 宽度（磅）（范围 0.01..） |
| `height` | number |  | `80.0` | 高度（磅）（范围 0.01..） |
| `font_size` | number |  |  | 字号（磅）（范围 1..4000） |
| `bold` | boolean |  |  | 是否加粗 |
| `color` | string |  |  | 文字颜色，形如 #4D6BFE |
| `align` | string |  |  | 段落对齐（可选：left / center / right / justify） |

### `ppt_set_geometry` — 移动或缩放形状

改某个形状的位置与尺寸（只改你给的项）

- 返回：改动前后的位置尺寸
- 依赖：需要 PowerPoint（走 COM）
- 标签：shape

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 页码，从 1 开始；留空用当前正在看的那页 |
| `shape` | string | 是 |  | 形状名、序号或组合 ref 路径（如 3/2；可用 title 指标题占位符） |
| `left` | number |  |  | 左边距（磅） |
| `top` | number |  |  | 上边距（磅） |
| `width` | number |  |  | 宽度（磅）（范围 0.01..） |
| `height` | number |  |  | 高度（磅）（范围 0.01..） |
| `rotation` | number |  |  | 旋转角度 |

### `ppt_add_picture` — 插入图片

把一张本地图片插到某一页（必须给绝对路径）

- 返回：新图片形状的名称与序号
- 依赖：需要 PowerPoint（走 COM）
- 标签：shape, media

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 页码，从 1 开始；留空用当前正在看的那页 |
| `path` | string | 是 |  | 图片的绝对路径 |
| `left` | number |  | `60.0` | 左边距（磅） |
| `top` | number |  | `60.0` | 上边距（磅） |
| `width` | number |  |  | 宽度（磅）；留空按原图尺寸（范围 0.01..） |
| `height` | number |  |  | 高度（磅）；留空按原图尺寸（范围 0.01..） |

### `ppt_set_notes` — 写演讲者备注

替换某一页的备注文字

- 返回：写入后的备注内容
- 依赖：需要 PowerPoint（走 COM）
- 标签：slide, notes

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 页码，从 1 开始；留空用当前正在看的那页 |
| `text` | string | 是 |  | 备注内容；传空字符串表示清空 |

### `ppt_set_table_cell` — 编辑表格单元格

修改真实 PowerPoint 表格中指定单元格的文字与格式

- 返回：目标页、表格、行列和修改前后的文字
- 依赖：需要 PowerPoint（走 COM）
- 标签：shape, text, table

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 页码，从 1 开始；留空用当前正在看的那页 |
| `shape` | string | 是 |  | 表格形状名称或 ref（支持组合序号路径，如 3/2） |
| `row` | integer | 是 |  | 行号，从 1 开始（范围 1..） |
| `column` | integer | 是 |  | 列号，从 1 开始（范围 1..） |
| `text` | string | 是 |  | 单元格文字；空字符串清空 |
| `font_size` | number |  |  | 字号（磅）（范围 1..4000） |
| `bold` | boolean |  |  | 是否加粗 |
| `color` | string |  |  | 文字颜色，如 #4D6BFE |
| `align` | string |  |  | 段落对齐（可选：left / center / right / justify） |

### `ppt_undo` — 撤销

撤销上一步操作（用的是 PowerPoint 自己的撤销栈）

- 返回：实际撤销了几步、为什么停下
- 依赖：需要 PowerPoint（走 COM）
- 标签：slide, shape

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `steps` | integer |  | `1` | 最多撤销几步（范围 1..20） |

### `ppt_open` — 打开演示

打开一个演示文稿并让它在 PowerPoint 里显示出来

- 返回：被打开演示的摘要（页数、路径、是否只读）
- 依赖：需要 PowerPoint（走 COM）
- 标签：deck

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `path` | string | 是 |  | 演示文件路径（.pptx / .ppt） |
| `read_only` | boolean |  | `false` | 以只读方式打开，避免误改原文件 |
| `bring_to_front` | boolean |  | `true` | 打开后把 PowerPoint 窗口切到最前，让用户看得见 |

### `ppt_new` — 新建演示

新建一份空白演示并显示出来

- 返回：新演示的摘要（页数、路径可能为「未保存」）
- 依赖：需要 PowerPoint（走 COM）
- 标签：deck

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `template` | string |  |  | 可选：以某个 .potx / .pptx 模板为底，需要绝对路径 |
| `bring_to_front` | boolean |  | `true` | 新建后把 PowerPoint 切到最前，让用户看得见 |

### `ppt_close` — 关闭演示

关闭一个已打开的演示（默认不保存，避免误改用户文件）

- 返回：被关闭演示的名称
- 依赖：需要 PowerPoint（走 COM）
- 标签：deck

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空表示当前活动演示 |
| `save` | boolean |  | `false` | 是否先保存再关闭。默认否——未保存的改动会被丢弃 |

## 删（破坏性，动手前先备份）

### `ppt_delete_slide` — 删除一页

删除指定的幻灯片（必须显式给页码，不会默认删当前页）

- 返回：删除了第几页、剩余总页数
- 依赖：需要 PowerPoint（走 COM）
- 标签：slide

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer | 是 |  | 要删除的页码，从 1 开始（必填） |
| `count` | integer |  | `1` | 连续删除几页（范围 1..50） |

### `ppt_delete_shape` — 删除形状

删掉某一页里的某个形状

- 返回：被删形状的名称与剩余形状数
- 依赖：需要 PowerPoint（走 COM）
- 标签：shape

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `deck` | string |  |  | 演示名、路径或序号；留空用当前活动演示 |
| `slide` | integer |  |  | 页码，从 1 开始；留空用当前正在看的那页 |
| `shape` | string | 是 |  | 形状名或序号（必填，不会默认删某个形状） |
