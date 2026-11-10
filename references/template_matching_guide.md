# V4.0 模板对标与风格迁移指南

模板对标的最终产物是**目标等级可执行的设计配置**，不是对参考文件外观的逐像素复制。解析结果必须经过等级边界过滤，才能交给生成器。

## 1. 处理链路

```text
参考文件
  → template_parser.py
  → TemplateStyle JSON
  → style_migrator.py
  → 等级安全设计配置
  → LeveledDeckBuilder.apply_template_style()
  → 实际生成与渲染对比
```

```bash
python scripts/template_parser.py 模板.pptx --json work/template.json
python scripts/style_migrator.py 模板.pptx L2 --json work/template_L2.json
```

```python
from template_parser import parse_template
from style_migrator import migrate_template
from leveled_pptx import LeveledDeckBuilder

parsed = parse_template("template.pptx")
style = migrate_template(parsed, "L2")
deck = LeveledDeckBuilder("L2", template_style=style)
```

## 2. 标准输出

`TemplateStyle` 至少包含：

| 字段 | 内容 |
|---|---|
| `palette` | 背景、面板、正文、次级文字、强调、分隔线、对比、高亮 |
| `fonts` | 标题、正文、强调、数字与主题字体角色 |
| `grid` | 列数、边距、栏距、顶部留白、组件圆角 |
| `components` | 检测到的卡片、图文、图表、表格、图标、拼贴等信号 |
| `rhythm` | 信息密度、转场提示、自动播放间隔 |
| `detected` | 页面、母版、版式、媒体、构件尺寸等证据 |
| `limits` | 无法直接恢复或需要人工确认的内容 |
| `confidence` | 解析置信度，不是设计质量分 |

## 3. 13 类来源

| 优先级 | 来源 | 可提取内容 | 限制 |
|---:|---|---|---|
| 1 | PPTX / POTX / POTM | 主题色、主题字体、母版、版式、页面用色、媒体、构件数量 | 不自动复制整套母版 |
| 2 | THMX | Office 主题色与字体 | 无页面内容结构 |
| 3 | PDF | 页面数、内容流颜色、嵌入字体名 | 不声称获得原生母版 |
| 4 | Word / VIS | 品牌色、字体、图片信号 | 仅提取可识别品牌规则 |
| 5 | Figma / PSD / XD / Sketch | 需导出 JSON、SVG、PPTX 或 PNG | 不直接声明完整可编辑还原 |
| 6 | SVG | 尺寸、颜色、路径/矩形等图元 | 滤镜和复杂效果不是一对一映射 |
| 7 | Markdown | 标题层级、表格、Mermaid | 仅作为内容结构 |
| 8 | Excel | 需先转为数据表或图表配置 | 视觉模板信号有限 |
| 9 | 网页 | 需截图或导出设计配置 | 不做远程页面抓取假设 |
| 10 | 原型 | 需导出图片或配置 | 交互逻辑不等于演示节奏 |
| 11 | 图片 | 色彩、明暗、分区、构图原型 | 字体、原文和真正母版不可恢复 |
| 12 | 视频 | 容器信息、分辨率、帧率、时长；有 ffmpeg 时抽帧 | 无 ffmpeg 时只读取元数据 |
| 13 | JSON / YAML | 显式色板、字体、网格、组件配置 | 仍需等级边界过滤 |

## 4. 等级继承矩阵

| 模板元素 | L1 | L2 | L3 | L4 |
|---|---|---|---|---|
| 黑白灰配色 | 是 | 是 | 是 | 是 |
| 品牌强调色 | 否 | 少量 | 单强调体系 | 完整主题 |
| 基础字体 | 1 族 | ≤2 族 | ≤3 族 | ≤4 族 |
| 直角卡片 | 否 | 是 | 是 | 是 |
| 圆角与柔和阴影 | 否 | 否 | 是，统一参数 | 主题控制 |
| 数据看板/指标卡 | 否 | 否 | 是 | 是 |
| 图片蒙版 | 否 | 否 | 克制使用 | 主题化使用 |
| 拼贴/纹理/异形 | 否 | 否 | 否 | 是 |
| 对象动画 | 否 | 整页单次 | 元素序列 | 全类型叙事 |

迁移器必须输出 `retained` 和 `discarded`。交付说明要明确哪些模板特征被保留、哪些因越级被过滤。

## 5. 匹配度验收

`check_pptx.py --template-style` 计算：

- 配色余弦相似度；
- 字体匹配率；
- 综合匹配度 `0.65 × 配色 + 0.35 × 字体`；
- 等级边界违规。

模板匹配度不能覆盖内容准确性、策略边界和可读性。若参考模板与目标等级冲突，优先执行目标等级边界，并记录原因。

