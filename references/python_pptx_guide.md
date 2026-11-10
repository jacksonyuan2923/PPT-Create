# V4.0.0 python-pptx 实现指南

本指南说明旧 API 如何映射到统一的四级表达策略。所有字体、网格、组件和版式权限来自 `level_profiles.py`；不得在页面函数中另设固定参数。

## 1. 策略解析

```python
from level_profiles import (
    get_profile,
    font_stack,
    resolve_layout,
    resolve_style_branch,
    match_content_layout,
)

profile = get_profile("L3")
fonts = font_stack("L3")
layout = resolve_layout(
    "L3",
    content_type="process",
    has_data=False,
)
```

推荐工作流：

1. 先根据内容类型调用 `match_content_layout()`。
2. 用 `resolve_layout()` 处理白名单回退。
3. 组件必须检查 `profile.components.allowed_components`。
4. 生成时记录策略版本、等级、版式、组件和风格分支。

L4 需要风格分支：

```python
branch = resolve_style_branch("L4", "luxury_brand")
builder = LeveledDeckBuilder("L4", style_branch="luxury_brand")
```

四个分支键：

- `technology_launch`
- `luxury_brand`
- `cultural_creative`
- `product_roadshow`

## 2. `LeveledDeckBuilder` 基础

```python
from leveled_pptx import LeveledDeckBuilder

builder = LeveledDeckBuilder("L3")
builder.add_cover(
    "项目方案",
    "用可视化证据强化关键结论",
    kicker="PROJECT REVIEW",
)
builder.add_content(
    "关键结论",
    ["结论一", "证据二", "行动三"],
    content_type="narrative",
    eyebrow="KEY POINT",
)
builder.save("presentation.pptx", audit_path="builder_audit.json")
```

### 公开 API

| 方法 | 内容任务 |
|---|---|
| `add_cover` | 封面与主题定位 |
| `add_content` | 通用内容，自动匹配版式 |
| `add_section` | 章节过渡 |
| `add_chart` | 基础原生图表页 |
| `add_closing` | 总结与行动 |
| `add_metric_cards` | 3–4 个关键指标 |
| `add_data_dashboard` | 指标 + 原生图表 |
| `add_process_steps` | 3–5 步流程 |
| `add_timeline` | 阶段与里程碑 |
| `add_icon_text_modules` | 图标/编号文字模块 |
| `add_architecture_compare` | 现状与目标或方案比较 |
| `add_case_study` | 挑战、行动、结果 |
| `add_product_roadshow` | L4 产品价值叙事 |

`add_content()` 新增参数：

```python
builder.add_content(
    title,
    points,
    content_type="timeline",
    requested_layout="",
    eyebrow="",
    image_path=None,
)
```

- `requested_layout` 仅在该等级白名单内直接采用。
- `content_type` 使用 `title`、`bullets`、`comparison`、`metrics`、`data`、`process`、`timeline`、`architecture`、`case`、`table`、`section`、`closing` 等语义。
- 上层请求不允许的版式时，`resolve_layout()` 返回该策略的合法回退。

## 3. L3 标准组件

### 指标卡

```python
builder.add_metric_cards(
    "关键指标",
    [
        ("周期", "6 周", "计划值"),
        ("覆盖", "92%", "样本：100%"),
        ("规模", "35", "单位：件"),
    ],
    insight="指标值、单位与口径在同一行完成对照。",
    eyebrow="METRICS",
)
```

### 数据看板

```python
builder.add_data_dashboard(
    "季度完成量持续上升",
    [
        ("Q1", "12", "单位：件"),
        ("Q2", "19", "单位：件"),
        ("Q3", "27", "单位：件"),
    ],
    ["Q1", "Q2", "Q3"],
    [("完成量", [12, 19, 27])],
    insight="增长来自已确认的项目数据。",
    eyebrow="DASHBOARD",
    chart_hint="bar",
)
```

数据必须来自真实资料。`insight` 不能夸大趋势或推断未验证原因。

### 流程步骤

```python
builder.add_process_steps(
    "交付流程",
    [
        ("定义", "确认受众、目标与证据。"),
        ("设计", "选择标准版式与组件。"),
        ("生成", "写入可编辑页面。"),
        ("验收", "渲染并逐页复查。"),
    ],
)
```

### 时间轴

```python
builder.add_timeline(
    "项目节奏",
    [
        ("第 1 周", "范围", "确认需求和策略。"),
        ("第 2 周", "设计", "完成页面矩阵。"),
        ("第 3 周", "交付", "生成与审查。"),
    ],
)
```

### 图标文字模块

```python
builder.add_icon_text_modules(
    "三项能力",
    [
        ("结构", "版式由内容任务决定。"),
        ("证据", "图表与案例承担证明任务。"),
        ("审查", "以实际渲染结果验收。"),
    ],
)
```

图标仅用于识别，不替代明确标题和说明。

### 架构对比

```python
builder.add_architecture_compare(
    "架构调整",
    "现状",
    ["信息层级混杂", "组件权限缺失"],
    "目标",
    ["四级策略统一", "版式与组件可校验"],
)
```

### 案例图文

```python
builder.add_case_study(
    "案例",
    "原有页面缺乏视觉重点。",
    "按证据选择标准版式。",
    "结论、证据和素材形成单一视觉中心。",
    image_path="case.png",
)
```

图片必须与案例直接相关，不得用通用图库代替证据。

## 4. L4 风格分支

```python
for branch in (
    "technology_launch",
    "luxury_brand",
    "cultural_creative",
    "product_roadshow",
):
    builder = LeveledDeckBuilder("L4", style_branch=branch)
    builder.add_cover(...)
    builder.add_content(...)
    builder.save(f"{branch}.pptx")
```

分支只改变主题视觉，不改变事实或用户内容。`level_smoke_test.py` 会比对四份输出的哈希和主题配置，防止四个分支只是换标题。

## 5. 生成审计

`LeveledDeckBuilder.slide_audit` 和 `save(audit_path=...)` 记录：

- `version`
- `level`
- `strategy`
- `style_branch`
- 每页 `layout`
- 每页 `content_type`
- `components`
- `grid_columns`

不要把 `slide_audit` 当作视觉质检证明。它只证明生成路径使用了对应策略。

## 6. 原生图表

```python
builder.add_chart(
    "完成量逐季上升",
    ["Q1", "Q2", "Q3", "Q4"],
    [("完成量", [12, 19, 27, 35])],
    chart_hint="bar",
)
```

规则：

- L1 允许必要基础图表。
- L2 使用正式图表和表格。
- L3 图表辅助核心结论，使用等宽数字和更克制的视觉样式。
- L4 可按主题调整颜色、线宽和信息图形式，但不得改变数据事实。
- 柱状图/条形图数值轴从 0 开始。
- 图表数据标签、图例和网格由策略配置控制。

## 7. 表格

L2 与 L3 可直接生成原生表格。表格必须：

- 保持行高一致。
- 表头使用策略强调色。
- 数字右对齐或统一宽度。
- 不把长文硬塞进窄列。
- 必要单位、口径与来源放在表内或紧邻注释。

## 8. 自定义页面

新增角色时：

1. 在 `LayoutRule.content_layout_map` 注册内容类型。
2. 将版式加入对应等级白名单。
3. 在 `ComponentRule` 注册使用的组件。
4. 在 builder 中调用 `_assert_layout()`、`_assert_component()`。
5. 写入 `slide_audit`。
6. 更新 `level_smoke_test.py` 覆盖。
7. 运行 `check_pptx.py --level` 和实际渲染。

## 9. 兼容性

旧别名如“一般效果”“标准效果”“精美效果”“高端效果”只作为等级解析别名使用，不再代表视觉高低。旧字段 `score_range` 已移除，应改用 `profile.acceptance_range` 与 `strategy.score_dimensions`。

旧内容可以通过 `content_type` 自动选择新策略版式；不要继续硬编码旧 `asymmetric_cards` 等版式名。

## 10. V4.0 模板与动态 API

### 模板迁移

```python
from template_parser import parse_template
from style_migrator import migrate_template

parsed = parse_template("template.pptx")
safe_style = migrate_template(parsed, "L2")
deck = LeveledDeckBuilder("L2", template_style=safe_style)
```

也可以直接把解析对象传给构建器：

```python
deck = LeveledDeckBuilder("L3", template_style=parsed)
```

构建器会再次经过 `style_migrator`，调用方不能手工绕过等级边界。

### 动态 PPT

```python
deck = LeveledDeckBuilder("L3", mode="dynamic")
slide = deck.add_content(...)
deck.set_page_transition(slide, "fade", duration_ms=600)
deck.add_shape_animation(slide, slide.shapes[0], "wipe_up")
deck.save("dynamic.pptx")
```

公开接口：

| 方法 | 作用 |
|---|---|
| `apply_template_style()` | 应用解析/迁移后的模板配置 |
| `set_page_transition()` | 写入并校验等级允许的转场 |
| `add_shape_animation()` | 写入并校验对象动画 |
| `finalize_dynamic()` | 按等级生成默认转场和对象编排 |

### 动效底层

`dynamic_pptx.py` 提供：

- `set_transition()`
- `animate()`
- `transition_summary()`
- `animation_summary()`

直接调用底层函数只用于工具开发；业务生成仍应通过 `LeveledDeckBuilder`，以便执行权限校验和审计。
