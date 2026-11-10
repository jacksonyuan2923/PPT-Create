---
name: pro-ppt-studio
description: Create, redesign, or benchmark PowerPoint presentations with a V4.0.0 strategy-aware, template-aware, static-or-dynamic workflow. Use when the user asks to build a PPT/PowerPoint/演示文稿/幻灯片 from notes or source files, beautify an existing deck, apply a visual style, reference a template, make an animated deck, or perform layout and visual QA. Supports PPTX/POTX, THMX, PDF, Word/VIS guides, SVG, Markdown, images, videos, and JSON/YAML design configurations. Defaults to Chinese and supports other languages. Do not activate for general PowerPoint theory questions that do not require creating, editing, or inspecting a file.
metadata:
  short-description: "V4 四级设计基因、模板迁移、动态 PPT 与全链路质检"
  version: "4.0.0"
  tags:
    - ppt
    - powerpoint
    - presentation
    - content-strategy
    - visual-design
    - quality-check
    - template-matching
  author: pro-ppt-studio
  category: office
---

# pro-ppt-studio V4.0.0

专业级 PowerPoint 演示文稿制作、改版、模板对标和交付质检 Skill。默认中文，支持英文及其他指定语言。

V4.0.0 将 L1–L4 从权限白名单升级为四套完整设计基因。等级不存在优劣，只存在场景匹配差异。所有字体、版式、组件、网格、信息密度、模板迁移边界、动画权限和 L4 风格分支都必须由 `scripts/level_profiles.py` 统一控制。

## 核心目标

根据客户需求、原始资料、参考模板或现有 PPT，制作内容准确、结构清晰、可编辑并经过实际渲染审查的演示文稿。

所有设计必须同时满足：

1. 表达任务匹配：内容目标决定策略与版式。
2. 策略边界清晰：低等级不能使用高等级专属组件；高等级可复用基础结构，但不能退回通用模板。
3. 信息完整：视觉处理不能牺牲原始事实、限定条件、数据口径或关键文本。
4. 可编辑：必要文字、表格、图表、流程与图示优先使用 PowerPoint 原生对象。
5. 可验证：策略、版式、组件、字体、网格和渲染状态都应有机器可读记录。

## 触发条件

用户提出以下任务时使用本 Skill：

1. 从零制作 PPT / PowerPoint / 演示文稿 / 幻灯片。
2. 根据 Word、PDF、Markdown、Excel、图片、网页或其他资料制作 PPT。
3. 根据客户 PPT、Office 主题、PDF、VIS、Figma、图片或网页模板生成新演示文稿。
4. 对现有 PPT 进行改版、美化、重排、扩充、精简、修复或视觉统一。
5. 设计或优化页面结构、流程、数据看板、时间轴、案例页、图表和总结页。
6. 执行文字溢出、重叠、版式重复、字体层级、网格对齐和实际渲染检查。

仅询问 PowerPoint 理论、软件操作或一般设计建议，且不涉及文件创建、编辑或检查时，不进入完整制作流程。

## 四级内容表达策略

### L1 基础信息表达 · 信息优先策略

- 核心目标：最低成本完成准确、高效的信息传递。
- 适用：课堂作业、内部草稿、原始资料快速转写、信息留档。
- 字体：全程 1 种系统无衬线字体。
- 层级：2 级文字层级。
- 版式：只允许 `title_body` 与 `simple_split`。
- 视觉：纯白背景，无装饰元素。
- 内容：完整保留原文，不做关键词化或卡片化提炼。
- 验收：信息完整、清晰可读、无明显排版错误。

L1 不因“看起来普通”而失职。为视觉化删减原文、加入卡片或打散信息结构，反而属于失败。

### L2 规范商务表达 · 结构优先策略

- 核心目标：通过标准化结构建立正式感，提升结构化阅读效率。
- 适用：部门汇报、项目申报、迎检材料、内部评审。
- 字体：正式无衬线组合，最多 2 个字体家族。
- 层级：3 级固定层级，统一章节编号。
- 版式：固定页眉页脚、严格对称网格、色块分区。
- 组件：基础卡片、标准图文框、极简图标、标准表格与图表。
- 内容：章节化组织，保留完整论述，不压缩为空泛口号。
- 验收：正式统一、结构清晰、符合商务规范。

L2 禁止全屏蒙版标题、非对称艺术构图、图片拼贴、L3 数据看板和 L4 主题叙事。

### L3 专业视觉设计 · 重点强化策略

- 核心目标：用专业可视化手段强化核心论点，兼顾正式感和设计感。
- 适用：项目汇报、方案评审、咨询交付、经营分析、产品介绍、客户提案，以及绝大多数职场真实场景。
- 字体：标题、正文、数字等宽，最多 3 个字体家族。
- 层级：主标题、模块标题、正文、注释 4 级。
- 版式：数据看板、指标卡、流程步骤、时间轴、架构对比、案例图文、章节过渡等专业标准版式。
- 内容：结论先行，正文提炼为关键词与短句，证据用图表、流程、案例或图示承载。
- 图文：目标视觉占比 40%–60%，兼顾商务可读性。
- 验收：重点突出、逻辑可视化、现代专业、不牺牲业务可读性。

L3 是专业职场内容的核心策略，评价重点是“重点是否被强化、逻辑是否可视化、版式是否服务内容”，不是卡片数量。

### L4 高端定制设计 · 主题叙事策略

- 核心目标：围绕主题构建沉浸式视觉叙事和强识别度整套视觉系统。
- 适用：科技发布、轻奢品牌、文化文创、产品路演。
- 字体：允许主题化字体组合和有限字体视觉化。
- 版式：编辑式自由构图、沉浸大图、跨页视觉呼应、定制信息图、叙事拼贴。
- 内容：文字服务于主题叙事；长信息拆分到辅助页或附录。
- 验收：主题统一贯穿、风格独特完整、有明确艺术指导感。

L4 的四个内置风格分支：

1. `technology_launch`：科技发布。
2. `luxury_brand`：轻奢品牌。
3. `cultural_creative`：文化文创。
4. `product_roadshow`：产品路演。

L4 不得固定等同于国风或深色科技。必须根据主题选择或定义完整分支。

### V4 动态边界

| 等级 | 转场 | 对象动画 | 媒体与高级效果 |
|---|---|---|---|
| L1 | 无 | 无 | 无音视频、无动画 |
| L2 | 淡入、推进 | 最多 1 个页面级淡入 | 基础媒体，不用路径/触发 |
| L3 | 淡入、推进、擦除、Morph | 每页最多 3 个入场/强调 | 不变形为 L4 叙事 |
| L4 | 完整转场库 | 时序、强调、退出、路径 | 触发、音视频、自动播放 |

## 策略选择规则

1. 用户明确指定 L1/L2/L3/L4 时，采用用户指定。
2. 用户使用“基础信息表达”“规范商务表达”“专业视觉设计”“高端定制设计”等新名称时，直接映射到对应等级。
3. 旧名称“一般效果”“标准效果”“精美效果”“高端效果”仅作为兼容别名，不再代表视觉高低。
4. 用户未指定时，必须说明四种策略、推荐理由和场景差异，并请求确认。
5. 不得默认选择 L4，也不得把 L3 当作“升级版 L2”。等级选择只表达任务和场景。
6. 同一份内容可以因使用场景不同选择不同策略，但不能把一种策略的输出质量直接换算为另一种策略的等级分。

推荐映射：

- 原始资料完整留档、快速转写 → L1。
- 正式组织结构和商务合规 → L2。
- 需要强化关键论点并做专业可视化 → L3。
- 需要统一主题叙事和艺术指导 → L4。

## 机器可执行配置

`scripts/level_profiles.py` 是唯一策略源，包含：

- `strategy`：目标、场景、验收语句、专属评分维度。
- `font`：字体角色、字体家族上限、字号/字重/行距层级。
- `grid`：列数、边距、栏距、基线、圆角、阴影和对齐允差。
- `components`：组件白名单与禁止项。
- `layout`：版式白名单、内容类型映射、页眉页脚与方向约束。
- `images`：图片 fit、蒙版、边框、拼贴与数量。
- `charts`：图表类型、标签、图例、网格和零轴规则。
- `information`：密度系数、文本预算、视觉占比、证据政策。
- `style_branches`：L4 四类主题分支。
- `design_gene`：设计原则、叙事节奏、构图签名、视觉中心、对象性格、留白和标志动作。
- `palette_system`：2/5/8/12 色色板规模、功能色、使用比例和对比策略。
- `animation`：四级转场、对象动画、强调/退出/路径、触发、媒体、密度和时长。

核心接口：

- `get_profile(level, style_branch=None)`：取得策略配置；L4 可应用风格分支。
- `font_stack(level, style_branch=None)`：取得标题、正文、模块、强调、数字与注释字体。
- `resolve_layout(level, requested_layout, content_type=..., has_image=..., has_data=...)`：返回白名单版式。
- `match_content_layout(level, content_type, **flags)`：由内容类型自动匹配版式。
- `resolve_style_branch(level, requested_style)`：解析 L4 风格。
- `validate_profile_matrix()`：校验四级唯一性、边界、版式映射、评分维度和 L4 分支。

示例：

```python
from level_profiles import get_profile, font_stack, resolve_layout
from leveled_pptx import LeveledDeckBuilder

level = "L3"
profile = get_profile(level)
fonts = font_stack(level)
layout = resolve_layout(level, content_type="process")

builder = LeveledDeckBuilder(level)
builder.add_process_steps(title, steps)
builder.save(output_path, audit_path=audit_path)
```

L4：

```python
builder = LeveledDeckBuilder("L4", style_branch="luxury_brand")
```

动态与模板：

```python
builder = LeveledDeckBuilder(
    "L3",
    mode="dynamic",
    template_style=parsed_template,
)
slide = builder.add_content(...)
builder.set_page_transition(slide, "fade")
builder.add_shape_animation(slide, slide.shapes[0], "wipe_up")
builder.save(output_path)
```

## 模板解析与迁移

`scripts/template_parser.py` 将 PPTX/POTX/THMX、PDF、DOCX、SVG、图片、视频、JSON/YAML 和 Markdown 统一解析为 `TemplateStyle`。

`scripts/style_migrator.py` 再按目标等级过滤：

- L1 只继承黑白与单一正文族，不继承强调色、圆角、蒙版、拼贴和动画。
- L2 最多继承 5 色、2 个字体族、直角卡片、标准图文和基础图表。
- L3 最多 8 色、3 个字体族，可继承指标卡、数据图表、圆角和单一柔和阴影。
- L4 最多 12 色、4 个字体族，可继承主题纹理、蒙版、拼贴和复杂动效。

迁移结果必须包含 `retained`、`discarded`、`warnings` 和 `fingerprint`。构建器收到原始模板时也必须再次经过迁移器，不能绕过边界。

## 生成器与组件

`scripts/leveled_pptx.py` 的 `LeveledDeckBuilder` 保留原有封面、内容、章节、图表、总结 API，并新增：

- `add_metric_cards()`：关键指标卡。
- `add_data_dashboard()`：指标 + 原生图表。
- `add_process_steps()`：步骤条。
- `add_timeline()`：时间轴。
- `add_icon_text_modules()`：图标/编号文字模块。
- `add_architecture_compare()`：架构、方案或现状/目标对比。
- `add_case_study()`：挑战、行动、结果。
- `add_product_roadshow()`：L4 产品叙事。

生成器必须执行：

1. 通过 `_assert_layout()` 检查版式白名单。
2. 通过 `_assert_component()` 检查组件权限。
3. 从 `LevelProfile` 读取字号、字重、行距、网格、圆角、阴影和颜色。
4. 在 `slide_audit` 记录每页版式、内容类型、组件、风格分支和网格列数。
5. 不绕过等级配置硬编码同一套视觉常量。

## 模板对标体系

参考模板只约束视觉语言、版式、品牌和排版逻辑，不得覆盖用户内容、事实或表达策略边界。

模板优先级：

1. PPT 源模板（`.potx` / `.pptx` / `.potm`）
2. `.thmx` Office 主题
3. PDF
4. Word / VIS 品牌规范
5. Figma / PSD / XD / Sketch
6. AI / SVG
7. Markdown
8. Excel
9. 网页
10. Axure / 墨刀 / 摹客原型
11. 图片
12. 视频
13. JSON/YAML 设计配置作为参数补充

解析要求：

- PPT/THMX：读取母版、主题色、字体、占位符、页眉页脚、Logo 和图表样式。
- PDF：复刻页面对应的视觉逻辑，但不宣称获得原生母版。
- Word/VIS：读取品牌色、字体、Logo 安全边距和禁则。
- Figma/PSD/网页：提取网格、组件和配色，不直接盗用示例资产。
- 图片/视频：提取静态版式、色彩、节奏和图形语言。
- JSON/YAML：提供参数，但不能覆盖用户策略或更高优先级品牌规范。

模板与策略冲突时：

1. 用户品牌规范继续生效。
2. 目标策略边界优先于模板的越级视觉。
3. 保留可迁移的网格、色彩和组件关系，移除越级装饰。
4. 在 `design_plan.md` 和 `quality_report.md` 记录适配。

无法解析原始格式时，要求导出图、PDF、截图或结构化配置，并明确未验证范围。

## V4.0.0 设计门槛

1. 未确认策略前，不进入视觉设计与生成。
2. 文件生成只代表技术完成；视觉通过必须包含实际渲染和逐页审查。
3. L1 不强制视觉中心；L2 每页应有清晰结构；L3 每页应有单一视觉中心；L4 应有主题叙事节点和跨页视觉呼应。
4. L2–L4 不得以“标题 + 项目符号 + 空白背景”作为整套默认设计。
5. L3 的版式必须由内容任务决定：数据用数据页，流程用步骤页，时间用时间轴，案例用案例图文。
6. L3 相邻页不得连续使用同一主版式超过 2 页；同一组件不得无序重复成卡片墙。
7. L4 必须先确定风格分支和主题视觉系统，再做页面。
8. 低等级不得越级美化；高等级可复用低等级基础结构，但不得退回通用模板质感。
9. 不得为了视觉丰富度虚构数据、来源、案例、Logo、产品图或实验结果。
10. 评分只用于策略匹配度判断；硬边界违规不能被平均分抵消。
11. L1-L2、L2-L3、L3-L4 视觉相似度必须分别低于 0.70、0.60、0.50。
12. L3 的单一视觉中心由主结论、主数字或主图表承担；L4 必须有跨页视觉锤。
13. L1 必须完全静态；L2 只允许基础转场和一次页面级淡入；L3 每页最多三个对象动画；L4 才可启用复杂转场、路径和触发。

## 工作流程

按以下 7 个阶段执行。发现问题可回到前序阶段修正，但不得跳过低等级边界、实际渲染和逐页审查。

### 阶段 0：环境探测

- 运行 `python scripts/probe_environment.py --json <work_dir>/environment.json`。
- 运行 `python scripts/level_profiles.py`，确认 V4.0.0 设计基因、色板、动画和策略矩阵有效。
- 首次运行、升级或怀疑工具链异常时，运行 `python scripts/smoke_test.py <work_dir>/smoke`。
- 修改策略或构建器后，运行 `python scripts/level_smoke_test.py <work_dir>/level-check --json <work_dir>/level-check.json`。
- 探测结果中渲染器只列为候选；运行时成功必须来自本次渲染清单。
- `render_preview.py` 必须通过 `--environment <work_dir>/environment.json` 读写 PowerPoint COM 状态。

渲染决策：

1. PowerPoint 原生 COM：优先 pywin32，失败后再尝试 PowerShell COM。
2. LibreOffice：转 PDF 后逐页转 PNG。
3. Artifact Tool：仅当环境显式提供 Node、`@oai/artifact-tool` node_modules 和 helper 时调用。
4. 全部失败：`verification_status=unverified`。

遇到 `0x80070520` 时：

- 将本次错误标记为“临时故障”，同时把 PowerPoint COM 标记为“当前环境的永久性限制”。
- 立即停止整个 PowerPoint COM 后端族，不继续尝试 pywin32 或 PowerShell COM。
- 直接流转 LibreOffice → Artifact Tool。
- 在 `environment.json` 写入 `retry_policy=skip_powerpoint_com`、`auto_retry_allowed=false`、错误码、hint、完整 detail 与 stack。
- 后续同一环境默认跳过 PowerPoint COM；只有环境改变后显式使用 `--reset-powerpoint-com-state` 才清除限制。
- 即使 Artifact Tool 导出全部成功，`render_backend_kind` 仍必须是 `alternative_artifact_tool`，`verification_status` 仍必须是 `rendered_alternative_artifact_tool`。不得改写成 `rendered_native_powerpoint`。

每个后端只尝试一次。每次使用独立临时目录；只有导出页数、清单页数和磁盘图片数均等于 PPTX 页数时才提升成功。`render_manifest.json` 必须记录：

- `skill_version`
- `backend`
- `backend_kind`
- `render_backend_kind`
- `native_powerpoint_verified`
- `verification_status`
- `environment_file`
- `expected_count`
- `rendered_count`
- `pngs`
- `backend_failures`
- `skipped_backends`
- `powerpoint_com_state`
- `powerpoint_com_attempted`
- `visual_quality_status`
- `design_gene`
- `palette_declared_colors` / `palette_color_count`
- `motion_intensity` / `animated_slides` / `transitioned_slides`
- `template_match`
- `dynamic_exports`

`backend_failures` 中针对 `0x80070520` 的记录必须至少包含 `error_code`、`hint`、`detail`、`stack`、`attempted`、`skipped_retry`；不得只保存截断摘要或吞掉 PowerShell 堆栈。

### 阶段 1：需求解析与澄清

识别：

- 主题与核心业务目标。
- 受众、场景和演示时长。
- 页数、语言、输出格式和可编辑要求。
- 表达策略 L1–L4。
- 模板类型、品牌规范、配色、字体和素材。
- 必须出现、禁止出现和需要保留的原文。
- 是否需要研究、来源和引用。

只有关键缺失会改变结果时才提问。未指定策略时必须澄清四种选项。非关键缺失采用合理专业默认值并在计划中记录。

### 阶段 2：内容与逐页结构

1. 提取事实、数据、观点、限制和结论。
2. 区分用户提供事实、外部查证信息、示例和设计假设。
3. 构建叙事链或业务结构。
4. 为每页确定：
   - 页面角色
   - 观众结论
   - 内容证据
   - 内容类型
   - 版式
   - 视觉中心
   - 组件
   - 素材需求
   - 与相邻页的节奏差异
5. 检查重复、遗漏、逻辑断裂和事实冲突。
6. 内容过多时拆分页面或调整层级，不直接删除重要信息。

L1 不得执行会损失原意的概括；L2 保留完整论述结构；L3 可以在保留限制条件的前提下提炼；L4 可将详细证据放到附录。

### 阶段 3：设计规划

必须读取：

- `references/visual_engine.md`
- `references/design_system.md`

生成 `design_plan.md`，至少包含：

- `skill_version`: `4.0.0`
- `selected_expression_level`
- `expression_strategy`
- `strategy_objective`
- `scene_match_reason`
- `acceptance_criteria`
- `font_level_spec`
- `grid_spec`
- `component_permissions`
- `layout_level_spec`
- `information_density_level`
- `style_branch`（L4 必填）
- `reference_template_type`
- `reference_style_summary`
- `template_level_compatibility`
- `brand_spec_lock`
- 配色、字体、图片、图表和图示策略
- 逐页规划表

调用 `get_profile()`、`font_stack()`、`resolve_layout()`、`match_content_layout()` 和必要的 `resolve_style_branch()`，确保计划中的参数来自配置而不是自然语言假设。

### 阶段 4：素材与数据

- 图片与内容直接相关，使用 Pillow 验证尺寸和方向。
- 等比缩放/裁切，禁止拉伸和遮挡主体。
- 记录真实素材来源；生成图不得冒充真实证据。
- 图表、表格和流程优先原生可编辑对象。
- 数据图必须标明单位、范围、来源和比较口径。
- L1 图片无装饰处理；L2 标准框；L3 条件使用圆角/蒙版；L4 按主题分支组织多图叙事。
- 无合适素材时使用原生图示、数据图或留白，不为丰富度堆图库。

### 阶段 5：生成

- 默认从 `scripts/leveled_pptx.py` 的 `LeveledDeckBuilder` 开始。
- 先通过内容类型自动匹配版式；必要时再传入 `requested_layout`，但仍必须经过白名单。
- 使用 L3 标准组件承载指标、流程、时间轴、架构和案例。
- L4 必须传入 `style_branch`。
- 新页面角色必须读取 `LevelProfile`，并登记版式、组件、网格和审计元数据。
- 生成后保存 builder audit，供等级回归和问题定位。
- 所有正文预算、字体家族、层级和组件权限由策略配置执行，不能在页面函数中绕过。

### 阶段 6：技术质检、渲染与视觉审查

必须读取 `references/quality_checklist.md`。

1. 文件与结构：`check_pptx.py` 检查可解析性、页数、比例、空白页、占位文本和内容完整性。
2. 文本与内容：检查错别字、数字、单位、术语、事实、来源、重复和逻辑。
3. 几何与策略：`check_pptx.py --level L1/L2/L3/L4` 检查溢出、重叠、越界、网格、字号数量、字体上限和版式重复。
4. 代码层策略：`level_smoke_test.py` 检查字体数量、版式权限、信息密度、组件差异、L3 多样性和 L4 四分支。
5. 实际渲染：`render_preview.py` 导出完整页集和缩略图。
6. 视觉审查：先看缩略图，再逐页查看高清。检查文字、图片、图表、层级、重复度和整体一致性。
7. 修复：记录页码、原因、动作，重新生成、渲染和复查。

视觉评分使用 `visual_engine.md` 的分策略、分维度标准。每个策略最低通过分 80；任何硬边界违规不得通过。

### 阶段 7：交付

默认交付：

1. `presentation.pptx`
2. `presentation.pdf`（环境支持且用户需要时）
3. `preview/`：逐页 PNG、缩略图和 `render_manifest.json`
4. `design_plan.md`
5. `quality_report.md`

文件名按项目命名，避免覆盖原始资料和已有客户文件。

## 质量门槛

必须分别记录：

- `generation_status`：PPTX 保存、解析和内容完整。
- `technical_quality_status`：结构、内容、几何和策略硬边界。
- `render_backend` / `render_backend_kind`
- `environment_file`
- `render_verification_status`
- `render_expected_pages` / `render_rendered_pages`
- `render_backend_failures`
- `visual_review_scope`
- `visual_quality_status`
- `strategy_match_score`
- `strategy_dimension_scores`
- `strategy_boundary_violations`
- `reverse_underperformance_pages`
- `grid_alignment_check`
- `layout_diversity_check`
- `font_hierarchy_check`
- `component_permissions_check`
- `unverified_items`

视觉通过需要：

- 实际渲染完成，或替代渲染明确标注。
- 整套缩略图和每个交付页高清预览已查看。
- 无未解释的文字截断、严重重叠、缺图、错位和越界。
- 策略匹配分不低于 80。
- 无越级组件、策略边界违规或反向失职。
- L4 主题系统跨页成立。

渲染成功、文件生成、脚本返回 0、平均分高，均不能单独代表视觉或策略匹配通过。原生 PowerPoint 未执行时，不得声称已完成原生验证。

## 资源

- `scripts/level_profiles.py`：V4.0.0 设计基因、策略、色板、动画和矩阵校验。
- `scripts/leveled_pptx.py`：策略驱动构建器和 L3/L4 组件库。
- `scripts/level_smoke_test.py`：设计基因、视觉相似度、模板迁移、动态权限和越级拒绝测试。
- `scripts/check_pptx.py`：结构、几何、模板匹配、动画合规、设计原则和跨等级差异检查。
- `scripts/template_parser.py`：多格式模板解析。
- `scripts/style_migrator.py`：目标等级风格继承与越级过滤。
- `scripts/dynamic_pptx.py`：原生转场、对象动画、动画摘要。
- `scripts/deepseek_workflow.py`：需求、策略、大纲、生成、质检、渲染编排。
- `scripts/render_preview.py`：PowerPoint → LibreOffice → Artifact Tool 回退渲染，可选 MP4/GIF。
- `scripts/probe_environment.py`：生成、模板、动态和渲染能力探测。
- `scripts/smoke_test.py`：三页工具链与渲染回退契约测试。
- `scripts/artifact_tool_render.mjs`：显式配置的 Artifact Tool helper。
- `references/visual_engine.md`：策略、版式矩阵、L4 风格和评分。
- `references/design_system.md`：字体、网格、配色、组件权限。
- `references/quality_checklist.md`：五层质检与策略匹配清单。
- `references/python_pptx_guide.md`：API、L3 组件和 L4 分支调用方式。
- `references/template_matching_guide.md`：模板解析、继承矩阵和匹配度。
- `references/dynamic_ppt_guide.md`：动画权限、时序、媒体和导出。
- `references/deepseek_integration_guide.md`：Codex/DeepSeek 七步工作流。
- `references/level_difference_standard.md`：设计基因、量化阈值与人工验收。
- `requirements.txt`：PPT 生成与图像处理必需依赖。
- `requirements-optional.txt`：色彩聚类、字体解析、视频处理和相似度增强依赖。

## 最终回复

简要说明：

1. 成品路径和页数。
2. 采用的表达策略及其场景匹配理由。
3. 主要版式、组件和 L4 风格分支。
4. 生成、技术、渲染、视觉和策略匹配状态。
5. L3 版式多样性或 L4 主题一致性结果。
6. 实际后端；替代渲染是否使用。
7. 已修复问题、剩余限制和未验证项。
