# V4.0.0 五层质检与策略匹配清单

文件生成、技术检查、实际渲染、视觉审查和策略匹配必须分别记录。渲染成功不等于视觉通过，平均分高也不能抵消硬边界违规。

## 零、环境与基线

- [ ] 运行 `python scripts/probe_environment.py --json environment.json`。
- [ ] 确认探测结果包含 `skill_version=4.0.0`。
- [ ] 运行 `python scripts/level_profiles.py`，确认四个策略指纹唯一、维度总分均为 100、版式映射均落在白名单。
- [ ] 修改生成器或策略配置后，运行 `python scripts/level_smoke_test.py <out> --json level_report.json`。
- [ ] 安装或首次运行工具链时，运行 `python scripts/smoke_test.py <out>`；该测试同时检查回退契约。
- [ ] 记录实际可用渲染后端；注册表或命令存在仅代表候选。
- [ ] 渲染顺序固定为 PowerPoint 原生 COM → LibreOffice → 显式 Artifact Tool → 未验证。
- [ ] 每个失败后端只尝试一次；记录实际错误并继续。
- [ ] 每次尝试使用独立临时目录；只有实际导出页数等于 PPTX 页数才提升结果。
- [ ] 替代渲染不得标记为 PowerPoint 原生通过。
- [ ] `0x80070520` 在 `environment.json` 中标记为临时运行故障，同时记录为当前环境的永久性限制。
- [ ] 捕获 `0x80070520` 后立即跳过整个 PowerPoint COM 后端族，不重复尝试 pywin32/PowerShell COM。
- [ ] `backend_failures` 保存 `error_code`、hint、完整 detail 和 stack，不吞 PowerShell 堆栈。
- [ ] Artifact Tool 成功后，`render_backend_kind=alternative_artifact_tool`，不得改写为 `rendered_native_powerpoint`。

## 一、文件与结构

- [ ] 文件存在、非空、可解析。
- [ ] 页数、比例和输出格式符合要求。
- [ ] 无空白页、重复页、占位文本、调试文本和模板残留。
- [ ] 必需内容、图片、表格和图表存在。
- [ ] PPTX 为可编辑对象，不以整页截图替代原生内容。

```powershell
python scripts/check_pptx.py <file.pptx> --level <L1|L2|L3|L4> --json check.json
```

## 二、文本与内容

- [ ] 错别字、术语、数字、单位、日期和名称一致。
- [ ] 标题反映页面实际内容，不写空泛口号。
- [ ] 结论有证据支撑；预测、假设和实际数据明确区分。
- [ ] 引用、来源、版权、口径和免责声明齐全。
- [ ] L1 完整保留原文，不为了视觉化删减、概括或改写关键信息。
- [ ] L2 的章节、编号和模块层级完整。
- [ ] L3 的结论先行且关键词没有丢失关键限定条件。
- [ ] L4 的主题叙事没有压过事实与可读性。

## 三、布局与几何

- [ ] 文本框无意外溢出、截断或越界。
- [ ] 图片、图表、表格和文字无意外重叠。
- [ ] 页边距、栏间距、行距和组件内边距一致。
- [ ] 图片等比裁切，不拉伸、不压缩、不遮挡主体。
- [ ] 标题、页码、图例和数据标签对齐。
- [ ] 字号可读，不为塞内容无节制缩小。

### L3 专项

- [ ] 使用 12 列网格；主要容器边缘对齐网格，允许误差 0.08in。
- [ ] 统一组件圆角约 0.12in；无多种圆角混用。
- [ ] 仅使用单一柔和环境阴影，不使用多层发光/阴影。
- [ ] 每页只有一个视觉中心。
- [ ] 相邻页面不机械复用同一主版式；连续同版式不超过 2 页。
- [ ] 数据页有单位、范围和来源；图表标题表达结论。
- [ ] 卡片数量与信息关系匹配，不把所有内容做成卡片墙。
- [ ] 关键数字使用等宽数字字体，比较列宽度稳定。

### L2 专项

- [ ] 固定页眉页脚、章节号和对称网格。
- [ ] 基础卡片为直角，不出现 L3 数据看板或 L4 蒙版拼贴。
- [ ] 图片使用标准图文框；标题带和正文分区稳定。

### L1 专项

- [ ] 纯白背景，无装饰色块、背景图、卡片、圆角或图标模块。
- [ ] 仅使用 `title_body` 与 `simple_split`。
- [ ] 字体 1 种、主层级 2 级。
- [ ] 允许长文本，但不得截断或过度概括。

### L4 专项

- [ ] 已明确四种风格分支之一，或其定制变体。
- [ ] 主题色、字体、图像、图形和跨页节奏形成完整系统。
- [ ] 自由构图仍保持阅读顺序和视觉层级。
- [ ] 图片和生成素材不被冒充为真实产品、事件或实验证据。
- [ ] 科技、轻奢、文创、产品路演分支没有落入默认模板。

## 四、实际渲染

```powershell
python scripts/render_preview.py <file.pptx> <preview_dir> `
  --environment <work_dir>/environment.json `
  --artifact-node <node.exe> `
  --artifact-node-modules <node_modules>
```

- [ ] 先查看整套缩略图，再逐页查看高清预览。
- [ ] 检查文字截断、换行、字体替换、缺图、遮挡、错位和风格漂移。
- [ ] 检查过密、过空、对比不足和信息层级。
- [ ] 检查页面是否有明确视觉中心，而不是普通文档排版。
- [ ] 检查图片与观点相关性，图表与数据口径一致性。
- [ ] 检查版式重复度和章节节奏。
- [ ] 记录实际后端、类别、验证状态、预期页数、实际页数、失败原因和未验证项。
- [ ] 验证 `render_backend_kind` 只使用后端类别值，`verification_status` 才使用 `rendered_*` 值。
- [ ] 同一 environment 的第二次渲染中，`powerpoint_com_attempted=false`，PowerPoint COM 因持久限制被跳过。
- [ ] manifest 的 `environment_file` 指向实际保存熔断状态的 `environment.json`。
- [ ] 渲染成功后，`visual_quality_status` 仍默认为 `not_assessed`，直到完成审查。

## 五、策略匹配度

### 硬边界

任一违规即不得判定通过：

- [ ] 字体家族数不超过配置上限。
- [ ] 主字体层级数量符合策略。
- [ ] 版式和组件均在白名单。
- [ ] L1 没有卡片、蒙版、图标、背景图或圆角装饰。
- [ ] L2 没有 L3/L4 专属数据看板、非对称、蒙版或拼贴。
- [ ] L3 没有 L4 纹理拼贴、厚重叙事包装或多重阴影。
- [ ] L4 没有退回通用模板质感。

### 反向失职

- [ ] L1 未因卡片化、关键词化或过度美化牺牲信息完整性。
- [ ] L2 未用大图蒙版或艺术排版替代正式结构。
- [ ] L3 未只换视觉主题而没有突出论点、可视化逻辑或版式变化。
- [ ] L4 未只堆视觉元素而没有主题系统和事实证据。

### 评分

每个策略总分 100，最低通过 80。必须按等级专属维度评分：

- L1：信息完整、可读性、基础结构、基础版式、边界、执行。
- L2：结构匹配、内容组织、正式一致、层级、商务组件、执行。
- L3：重点突出、逻辑可视化、版式丰富、信息层级、商务专业、执行。
- L4：主题叙事、视觉系统、构图创新、跨页一致、信息支撑、执行。

高平均分不能抵消任何硬边界失败。

### 可量化校验

- [ ] `level_smoke_test.py` 中四档字体数量与层级符合配置。
- [ ] L3 至少覆盖 7 种不同版式；测试样本实际达到 10 种不同布局。
- [ ] L3 版式重复不超过 2 页，网格、字号、组件差异均有 JSON 审计。
- [ ] L4 四个风格分支生成文件哈希不同，分别验证 chart theme、布局和调色。
- [ ] L1/L2 调用越级组件时必须抛错。
- [ ] `check_pptx.py --level L3` 检查网格、字号数量、字体上限和主标题一致性。

## 六、修复与回归

1. 记录页码、问题类型、原因和修复方式。
2. 优先调整结构、布局、文本框尺寸、组件选择或内容表达。
3. 修复后重新生成或重排受影响页。
4. 重新执行结构检查、策略检查和渲染。
5. 复查受影响页及整套缩略图；最多三轮自动修复。
6. 未解决项必须写入 `quality_report.md`，不得降低口径掩盖。

## 七、交付状态

- `生成通过`：文件已保存、可解析，内容完整。
- `技术质检通过`：结构、内容、几何与策略硬边界通过，或剩余问题已明确披露。
- `替代渲染通过`：指定兼容后端完成完整页集导出，不等于 PowerPoint 原生。
- `视觉通过`：已完成缩略图和逐页审查，无关键缺陷。
- `策略匹配通过`：总分 ≥80，且无硬边界或反向失职。

`quality_report.md` 至少记录：

- `skill_version`
- `selected_expression_level`
- `expression_strategy`
- `generation_status`
- `technical_quality_status`
- `render_backend` / `render_backend_kind`
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

## 八、V4.0 模板、差异与动态质检

### 模板

- [ ] 模板解析输出包含 `palette`、`fonts`、`grid`、`components`、`rhythm`、`limits`。
- [ ] 原始模板已通过 `style_migrator.py` 映射到目标等级。
- [ ] `discarded` 记录所有越级色板、组件、圆角、蒙版、纹理、拼贴和动画。
- [ ] `check_pptx.py --template-style` 的配色余弦、字体匹配和综合匹配度已记录。

### 等级差异

- [ ] `level_profiles.py` 输出 `design_genes`、`palette_matrix` 和 `animation_matrix`。
- [ ] `level_smoke_test.py` 的 L1-L2、L2-L3、L3-L4 视觉相似度均低于阈值。
- [ ] L3 每页有单一视觉中心；L4 有跨页主题视觉锤。
- [ ] 不得仅通过改变圆角、字体数量或背景色冒充等级差异。

### 动态

- [ ] L1 无 `p:transition` 和 `p:timing`。
- [ ] L2 只使用基础转场与最多一个页面级入场。
- [ ] L3 每页最多三个对象动画，使用克制入场和强调。
- [ ] L4 可按叙事使用复杂转场、强调、退出、路径和触发。
- [ ] `--mode dynamic` 检查动态模式；`--mode static` 检查静态泄漏。
- [ ] 视频导出优先 PowerPoint 原生；ffmpeg 仅作为静态帧降级。
- [ ] GIF 是页面节奏预览，不得表述为对象动画已原生验证。
