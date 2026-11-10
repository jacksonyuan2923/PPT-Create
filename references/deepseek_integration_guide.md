# V4.0 DeepSeek / Codex 集成指南

`scripts/deepseek_workflow.py` 把模型内容规划与确定性生成、质检连接起来。模型负责理解和表达，代码负责等级边界、模板过滤、文件生成和验证。

## 1. 七步链路

1. `parse_requirement()`：提取主题、场景、页数、等级、模板、静态/动态和必须保留内容。
2. `select_strategy()`：返回设计基因、版式白名单、组件白名单和动画规则。
3. `parse_template_file()` + `migrate_template()`：解析模板并过滤越级元素。
4. `generate_outline()`：生成逐页角色、内容类型、版式、视觉中心、组件和证据。
5. `build_pptx()`：调用 `LeveledDeckBuilder`，保存 `builder_audit.json`。
6. `full_quality_check()`：执行结构、模板、动画和策略质检。
7. `render_result()`：渲染逐页 PNG、缩略图，并可按需输出视频或 GIF。

## 2. 强制等级确认

用户未指定等级时，`parse_requirement()` 返回：

```json
{
  "requires_level_confirmation": true,
  "level": null
}
```

此时不得默认 L3/L4。模型应说明：

| 等级 | 适用任务 | 视觉逻辑 |
|---|---|---|
| L1 | 原文留档、快速转写 | 纯信息、零装饰 |
| L2 | 正式汇报、内部评审 | 对称结构、商务规范 |
| L3 | 经营分析、方案评审 | 结论先行、单一视觉中心 |
| L4 | 发布、品牌、文创、路演 | 主题叙事、跨页视觉锤 |

## 3. 模型提示词

`DEEPSEEK_SYSTEM_PROMPT` 包含不可绕过的规则：

- 先策略后版式，再填内容；
- 不得模板越级；
- 不得用“标题 + 列表”应付所有页面；
- L3 必须有单一视觉中心；
- L4 必须有主题系统和叙事节奏；
- 不得虚构事实或来源；
- 生成后必须质检和实际渲染。

`generate_outline(model_callable=...)` 接收模型返回 JSON，并对页数、版式和组件做二次校验。任何非法项都会抛错，不会静默进入生成器。

## 4. Python 调用

```python
from deepseek_workflow import run_workflow

result = run_workflow(
    {
        "prompt": "制作 L3 项目复盘，8 页，参考 template.pptx",
        "template_path": "template.pptx",
        "page_count": 8,
    },
    output_dir="build/project",
)
```

输出：

- `design_plan.md`
- `presentation.pptx`
- `builder_audit.json`
- `quality_report.json`
- 渲染后可得到 `preview/`、`render_manifest.json`、`preview.gif` 或 `preview.mp4`

## 5. 任务状态

| 状态 | 含义 |
|---|---|
| `level_confirmation_required` | 需要用户确认策略，禁止继续生成 |
| `generated` | PPTX 已生成并完成自动化检查 |
| `unverified` | 渲染后端不可用，不能声称视觉通过 |
| `failed` | 明确的模型、模板、生成或质检错误 |

自动化通过只代表结构和权限通过。视觉审查仍需查看整套缩略图与每个交付页。

