# pro-ppt-studio V4.0.0 升级报告

## 架构

本次升级把四级策略从“权限白名单”扩展为完整的设计系统：

| 层级 | 实现 |
|---|---|
| 策略配置 | `DesignGene`、`PaletteSystem`、`AnimationRule` 写入 `level_profiles.py` |
| 模板解析 | `template_parser.py` 统一解析 PPTX/POTX/THMX、PDF、DOCX、SVG、图片、视频、JSON/YAML、Markdown |
| 风格迁移 | `style_migrator.py` 按目标等级保留/过滤色彩、字体、网格、组件、效果和动效 |
| 生成执行 | `leveled_pptx.py` 接入模板配置、动态模式、转场、对象动画、媒体、路径和逐页审计 |
| 动态能力 | `dynamic_pptx.py` 写入原生 `p:transition` / `p:timing`，提供动画摘要 |
| 质检 | `check_pptx.py` 增加模板匹配、动画权限、媒体/路径、视觉中心、跨等级相似度 |
| 渲染 | `render_preview.py` 增加 PowerPoint 原生 MP4 与 GIF 页面节奏预览 |
| 编排 | `deepseek_workflow.py` 串联需求、策略、模板、大纲、生成、质检和渲染 |

## 四级差异

| 指标 | L1 | L2 | L3 | L4 |
|---|---:|---:|---:|---:|
| 声明色板 | 2 | 5 | 8 | 12 |
| 设计基因 | 原文转写 | 对称结构 | 结论先行 | 主题叙事 |
| 动效强度 | none | minimal | structured | cinematic |
| 对象动画上限 | 0 | 1 | 3 | 按叙事编排 |

最终回归的相邻视觉相似度：

| 相邻等级 | 实测 | 阈值 |
|---|---:|---:|
| L1–L2 | 0.321 | < 0.70 |
| L2–L3 | 0.260 | < 0.60 |
| L3–L4 | 0.389 | < 0.50 |

## 回归结果

- `python scripts/level_profiles.py`：通过，V4 策略、色板和动画矩阵有效。
- `python scripts/level_smoke_test.py ...`：通过，L1–L4 样例、L4 四分支、模板迁移、动态权限和视觉相似度全部通过。
- 静态 L1–L4 的 `check_pptx.py --mode static`：0 错误、0 警告。
- 动态 L2–L4 的 `check_pptx.py --mode dynamic`：0 错误、0 警告。
- `deepseek_workflow.py`：可完成显式等级下的动态 L3 生成流程；未指定等级时返回 `level_confirmation_required`，不默认选择等级。
- 动态文件已验证可被 `python-pptx` 重新解析，转场、动画、媒体关系和运动路径可读取。

## 当前环境限制

本机 PowerPoint COM 返回 `0x80070520`，且未安装 LibreOffice 与 ffmpeg。因此：

- 不能在本轮声称已完成 PowerPoint 原生动态播放或 MP4 导出验证。
- GIF 导出使用 Pillow 和逐页 PNG，可实现页面交叉淡入预览，但不代表对象动画原生验证。
- PowerPoint 原生视频导出代码已实现，待可用桌面会话重新运行。
- 模板视频解析在无 ffmpeg 时只读取容器元数据，会在输出 `limits` 中明确标注。
