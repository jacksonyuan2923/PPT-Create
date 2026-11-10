# V4.0 动态 PPT 指南

动态能力由 `scripts/dynamic_pptx.py` 写入 PowerPoint 原生 `p:transition` 与 `p:timing`。动态不是装饰层，必须服务于信息顺序、叙事转折或交互节奏。

## 1. 等级权限

| 能力 | L1 | L2 | L3 | L4 |
|---|---|---|---|---|
| 页面转场 | 无 | 淡入/推进 | 淡入、推进、擦除、Morph | 完整转场库 |
| 对象动画 | 无 | 整页单元素淡入 | 每页最多 3 个元素 | 按时序编排 |
| 强调动画 | 无 | 无 | 脉冲、放大、变色 | 全类型强调 |
| 退出动画 | 无 | 无 | 无 | 允许 |
| 运动路径 | 无 | 无 | 无 | 允许 |
| 触发/媒体 | 仅静态媒体 | 基础媒体 | 基础媒体 | 触发、音频、视频、自动播放 |

权限值以 `level_profiles.py` 的 `AnimationRule` 为唯一来源。

## 2. 生成 API

```python
from leveled_pptx import LeveledDeckBuilder

deck = LeveledDeckBuilder("L3", mode="dynamic")

cover = deck.add_cover("年度经营复盘", "结论、证据与行动")
deck.set_page_transition(cover, "fade", duration_ms=600)

page = deck.add_metric_cards(
    "三项关键指标",
    [
        ("完成率", "92%", "样本：全部项目"),
        ("周期", "6 周", "计划值"),
        ("覆盖", "12", "单位：部门"),
    ],
)
deck.add_shape_animation(page, page.shapes[0], "wipe_up")
deck.save("dynamic.pptx", audit_path="builder_audit.json")
```

未显式指定对象动画时，`save()` 会按等级生成克制序列：L2 一页一个淡入，L3 前三项入场，L4 按叙事对象生成更多层次。L1 在 `mode="dynamic"` 时直接拒绝。

也可以在构建器初始化时提供逐页动画计划：

```python
plan = {
    "slides": [
        {
            "index": 1,
            "transition": "ripple",
            "animations": [
                {
                    "shape_index": 0,
                    "effect": "fly_up",
                    "trigger": "click",
                    "duration_ms": 600,
                }
            ],
        }
    ]
}
deck = LeveledDeckBuilder("L4", mode="dynamic", animation_plan=plan)
```

`shape_id` 或 `shape_index` 均可定位对象。缺失对象、非法效果、超密度或超时长都会中断生成。

## 3. 动效编排

- 续接关系用淡入。
- 章节推进用水平或垂直推入。
- 同构内容变化用 Morph。
- 强调只用于主结论、关键数字或状态变化。
- 每页动画对象数必须与阅读决定一致，不能用动画掩盖结构混乱。
- 图表优先让整个图表入场，避免逐数据点弹跳造成比较困难。
- L4 的跨页 Morph 要求对象在版式中具有可追踪的形状或位置关系。

## 4. 质检

```bash
python scripts/check_pptx.py 动态文件.pptx --level L3 --mode dynamic
```

检查项：

- 转场是否在白名单；
- 单页对象动画是否超限；
- 动态模式是否缺转场；
- 静态模式是否误含动画；
- L1 是否零动画；
- `slide_audit.json` 是否记录模式、动效和设计基因。

## 5. 预览与导出

```bash
python scripts/render_preview.py 动态文件.pptx preview --video --gif
```

- `preview.mp4` 优先调用 PowerPoint `CreateVideo`，保留原生转场和对象动画。
- PowerPoint 不可用时，才允许 ffmpeg 将逐页 PNG 转为静态帧视频。
- `preview.gif` 由逐页 PNG 和交叉淡入生成，只用于快速查看页面节奏，不代表对象动画已通过原生验证。
- `render_manifest.json` 的 `dynamic_exports` 必须记录实际状态和降级说明。

