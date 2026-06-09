# PPT Optimizer Skill

对现有 PowerPoint 演示文稿进行智能优化 —— 放大字体、丰富内容、美化排版、统一配色。

## 适用场景

- 老板说"字太小了"😤
- 同事说"内容好空啊"😅
- 自己觉得"怎么这么丑"🤔
- 领导要求"重新美化一下"📊

## 功能

| 功能 | 说明 |
|------|------|
| 🔤 字体放大 | 智能缩放，标题 ≥36pt，正文 ≥14pt |
| 📝 内容增强 | 为空洞页面补充实质性内容 |
| 🎨 配色统一 | 内置 3 套专业配色方案 |
| 📐 布局优化 | 标题栏、卡片、页脚等视觉元素 |
| ✅ 质量检查 | 自动检测字号过小、内容溢出等问题 |

## 使用方法

在 Claude Code 中，直接将 PPT 文件拖入对话，然后说：

> "帮我优化这个PPT"

或者更具体的：

> "把这个PPT的字体全部放大，内容也丰富一下"

## 文件结构

```
ppt-optimizer/
├── SKILL.md                          # 主 skill 定义（工作流程+代码）
├── README.md                         # 本文件
└── references/
    ├── python-pptx-api.md            # python-pptx 完整 API 参考
    └── design-guide.md               # 演示文稿设计最佳实践
```

## 技术栈

- **python-pptx** — PPT 文件操作
- **lxml** — XML 底层操作（东亚字体）

## 安装

```bash
pip install python-pptx
```
