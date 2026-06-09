---
name: ppt-optimizer
description: 优化已有的 PowerPoint 演示文稿 — 放大字体、丰富内容、美化排版、统一配色。适用于任何 PPT 文件的批量优化、格式统一、内容增强。触发词：优化PPT、PPT字体放大、美化PPT、PPT改版、PPT润色、enhance PPT、polish slides。
version: 1.0.0
author: d82959222-creator
---

# PPT Optimizer — 演示文稿智能优化

## 概述

本 skill 用于对现有 PowerPoint（.pptx）文件进行全方位优化，解决常见痛点：

| 痛点 | 优化方案 |
|------|----------|
| 字体太小看不清 | 智能放大到可读尺寸 |
| 内容空洞单薄 | 基于关键词扩充内容 |
| 配色杂乱 | 统一为专业配色方案 |
| 排版拥挤/松散 | 调整布局和间距 |
| 缺少视觉层次 | 添加标题栏、卡片、图标 |

## 触发条件

当用户提出以下需求时，调用本 skill：
- "优化这个PPT" / "把这个PPT改好看一点"
- "PPT字体放大" / "字体太小了"
- "丰富一下PPT内容" / "内容太少了"
- "美化PPT" / "PPT改版"
- 任何对现有 PPT 文件的改进请求

## 工作流程

### 第一步：分析现有 PPT

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import copy

prs = Presentation("target.pptx")

# 收集信息
info = {
    "slide_count": len(prs.slides),
    "slide_width": prs.slide_width,
    "slide_height": prs.slide_height,
    "slides": []
}

for i, slide in enumerate(prs.slides):
    slide_info = {"index": i, "shapes": []}
    for shape in slide.shapes:
        shape_info = {
            "name": shape.name,
            "type": str(shape.shape_type),
            "left": shape.left,
            "top": shape.top,
            "width": shape.width,
            "height": shape.height,
        }
        if shape.has_text_frame:
            shape_info["text_runs"] = []
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    shape_info["text_runs"].append({
                        "text": run.text[:100],  # 截取前100字符
                        "font_size": run.font.size,
                        "bold": run.font.bold,
                        "color": str(run.font.color.rgb) if run.font.color and run.font.color.rgb else None
                    })
        slide_info["shapes"].append(shape_info)
    info["slides"].append(slide_info)
```

分析要点：
- 找出最小和最大字号 → 决定放大比例
- 统计文本总量 → 判断内容是否空洞
- 检查配色是否统一 → 决定是否需要重新着色
- 识别幻灯片母版/版式使用情况

### 第二步：确定优化策略

根据分析结果，向用户确认优化方向：

1. **字体放大** — 建议比例：标题 ≥36pt，正文 ≥14pt，表格 ≥11pt
2. **内容增强** — 为空洞的幻灯片补充要点、数据、说明
3. **配色统一** — 应用专业配色方案
4. **布局优化** — 调整间距、对齐、留白

### 第三步：执行优化

#### 3.1 字体缩放（核心功能）

```python
def set_font(run, name="Microsoft YaHei", size=Pt(14), bold=False, color=None):
    """设置字体 — 兼容 python-pptx 1.0.x 和 2.x"""
    run.font.name = name
    
    # 处理东亚字体（兼容不同 python-pptx 版本）
    try:
        # python-pptx 2.x API
        rPr = run._element.get_or_add_rPr()
    except AttributeError:
        # python-pptx 1.0.x API
        rPr = run._r.get_or_add_rPr()
    
    from lxml import etree
    nsmap = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    rFonts = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rFonts")
    if rFonts is None:
        rFonts = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}rFonts")
    rFonts.set("w:eastAsia", name)
    
    run.font.size = size
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


def scale_font(run, scale_factor=1.5, min_size=Pt(12)):
    """按比例放大字体"""
    if run.font.size and run.font.size < Pt(80):  # 不放大已是标题级别的
        new_size = Pt(int(run.font.size / 12700 * scale_factor))
        if new_size < min_size:
            new_size = min_size
        run.font.size = new_size


def optimize_all_text(prs, scale=1.5):
    """遍历所有文本并放大"""
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.text.strip():
                            scale_font(run, scale)
```

#### 3.2 内容增强

对于内容不足的幻灯片：
- 从标题提取关键词
- 为每个要点补充具体说明
- 添加数据/案例支撑
- 确保每页至少有 3-5 个实质信息点

```python
def enrich_slide_content(slide, context=""):
    """为幻灯片补充内容"""
    text_content = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_content.append(shape.text_frame.text)
    
    total_text = " ".join(text_content)
    
    # 判断内容是否空洞
    if len(total_text) < 50:
        # 返回内容建议，由 AI 生成补充文本
        return {
            "needs_enrichment": True,
            "current_text": total_text,
            "suggestion": "内容过少，建议补充具体要点"
        }
    return {"needs_enrichment": False}
```

#### 3.3 配色方案

```python
# 推荐配色方案
COLOR_SCHEMES = {
    "business_dark": {
        "primary": RGBColor(0x1F, 0x4E, 0x79),      # 深蓝
        "accent": RGBColor(0xC0, 0x39, 0x2B),        # 暗红
        "light": RGBColor(0xF2, 0xF2, 0xF2),         # 浅灰
        "text": RGBColor(0x33, 0x33, 0x33),           # 深灰文字
        "background": RGBColor(0xFF, 0xFF, 0xFF),     # 白色背景
    },
    "modern_teal": {
        "primary": RGBColor(0x00, 0x6D, 0x77),       # 青色
        "accent": RGBColor(0xE2, 0x95, 0x35),         # 金色
        "light": RGBColor(0xED, 0xF6, 0xF9),         # 浅青
        "text": RGBColor(0x2D, 0x2D, 0x2D),
        "background": RGBColor(0xFF, 0xFF, 0xFF),
    },
    "warm_simple": {
        "primary": RGBColor(0x8B, 0x45, 0x13),       # 暖棕
        "accent": RGBColor(0xD4, 0x8B, 0x2C),         # 暖橙
        "light": RGBColor(0xFD, 0xF5, 0xE6),         # 暖白
        "text": RGBColor(0x3D, 0x3D, 0x3D),
        "background": RGBColor(0xFF, 0xFA, 0xF5),
    }
}
```

#### 3.4 添加视觉元素

```python
def add_title_bar(slide, text, color=RGBColor(0x1F, 0x4E, 0x79), 
                  left=0, top=0, width=None, height=Inches(0.9)):
    """添加顶部标题栏"""
    if width is None:
        width = slide.part.slide_layout.slide_width if hasattr(slide.part.slide_layout, 'slide_width') else Inches(13.33)
    
    from pptx.util import Inches
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()  # 无边框
    
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.alignment = PP_ALIGN.LEFT
    tf.margin_left = Inches(0.6)
    return shape


def add_footer(slide, text="", page_num=None):
    """添加底部页脚"""
    width = Inches(13.33)
    shape = slide.shapes.add_shape(1, 0, Inches(7.0), width, Inches(0.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    shape.line.fill.background()
    
    tf = shape.text_frame
    p = tf.paragraphs[0]
    display_text = text
    if page_num is not None:
        display_text += f"  |  {page_num}"
    p.text = display_text
    p.font.size = Pt(9)
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.alignment = PP_ALIGN.RIGHT
    tf.margin_right = Inches(0.5)


def add_card(slide, left, top, width, height, title, content, 
             bg_color=None, title_color=None):
    """添加卡片式文本块"""
    if bg_color is None:
        bg_color = RGBColor(0xF2, 0xF2, 0xF2)
    
    shape = slide.shapes.add_shape(
        1, left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.fill.background()
    
    # 使用 text_frame 添加标题和内容
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.15)
    
    # 标题
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = title_color or RGBColor(0x1F, 0x4E, 0x79)
    
    # 内容
    p2 = tf.add_paragraph()
    p2.text = content
    p2.font.size = Pt(12)
    p2.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    p2.space_before = Pt(6)
    
    return shape
```

### 第四步：质量检查

优化完成后，执行以下检查：

1. **字号检查** — 确保没有字号小于 11pt
2. **溢出检查** — 确保文本框内容不超出幻灯片边界
3. **颜色一致性** — 确保通篇配色统一
4. **内容完整性** — 确保没有空白幻灯片

```python
def quality_check(prs):
    """优化后质量检查"""
    issues = []
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.size and run.font.size < Pt(10):
                            issues.append(f"Slide {i+1}: 字号过小 ({run.font.size/12700:.0f}pt) — '{run.text[:30]}...'")
                        
                        # 检查溢出
                        if shape.left + shape.width > Inches(13.5):
                            issues.append(f"Slide {i+1}: 内容可能超出右边界")
    
    return issues
```

### 第五步：保存与报告

```python
def save_optimized(prs, original_path, suffix="_优化版"):
    """保存优化后的文件"""
    import os
    dir_path = os.path.dirname(original_path)
    base_name = os.path.splitext(os.path.basename(original_path))[0]
    ext = os.path.splitext(original_path)[1]
    new_path = os.path.join(dir_path, f"{base_name}{suffix}{ext}")
    prs.save(new_path)
    
    # 报告优化内容
    print(f"✅ 已保存到: {new_path}")
    return new_path
```

## 重要技术细节

### python-pptx 版本兼容

```python
# 获取内部 XML 元素的兼容写法
def get_rPr(run):
    """兼容 python-pptx 1.0.x 和 2.x"""
    try:
        return run._element.get_or_add_rPr()   # 2.x
    except AttributeError:
        return run._r.get_or_add_rPr()          # 1.0.x
```

### Windows 编码注意事项

在 Windows 环境下，print 输出避免使用 emoji（✅❌等），GBK 编码不支持：
```python
# 正确 ✅
print(f"Saved to: {file_path}")

# 错误 ❌ — Windows GBK 报错
print(f"✅ 已保存到：{file_path}")
```

### 字体大小对照表

| 元素 | 最小 | 推荐 | 最大 |
|------|------|------|------|
| 封面标题 | 36pt | 44-56pt | 72pt |
| 封面副标题 | 18pt | 22-28pt | 32pt |
| 幻灯片标题 | 28pt | 36-44pt | 52pt |
| 正文内容 | 14pt | 16-18pt | 22pt |
| 表格文字 | 11pt | 13-14pt | 16pt |
| 页脚/注释 | 8pt | 9-10pt | 11pt |

### 幻灯片尺寸

现代宽屏 16:9：
- 宽度：13.33 inches (33.87 cm)
- 高度：7.5 inches (19.05 cm)

## 参考文件

需要深入了解时，加载以下参考文件：

- [references/python-pptx-api.md](references/python-pptx-api.md) — python-pptx 完整 API 参考
- [references/design-guide.md](references/design-guide.md) — 演示文稿设计最佳实践

## 快速开始

```python
# 最简单的优化调用
from pptx import Presentation
from ppt_optimizer import optimize_all_text, add_title_bar, save_optimized

prs = Presentation("input.pptx")
optimize_all_text(prs, scale=1.5)
save_optimized(prs, "input.pptx")
```
