# python-pptx API 参考

## 安装

```bash
pip install python-pptx
```

## 基本操作

### 打开/创建/保存

```python
from pptx import Presentation

# 打开已有文件
prs = Presentation("input.pptx")

# 创建新文件
prs = Presentation()  # 使用默认模板

# 保存
prs.save("output.pptx")
```

### 幻灯片尺寸

```python
from pptx.util import Inches, Cm, Pt, Emu

# 读取尺寸
width = prs.slide_width   # 默认 10 inches (4:3)
height = prs.slide_height  # 默认 7.5 inches

# 设置为 16:9 宽屏
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
```

## 幻灯片操作

### 添加幻灯片

```python
# 使用空白版式（版式索引 6 通常是空白）
blank_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_layout)

# 查看所有可用版式
for i, layout in enumerate(prs.slide_layouts):
    print(f"{i}: {layout.name}")
```

### 删除幻灯片

```python
def delete_slide(prs, slide_index):
    """删除指定索引的幻灯片"""
    rId = prs.slides._sldIdLst[slide_index].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[slide_index]
```

## 形状操作

### 添加形状

```python
from pptx.enum.shapes import MSO_SHAPE

# 矩形
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)

# 圆角矩形
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)

# 椭圆
shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)

# 常用 MSO_SHAPE 值:
# RECTANGLE = 1
# ROUNDED_RECTANGLE = 5
# OVAL = 9
# CHEVRON = 55
# PENTAGON = 56
```

### 设置填充

```python
from pptx.dml.color import RGBColor

# 纯色填充
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)

# 无填充
shape.fill.background()

# 渐变填充
shape.fill.gradient()
shape.fill.gradient_angle = 90.0
# 设置渐变色停止点
```

### 设置边框

```python
# 无边框
shape.line.fill.background()

# 实线边框
shape.line.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
shape.line.width = Pt(1.5)
```

## 文本操作

### 文本框

```python
from pptx.util import Inches, Pt

# 添加文本框
left = Inches(1)
top = Inches(2)
width = Inches(11)
height = Inches(3)
textbox = slide.shapes.add_textbox(left, top, width, height)

# 获取 text_frame
tf = textbox.text_frame
tf.word_wrap = True

# 设置边距
tf.margin_left = Inches(0.5)
tf.margin_right = Inches(0.5)
tf.margin_top = Inches(0.3)
tf.margin_bottom = Inches(0.3)
```

### 段落和文本运行

```python
# 第一个段落（默认存在）
p = tf.paragraphs[0]
p.text = "这是第一段"
p.alignment = PP_ALIGN.LEFT  # CENTER, RIGHT, JUSTIFY

# 段落间距
p.space_before = Pt(12)
p.space_after = Pt(6)
p.line_spacing = Pt(22)

# 添加新段落
p2 = tf.add_paragraph()
p2.text = "这是第二段"

# 文本运行（Run）— 同一段落的不同格式
run = p.add_run()
run.text = "加粗的文字"
run.font.bold = True
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
```

### 字体设置（完整示例）

```python
from lxml import etree

def set_font_complete(run, name="Microsoft YaHei", size=Pt(14), 
                       bold=False, italic=False, color=None):
    """完整的字体设置，支持东亚字体"""
    # 基本属性
    run.font.name = name
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    
    # 颜色
    if color:
        run.font.color.rgb = color
    
    # 东亚字体 — 兼容 python-pptx 1.0.x 和 2.x
    try:
        rPr = run._element.get_or_add_rPr()   # python-pptx >= 2.0
    except AttributeError:
        rPr = run._r.get_or_add_rPr()          # python-pptx < 2.0
    
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    rFonts = rPr.find(f"{{{ns_a}}}rFonts")
    if rFonts is None:
        rFonts = etree.SubElement(rPr, f"{{{ns_a}}}rFonts")
    rFonts.set("{http://schemas.openxmlformats.org/win/2006}wordprocessingDrawing}eastAsia", name)
```

### 项目符号

```python
from pptx.oxml.ns import qn

# 添加项目符号
p = tf.add_paragraph()
p.text = "这是带项目符号的文本"
p.level = 0  # 缩进级别

# 设置项目符号字符
pPr = p._pPr
if pPr is None:
    pPr = p._p.get_or_add_pPr()
buChar = pPr.makeelement(qn('a:buChar'), {'char': '•'})
pPr.append(buChar)
```

## 表格操作

```python
# 添加表格
rows = 4
cols = 3
left = Inches(1)
top = Inches(2)
width = Inches(11)
height = Inches(3)
table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
table = table_shape.table

# 设置列宽
table.columns[0].width = Inches(3)
table.columns[1].width = Inches(4)
table.columns[2].width = Inches(4)

# 写入单元格
cell = table.cell(0, 0)
cell.text = "表头"

# 设置单元格格式
for paragraph in cell.text_frame.paragraphs:
    paragraph.alignment = PP_ALIGN.CENTER
    for run in paragraph.runs:
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

# 单元格填充
tcPr = cell._tc.get_or_add_tcPr()
solidFill = tcPr.makeelement(qn('a:solidFill'), {})
srgbClr = solidFill.makeelement(qn('a:srgbClr'), {'val': '1F4E79'})
solidFill.append(srgbClr)
tcPr.append(solidFill)

# 合并单元格
cell1 = table.cell(0, 0)
cell2 = table.cell(0, 1)
cell1.merge(cell2)
```

## 图片操作

```python
# 添加图片
img_left = Inches(1)
img_top = Inches(2)
img_width = Inches(4)
pic = slide.shapes.add_picture("image.png", img_left, img_top, img_width)

# 等比例缩放
from pptx.util import Inches
pic.height = Inches(3)  # 手动设置高度

# 图片裁剪
pic.crop_left = Inches(0.5)
pic.crop_top = Inches(0.2)
```

## 组合形状

```python
# 组合多个形状
shapes_to_group = [shape1, shape2, shape3]
group = slide.shapes.add_group_shape(shapes_to_group)
```

## 动画（有限支持）

```python
# python-pptx 对动画支持有限
# 需要使用 lxml 直接操作 XML
from lxml import etree
# 复杂的动画建议用 VBA 或其他工具
```

## 常用枚举值

```python
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTOSIZE

# 文本对齐
PP_ALIGN.LEFT    # 左对齐
PP_ALIGN.CENTER  # 居中
PP_ALIGN.RIGHT   # 右对齐
PP_ALIGN.JUSTIFY # 两端对齐

# 垂直锚定
MSO_ANCHOR.TOP       # 顶部对齐
MSO_ANCHOR.MIDDLE    # 垂直居中
MSO_ANCHOR.BOTTOM    # 底部对齐

# 自动调整
MSO_AUTOSIZE.NONE          # 不自动调整
MSO_AUTOSIZE.SHAPE_TO_FIT_TEXT  # 形状适应文本
MSO_AUTOSIZE.TEXT_TO_FIT_SHAPE  # 文本适应形状
```

## 版式与母版

```python
# 访问幻灯片版式
slide_layout = slide.slide_layout

# 版式中的占位符
for ph in slide.placeholders:
    print(f"Placeholder {ph.placeholder_format.idx}: {ph.name}")

# 访问占位符
body_shape = slide.placeholders[1]  # 通常是正文占位符
title_shape = slide.placeholders[0]  # 通常是标题占位符
```

## 常见陷阱

1. **缩进使用 Inches/Cm/Pt** — 不要直接用数字，用 `Inches()`, `Cm()`, `Pt()` 包装
2. **颜色使用 RGBColor** — 不是字符串，是 `RGBColor(0xRR, 0xGG, 0xBB)`
3. **字体大小用 Pt** — `run.font.size = Pt(14)`，1pt = 12700 EMU
4. **文本框默认有一个段落** — `tf.paragraphs[0]` 始终存在
5. **东亚字体需要 XML 操作** — 仅设置 `run.font.name` 不够
6. **python-pptx 1.0.x 用 `run._r`** — 不是 `run._element`
