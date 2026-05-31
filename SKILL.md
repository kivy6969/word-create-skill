---
name: word-create
description: >
  修改 .docx 实训报告（改姓名/学号/实训内容/心得），保留图片排版截图原封不动。
  Trigger: 用户要求"修改实训报告"、"改docx"、"改名字"、"改学号"、"改实训报告"、
  "修改word"、"修改文档"、提供学分认定报告或实训报告要改个人信息等。
  IMPORTANT: 用户要求修改任何 .docx 文件的文本内容时都应触发此 skill。
---

# WordCreate — .docx 实训报告修改

修改 Word 实训报告，保留所有图片、排版、表格结构。只改文字，不动图。

## 依赖

- Python `python-docx` 库
- PowerShell（用于写脚本文件）

## 核心原则

1. **永远操作原文件，保存到新文件名**。原文件是最后的备份。
2. **图片段是禁区**。任何包含 `<blip>` 的 paragraph 绝对不碰。
3. **段落数必须匹配**。替换多段内容时，拆成和原来一样多的段落数。
4. **中文走文件，不走命令行**。命令行传递中文必然乱码。

## 工作流程

### 第一步：深度侦查

**必须做的**——在执行任何修改之前，先完整检查文档结构：

```bash
cd "目标目录"; python -c "
import docx
from docx import Document

doc = Document(r'原文件.docx')

for ti, table in enumerate(doc.tables):
    print(f'=== Table {ti} ===')
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            xml = cell._element.xml
            drawings = cell._element.findall(
                './/{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline')
            print(f'Row{ri} Cell{ci}: paragraphs={len(cell.paragraphs)}, images={len(drawings)}')
            for pi, p in enumerate(cell.paragraphs):
                has_img = 'blip' in p._element.xml.lower()
                print(f'  P{pi}: has_image={has_img}, text_len={len(p.text)}, text=[{p.text[:80]}]')
"
```

**侦查清单**：
- 哪个 table 是信息表（姓名/学号）→ 改 run.text 替换即可
- 哪个 table 是实训操作内容 → 查哪些 cell 有图片，哪个 paragraph 有图片
- 哪个 table 是实训心得 → 数有多少个 cell 有心得文字，每个 cell 有几个 paragraph
- 标注：**图片 paragraph = 禁区，文字 paragraph = 操作区**

### 第二步：准备中文内容

中文内容写到独立的 UTF-8 txt 文件，**绝不写在 Python 脚本里**：

```powershell
$content = @'
这里是要写的中文内容……
可以有多行，但不能带中文引号""，全部用英文引号替代或用其他方式处理
'@
$content | Out-File -FilePath "e:\目标目录\_content.txt" -Encoding utf8
```

### 第三步：写修改脚本

Python 脚本通过 PowerShell Out-File 写入（避免编码问题）：

```powershell
cd "目标目录";
$pyscript = @'
# -*- coding: utf-8 -*-
import docx
from docx import Document

doc = Document(r"原文件路径")

# 读取中文内容（用 utf-8-sig 自动去掉 BOM）
with open("content_file.txt", "r", encoding="utf-8-sig") as f:
    new_text = f.read().strip()

# === 改姓名/学号：替换 run.text ===
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    if "旧名字" in run.text:
                        run.text = run.text.replace("旧名字", "新名字")
                    if "旧学号" in run.text:
                        run.text = run.text.replace("旧学号", "新学号")

# === 改图片旁边的文字：只改特定的文字 paragraph，图不动 ===
# 例如 Table 2 Row0 Cell1 有3段：P0=图, P1=空行, P2=图
# 只给 P1 加文字：
cell = doc.tables[2].rows[0].cells[1]
p1 = cell.paragraphs[1]  # 第2段是空行，在两图之间
p1.text = ""
p1.add_run(operation_text)

# === 改心得：逐段替换，保持段落数 ===
# 例如心得在 Table 3 Row0 Cell1~Cell5，每个cell有5段
row0 = doc.tables[3].rows[0]
for ci in range(1, 6):  # 遍历每个内容cell
    cell = row0.cells[ci]
    for pi in range(5):  # 逐段替换
        cell.paragraphs[pi].text = ""
        cell.paragraphs[pi].add_run(new_paragraphs[pi])

# 保存到临时文件名
doc.save(r"输出文件_temp.docx")
print("SAVED OK")
'@
$pyscript | Out-File -FilePath "_modify.py" -Encoding utf8
python _modify.py
```

### 第四步：验证输出

**修改完必须验证**，不能跳过：

```bash
python -c "
import docx
from docx import Document

doc = Document(r'输出文件.docx')

# 1. 确认图片没丢
t2 = doc.tables[2]
cell = t2.rows[0].cells[1]
img_count = sum(1 for p in cell.paragraphs if 'blip' in p._element.xml.lower())
print(f'Table2 Row0 Cell1 images: {img_count}')  # 应该和原来一样

# 2. 确认段落结构
t3 = doc.tables[3]
for ci in range(1, 6):
    pc = len(t3.rows[0].cells[ci].paragraphs)
    print(f'Table3 Row0 Cell{ci} paragraphs: {pc}')  # 应该和原来一样

# 3. 确认名字和学号
print(f'Name: [{doc.tables[0].rows[0].cells[1].text}]')
print(f'ID: [{doc.tables[0].rows[0].cells[5].text}]')
"
```

验证三关全过 → 进入下一步。

### 第五步：最终命名 + 清理

```powershell
cd "目标目录";
# 如果旧文件被占用，先让用户关掉Word
Remove-Item -Path "最终文件名.docx" -Force -ErrorAction SilentlyContinue
if (Test-Path "最终文件名.docx") {
    Write-Output "文件被占用，请关闭Word后手动删除"
} else {
    Rename-Item -Path "输出文件_temp.docx" -NewName "最终文件名.docx"
    # 清理临时文件
    Remove-Item -Path "_modify.py", "_content.txt" -Force
}
```

## 踩坑记录（为什么要这样做）

| 坑 | 原因 | 正确做法 |
|---|---|---|
| 图片消失 | `p.text = ''` 清了所有段落，包括含图片的 P0/P2 | 先侦查，只改 has_image=False 的段落 |
| 心得变一坨 | 整段替换没保持段落数 | 原5段就拆成5段逐段替换 |
| 中文乱码 | 中文经过命令行→Python 编码链断裂 | 中文存 UTF-8 txt，Python 读文件 |
| 文件保存失败 | 输出文件已被 Word 打开 | 先存临时名，验证后 rename |
| 脚本语法错误 | Write/Edit 工具写 `.py` 时编码损坏 | 用 PowerShell `Out-File -Encoding utf8` 写脚本 |

## 边界

- 本 skill 假设文档是标准 .docx 格式（非 .doc 老格式）
- 图片检测通过 XML 中的 `<blip>` 标签判断
- 如果表格结构异常（如合并单元格数量不匹配），先问用户确认再动手
- 修改完成后提醒用户打开新文件检查，确认无误后再删原文件
