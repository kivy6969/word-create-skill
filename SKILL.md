---
name: word-create
description: >
  修改 .docx 文档的文本内容（字段替换、段落重写、内容补充），完整保留图片、
  排版、表格结构。Trigger: 用户要求修改 Word 文档、替换文档中的文本、更新报告内容、
  修改 .docx 文件、批量处理文档、填充模板等任何涉及 .docx 文本修改的请求。
---

# WordCreate — .docx 结构化文档修改

对 .docx 文档进行精确的文本修改，完整保留所有图片、排版与表格结构。
适用于报告模板填充、批量文档更新、表格内容替换等场景。

## 依赖

- Python `python-docx`
- PowerShell

## 核心原则

1. **始终从原文件读取，输出到新文件**。永远不会覆盖源文件。
2. **图片段落是只读区域**。任何包含 `<blip>` 元素的 paragraph 不作任何修改。
3. **段落结构必须保持一致**。替换多段文本时，输出段落数须与输入段落数完全相同。
4. **中文字符串通过文件传递，不经过命令行**。防止编码链损坏。

## 工作流

### 第一步：文档侦查

修改之前必须完整扫描文档结构。获取每一张表、每一个单元格的段落数以及图片分布情况：

```bash
python -c "
import docx
from docx import Document

doc = Document(r'input.docx')

for ti, table in enumerate(doc.tables):
    print(f'=== Table {ti} ===')
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            ns = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
            drawings = cell._element.findall(f'.//{ns}inline')
            print(f'Row{ri} Cell{ci}: paragraphs={len(cell.paragraphs)}, images={len(drawings)}')
            for pi, p in enumerate(cell.paragraphs):
                has_image = 'blip' in p._element.xml.lower()
                text_preview = p.text[:80] if p.text.strip() else '(empty)'
                print(f'  P{pi}: image={has_image}, len={len(p.text)}, [{text_preview}]')
"
```

侦查结束后应明确：
- 哪些单元格需要修改文本
- 哪些段落包含图片（标记为不可修改）
- 每个目标区域当前的段落数量

### 第二步：准备文本内容

所有中文或其他非 ASCII 文本写入独立的 UTF-8 文件，禁止直接嵌入 Python 源码：

```powershell
$content = @'
替换文本内容...
'@
$content | Out-File -FilePath "content.txt" -Encoding utf8
```

### 第三步：执行修改

Python 脚本通过 PowerShell 的 `Out-File -Encoding utf8` 写入，以避免工具链编码问题：

```powershell
$pyscript = @'
# -*- coding: utf-8 -*-
from docx import Document

doc = Document(r"input.docx")

with open("content.txt", "r", encoding="utf-8-sig") as f:
    new_text = f.read().strip()

# 示例 A：替换表格中的字段值（如公司名称、日期、编号）
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    if "旧值" in run.text:
                        run.text = run.text.replace("旧值", "新值")

# 示例 B：在图片之间的空白段落插入文字说明
# 某单元格有 3 段：P0=图片, P1=空白, P2=图片
# 仅操作 P1，P0 和 P2 保持不动
cell = doc.tables[2].rows[0].cells[1]
middle_paragraph = cell.paragraphs[1]
middle_paragraph.text = ""
middle_paragraph.add_run(description_text)

# 示例 C：逐段替换多段文本（保持段落数不变）
# 某区域原有 N 段文本，替换内容也拆分为恰好 N 段
target_cell = doc.tables[3].rows[0].cells[1]
for i, paragraph_text in enumerate(replacement_paragraphs):
    target_cell.paragraphs[i].text = ""
    target_cell.paragraphs[i].add_run(paragraph_text)

doc.save(r"output_temp.docx")
print("done")
'@
$pyscript | Out-File -FilePath "_modify.py" -Encoding utf8
python _modify.py
```

### 第四步：验证

修改完成后必须对输出文件进行验证：

```bash
python -c "
from docx import Document

doc = Document(r'output_temp.docx')

# 1. 图片数量校验（与侦查结果对比）
for ti in [2]:  # 包含图片的表索引
    table = doc.tables[ti]
    for ri, ci in [(0, 1)]:  # 包含图片的单元格坐标
        cell = table.rows[ri].cells[ci]
        img_count = sum(1 for p in cell.paragraphs if 'blip' in p._element.xml.lower())
        print(f'Table{ti} Row{ri} Cell{ci} images: {img_count}')

# 2. 段落结构校验（与侦查结果对比）
table = doc.tables[3]
for ci in range(1, 6):
    pc = len(table.rows[0].cells[ci].paragraphs)
    print(f'Table3 Row0 Cell{ci} paragraphs: {pc}')

# 3. 关键字段值校验
print(f'Field value: [{doc.tables[0].rows[0].cells[1].text}]')
"
```

三项全部通过 → 进入最终命名。

### 第五步：命名与清理

```powershell
# 删除可能被锁定的旧文件（用户需先关闭 Word）
Remove-Item -Path "output.docx" -Force -ErrorAction SilentlyContinue
if (Test-Path "output.docx") {
    Write-Output "文件被占用，请关闭 Word 后重试"
} else {
    Rename-Item -Path "output_temp.docx" -NewName "output.docx"
    Remove-Item -Path "_modify.py", "content.txt" -Force
}
```

## 常见问题与原因

| 问题 | 原因 | 预防措施 |
|---|---|---|
| 图片消失 | 清理文本时误清除了包含图片的段落 | 侦查阶段标记所有 `has_image=True` 的段落，修改时跳过 |
| 段落合并为一段 | 将多段替换文本写入单个段落 | 保持段落数量一一对应，逐个替换 |
| 中文输出乱码 | 中文字符串经命令行传递时编码丢失 | 通过 UTF-8 文件传递所有非 ASCII 文本 |
| 保存失败 | 输出文件正被 Word 或其他进程占用 | 先写临时文件，验证后再改名 |
| Python 脚本语法错误 | 源码文件经某些工具写入时编码损坏 | 始终用 PowerShell `Out-File -Encoding utf8` 写入 .py 文件 |
