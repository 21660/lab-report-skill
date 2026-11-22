---
name: lab-report
description: Use when writing a course lab report end to end (实验报告 / lab report / experiment report / fill the report template): check environment, run the experiment for real, take real screenshots, fill the school Word template in place, export PDF to proofread, build the submission folder. Includes Word/WPS COM, python-docx, PrintWindow screenshot and DOCX→PDF recipes. 当用户说"写实验报告""做实验X""填报告模板""交实验"时使用。不用于 ML 研究结果汇报（那走 experiment-report）。
---

# 课程实验报告

优先级：**真实性 > 速度 > 报告美观 > 完成率**。宁可说某步做不了，也不伪造成功。

## 三条硬规则

1. **不擅自填个人信息** —— 姓名、学号、班级、指导教师、日期等，没给就留空或回问。
2. **有模板就在模板文件上原位改** —— 不另起新文档，不改模板结构，不缺栏目。
3. **正文不整份粘源码** —— 源码作为文件单独交，正文用"逐段解析 + 行内关键片段"。
4. **交付物只放老师要的** —— 默认提交文件夹 = 报告 docx + 源代码，两样。**没要求就不生成 PDF**；截图、脚本、日志、中间产物、README 一律不进提交文件夹。

## 流程

**阶段 0 读材料 + 查环境**
找齐指导书、任务书、模板、素材、以前的报告（格式参照）。查本机真实环境并实际启动验证；缺组件/权限/版本不兼容 → 立即停下报告，不擅自改系统环境。先汇报材料+环境+阻塞点，再开工。

**阶段 1 真跑 + 取证**
每步真执行，检查输出和最终状态；不报错 ≠ 成功（进程真起来？数据真变了？文件内容对？）。截图用真实界面、步骤完成后立刻截、语义化编号（`01_环境检查.png`），存**工作目录**（`shots/`），不进提交文件夹。遇问题立即报告：①哪步 ②什么错 ③试过什么 ④现状 ⑤需要用户做什么。

**阶段 2 写报告**
按模板章节顺序填，模板表格/图片/分工表保留；步骤=实际执行的，结果=真跑的，分析结合实际。三种状态分开写：**【已完成】/【未完成】(原因) /【待确认】(需要用户提供什么)**。与指导书有差异要说明原因。语言像普通本科生写的。

**阶段 3 交付（收尾必做）**

报告写完 ≠ 任务完成。建独立提交文件夹（如 `提交/` 或 `最终提交/`），**不压缩**。

| 放进去 | 不放进去 |
|--------|----------|
| 最终报告 **docx** | PDF（没要求就不生成、也不放） |
| 源代码（多目录保留子结构） | 截图文件夹（截图只嵌在报告正文里） |
| 老师模板/指导书明确要求的东西 | 生成脚本、运行日志、数据/中间产物、测试文件、README |

- 从工作目录**只复制该交的文件**，不是整目录复制。
- 若指导书/模板写明了交什么，或以用户明确指示为准（如"这次把 PDF 也发我"）。
- 最后 `ls` 一遍提交夹，逐个文件问：老师会要这个吗？不会就移走。

---

## 工具调用

### 1. 填 Word/WPS 模板（保留原版式）

先复制模板再改副本；WPS 用 `KWPS.Application`，Word 用 `Word.Application`（本机两者都在）。

```python
import shutil, win32com.client
shutil.copyfile(tpl, out)
app = win32com.client.DispatchEx("KWPS.Application")
doc = app.Documents.Open(out)

rng = doc.Tables(1).Cell(3, 2).Range
rng.InsertAfter("正文内容")
rng.HighlightColorIndex = 0          # 必做：否则继承空段落的高亮
rng.Font.Color = -16777216           # 必做：否则继承红字
```

```python
# 替换模板里 "1. 2. 3." 这类自动编号段落：只换标记之前，别换段落标记
doc.Range(p.Range.Start, p.Range.End - 1).Text = "新正文"

# 插图：先量单元格可用宽度（常见学校模板 A4 版心实测安全值 13.8cm）
rng = doc.Tables(1).Cell(4, 2).Range
shp = doc.InlineShapes.AddPicture(img_path, False, True, rng)
shp.Width = 13.8 * 28.35
shp.Height = shp.Width * 原高 / 原宽
doc.Range(rng.Start, rng.End).ParagraphFormat.KeepWithNext = True
```

行距被"文档网格"锁死时（.doc 老模板常见），改 **Paragraph 对象**，不是 ParagraphFormat：

```python
p.DisableLineHeightGrid = True
p.Format.LineSpacingRule = 4      # wdLineSpaceExactly
p.Format.LineSpacing = 12.1       # 10.5pt 宋体单倍 ≈ 12.1pt/行
```

### 2. python-docx 从零生成中文报告

打开模板 → 找到"摘要"标题 → 删其后全部 body 元素（保留 `w:sectPr`）→ 追加内容。保留封面/样式/页面设置。

```python
from docx.oxml.ns import qn
from docx.shared import Pt

# 标题：用模板 Heading 1/2/3（目录靠它），但字号在 run 上显式覆盖
r = p.add_run("一、实验目的"); r.font.size = Pt(18); r.font.name = "黑体"

# 中文字体必须单独设 eastAsia；西文给 TNR（宋体的 O 会被读成 0）
def set_font(run, size, cn="宋体", en="Times New Roman"):
    run.font.size = Pt(size); run.font.name = en
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cn)
```

- 首行缩进 2 字符用 `w:firstLineChars="200"`，不要写固定 twips。
- 自动目录：插 TOC 域 ` TOC \o "1-3" \h \z \u `，并在 settings.xml 加 `<w:updateFields w:val="true"/>`；否则空白。
- 封面不编页码：`section.different_first_page_header_footer = True`，页码只放普通 footer。
- 插图宽度 = 页宽 − 左右边距（A4 常规边距取 14.5cm 安全）。

### 3. 图片/表格版面修复

```python
pic.paragraph_format.line_spacing = 1.0        # 模板 Normal 是固定值行距时，图会被裁成一条
```

```python
# 大表格模板：行带 cantSplit + trHeight 会把超页内容整段丢弃
trPr = row._tr.trPr
el = trPr.find(qn("w:cantSplit"))
if el is not None: trPr.remove(el)
trPr.find(qn("w:trHeight")).set(qn("w:hRule"), "auto")
```

表格：列宽之和 ≤ 版心（A4 约 14.66cm）；跨页加 `<w:cantSplit/>` + 表头 `<w:tblHeader/>`；固定版式要配 `w:tblLayout type=fixed` 并改 `w:gridCol/w:w`。

### 4. DOCX → PDF（校版用）

先试 Word，挂死再换 WPS：

```powershell
$app = New-Object -ComObject KWPS.Application      # 或 Word.Application
$doc = $app.Documents.Open($In, $false, $false)
$doc.Fields.Update() | Out-Null
$doc.TablesOfContents.Item(1).Update() | Out-Null
$doc.Save()
$doc.SaveAs([ref]$Out, [ref]17)                    # 17 = wdFormatPDF
$doc.Close([ref]0); $app.Quit()
```

- 导出目标先用纯 ASCII 路径，再 Copy-Item 进中文目录（中文路径会挂）。
- 卡住且无输出 → 枚举顶层窗口找"安全声明/安全警告"弹窗，PostMessage 回车。
- 带 TOC 域反复挂死 → 删掉 body 里全部 `w:fldChar`/`w:instrText`（拍平成静态文字），保留 Hyperlink 样式只改颜色。
- 失败后 `taskkill /F /IM WINWORD.EXE`（或 wps），否则残留进程锁住工作目录。
- 没有 Word/WPS 时：HTML + `msedge.exe --headless=new --print-to-pdf=out.pdf file:///...`。

**PDF 必须逐页渲染看过**：`pymupdf.open(pdf)` → `page.get_pixmap(dpi=110).save(...)`。

### 5. 截图（真实运行现场）

窗口被遮挡时不要抓屏，用 PrintWindow 直接渲染窗口（Z 序无关）：

```python
hdc = user32.GetWindowDC(hwnd); mdc = gdi32.CreateCompatibleDC(hdc)
bmp = gdi32.CreateCompatibleBitmap(hdc, w, h); gdi32.SelectObject(mdc, bmp)
user32.PrintWindow(hwnd, mdc, 2)      # PW_RENDERFULLCONTENT
# GetDIBits(biHeight=-h) → Image.frombuffer("RGB", (w,h), buf, "raw", "BGRX", 0, 1)
```

控制台程序：`wt.exe -w new --size 106,62 new-tab --title <唯一tag> cmd /k ...`，`chcp 65001`，stdout 一律 DEVNULL；用标题写哨兵 `<tag>_DONE` 判断跑完，别靠"画面两帧相同"。**绝不能按标题子串 WM_CLOSE**（会匹配到用户自己的终端）。

### 6. 提交文件夹

```powershell
New-Item -ItemType Directory 提交
Copy-Item report.docx, src -Destination 提交 -Recurse
Get-ChildItem 提交 -Recurse | Select-Object FullName
```

只复制该交的文件，不是整目录复制。**不要顺手生成 PDF**——PDF 是校版用的中间产物，除非老师要求或用户明确说要，否则不进提交夹。交付前对照上面的"放进去/不放进去"表逐项核对。

---

## 交付前自查

每项要求真跑了？有没有没跑就写的步骤？截图是这次的、和正文对得上？结果没被改预期值掩盖？未解决的事都写进报告并告知了用户？提交夹里只有该交的？

## 深坑速查（作者本机 memory 索引，可忽略；细节按名字自行搜索）

| 场景 | 记忆文件 |
|------|----------|
| 模板 COM 全部坑 / 网格行距 | `word-fill-template-com` |
| 反向填进下发表单 / 编号跨文档 | `docx-fill-teacher-form` |
| python-docx 中文报告完整配方 | `python-docx-chinese-report` |
| 表格超宽、跨页、LO 转换 | `docx-table-layout-fixes` |
| 图片被裁、公式编号、文末空白页 | `docx-inline-picture-clipped` |
| DOCX→PDF 挂死与弹窗 | `office-docx-to-pdf` |
| HTML→PDF | `html-to-pdf-edge` |
| 控制台/IDE 截图 | `windows-console-screenshot`、`intellij-run-screenshot`、`pycharm-lab-run-screenshot` |
| 具体课程环境 | MEMORY.md 里按课程名 grep（mysql / kingbase / oracle / matlab / stm32 / java-web …） |
