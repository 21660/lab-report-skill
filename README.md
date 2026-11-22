# 接单了上百份计算机实验报告后让 Claude Code 总结的实验报告 skill

把「读指导书 → 查环境 → 真跑实验 → 现场截图 → 按模板填报告 → 转 PDF 校版 → 建提交文件夹」这套重复劳动固化成一个 Claude Code skill。

不是帮你编报告，而是帮你**真的把实验做完、把真实结果写进去、最后交对东西**。

> A Claude Code skill for course lab reports: run the experiment for real, capture real screenshots, fill the school template, export to PDF for proofreading, and hand in only what the teacher asked for.

**Keywords:** Claude Code skill, lab report, experiment report, 实验报告, docx template, Word COM automation, python-docx, Windows screenshot (PrintWindow), PDF export.

## 它解决什么问题

每门课的实验报告流程都一样，但每一步都有坑：

- 模板是 `.doc`，直接改会掉编号、继承红字高亮；
- 截图要拍真实运行界面，窗口被遮挡就截到别人的；
- 图片插进模板被裁成一条缝；
- 交作业时把截图、脚本、PDF 全塞进压缩包，老师要在里面挑。

这个 skill 把这些经验写成一份可执行的流程 + 配方，Claude Code 在写报告时会自动按它走。

## 安装

### 方式一：作为插件市场安装（推荐）

在 Claude Code 里：

```
/plugin marketplace add 21660/lab-report-skill
/plugin install lab-report@lab-report-skill
```

装完后重启会话，说「写实验报告」「做实验三」「填报告模板」即可触发。

### 方式二：手动复制

```bash
git clone https://github.com/21660/lab-report-skill.git
cp -r lab-report-skill/skills/lab-report ~/.claude/skills/
```

Windows PowerShell：

```powershell
git clone https://github.com/21660/lab-report-skill.git
Copy-Item -Recurse .\lab-report-skill\skills\lab-report "$env:USERPROFILE\.claude\skills\"
```

## 里面有什么

**流程**：阶段 0 读材料 + 查环境 → 阶段 1 真跑 + 取证 → 阶段 2 写报告 → 阶段 3 交付。

**工具配方**（Windows 实测可用）：

| 场景 | 做法 |
|------|------|
| 填 Word/WPS 模板 | COM 打开副本，清高亮红字，保编号换文字，插图按 13.8cm 安全宽 |
| 从零生成中文报告 | python-docx + eastAsia 字体 + TOC 域 + 封面页码 |
| 图片被裁 / 表格丢内容 | 修固定行距、`cantSplit`、`trHeight` |
| DOCX → PDF | Word/WPS COM、弹窗看门狗、TOC 拍平、ASCII 路径 |
| 真实运行截图 | PrintWindow 离屏渲染、wt.exe + 标题哨兵 |
| 提交文件夹 | 只放报告 docx + 源代码 |

**三条底线**：

1. 不擅自填个人信息——姓名、学号、班级没给就留空。
2. 有模板就在模板上原位改，不另起文档。
3. 正文不整份粘源码，源码单独交。

**交付默认**：提交文件夹只放报告 docx 和源代码；没要求就不生成 PDF，截图只嵌在正文里。

## 关于真实性

这个 skill 的优先级是 **真实性 > 速度 > 报告美观 > 完成率**。

它要求实验必须真的跑、截图必须来自真实界面、结果必须验证；做不了的步骤停下来告诉你，而不是编一个完美结果。抄作业式的"帮我伪造一份实验报告"不在适用范围。

## 适用范围

- 配方全部基于 Windows 11 + Microsoft Office / WPS + Claude Code；macOS/Linux 上流程部分仍可用，COM 与截图部分不适用。
- 不同学校模板差别很大，填模板前先量实际版心宽度。

## 目录

```
.claude-plugin/
├── marketplace.json     # 插件市场清单
└── plugin.json          # 插件清单
skills/
└── lab-report/
    └── SKILL.md
```

## License

MIT
