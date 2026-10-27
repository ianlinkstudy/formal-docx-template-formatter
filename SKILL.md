---
name: formal-docx-template-formatter
description: 中文 DOCX 模板排版与模板库。用于按 GB/T 9704 参考规范或本机最多三套自定义 Word 模板排版，导入 DOCX/DOTX 模板并让用户确认、命名和保存，选择、归档或恢复模板；处理中文正式文档的页面、标题、正文、表格、图片和公式，输出新 DOCX 与核验报告。不负责改写、润色或补写正文。
metadata:
  short-description: 选择 GB 或自定义模板，排版并核验中文 DOCX
compatibility: Requires Python 3.10 or later, python-docx, and docxcompose. A DOCX renderer is optional for visual checks.
---

# Word 模板排版

> 鼓翼成长星球 · 出品与持续维护

核心承诺：先选择模板，再改格式；用户模板必须经过“解析 → 确认 → 命名 → 保存”后才能使用。

本分享版由「鼓翼成长星球」整理与持续维护。品牌署名只用于 Skill 文档、示例和发布材料，不写入用户排版后的正文。

以下命令均从本 Skill 根目录运行。

## 每次排版先选择模板

在处理任何排版任务前运行：

```bash
python3 scripts/template_cli.py list
```

向用户展示：

- `GB/T 9704-2012 参考版`；
- 已保存的自定义模板名称；
- 剩余自定义模板槽位；
- 没有合适模板时，可以选择导入新模板。

用户已经明确说出 GB 或某个已保存模板名称时，该选择有效。否则先问用户选哪一套，不得静默沿用上次模板。

模板选择完成、且用户提供待排 DOCX 后，运行：

```bash
python3 scripts/format_docx.py INPUT.docx OUTPUT.docx --template "模板名称"
python3 scripts/verify_docx.py OUTPUT.docx --reference INPUT.docx --template "模板名称"
```

只有 `format_docx.py` 成功、`verify_docx.py` 返回 `status = passed` 时才能交付。需要页面预览时运行 `scripts/render_preview.py`。

## 导入用户模板

导入是独立流程，不启动文档排版。详细字段见 [references/template-workflow.md](references/template-workflow.md)。

1. 运行 `template_cli.py list` 检查槽位。已有三套时先让用户选择一套归档。
2. 对用户提供的 `.docx` 或 `.dotx` 运行：

   ```bash
   python3 scripts/template_cli.py inspect TEMPLATE.docx --mode auto
   ```

3. 把返回的 `summary` 整理成中文确认摘要，至少覆盖页面、标题与正文规则、表格、模式、母版锚点、低置信度项和阻断项。
4. 用户指出解析错误时，用 `amend` 修正。修正会清空旧确认并要求重新确认。
5. 只有用户明确表示“确认解析结果”或同义且无歧义的确认后，才运行 `confirm`，并带上当前 `candidate_sha256` 和用户的确认原话。
6. 确认成功后，才询问模板名称；随后运行 `save --name`。
7. 保存成功后结束导入流程。不要在本次导入中自动拿新模板去排版；下一次排版任务再让用户选择。

上传模板、沉默、说“先看看”、要求修改或直接要求排版都不等于确认。未确认、已修改、未命名或保存失败的候选不得出现在模板选择列表。

## 强制检查点

- 三套上限由 CLI 强制；GB 不占槽位，归档模板不可选用。
- 候选确认与 `candidate_sha256` 绑定；任何修改后都必须重新确认。
- 模板名称只能在确认之后询问；名称非法、重复或槽位不足时保持未保存状态。
- 母版模式必须有 `{{BODY}}` 或经用户确认的正文替换范围。
- 首版只接受 `.docx`/`.dotx` 模板和单节输入正文；`.docm`、`.doc`、PDF 与扫描件先转换。
- 不覆盖输入文件；不把用户模板写入公开分享包。
- 正文、数字、标点、表格文字和公式不得改写。模板确认摘要和核验报告必须区分“源内容”与“母版固定内容”。

## 按需读取

- 导入、确认、命名、归档的完整状态机和命令示例：[references/template-workflow.md](references/template-workflow.md)
- Profile 字段、锚点、正文范围与修正方法：[references/template-schema.md](references/template-schema.md)
- 图片、表格、公式、母版合并和核验规则：[references/verification-and-objects.md](references/verification-and-objects.md)

运行时需要 Python 3.10 或更高版本。如果环境没有依赖，先运行 `python3 -m pip install --user -r requirements.txt`；不要用聊天中的近似排版代替 DOCX 输出。
