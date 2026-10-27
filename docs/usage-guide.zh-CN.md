# Word 模板排版使用手册

> 鼓翼成长星球 · 出品与持续维护

## 1. 这个 Skill 解决什么问题

它把“按单位模板排版”拆成两条独立流程：

```text
导入模板：解析 → 确认 → 命名 → 保存
排版正文：选择模板 → 生成 → 核验 → 交付
```

模板确认与正文排版不会在同一次导入流程中自动接力。

## 2. 安装

需要 Python 3.10 或更高版本。

```bash
npx skills add ianlinkstudy/formal-docx-template-formatter
python3 -m pip install -r requirements.txt
```

如果卸载后重装，本地模板库默认仍位于 `~/.formal-docx-template-formatter/`。

## 3. 第一次检查模板

```bash
python3 scripts/template_cli.py list
```

返回内容包括：

- 内置 `GB/T 9704-2012 参考版`；
- 已保存的自定义模板；
- 当前有效自定义模板数量；
- 剩余槽位；
- 是否必须选择模板。

没有合适模板时先导入。不要用一个“看起来差不多”的内置版式替代单位模板。

## 4. 直接使用 GB 参考版

```bash
python3 scripts/format_docx.py 输入.docx 输出.docx \
  --template "GB/T 9704-2012 参考版"
```

成功后检查：

```text
输出.docx.format-report.json       content_integrity.passed = true
输出.docx.verify-report.json       status = passed
```

再次独立核验：

```bash
python3 scripts/verify_docx.py 输出.docx \
  --reference 输入.docx \
  --template "GB/T 9704-2012 参考版"
```

## 5. 导入普通样式模板

```bash
python3 scripts/template_cli.py inspect 单位模板.docx --mode auto
```

查看返回的 `candidate.summary`，重点确认：

- 页面尺寸和四边页边距；
- 标题、正文和分级标题；
- 表格样式；
- 自动判断的角色是否合理；
- 低置信度项；
- 阻断项。

用户明确确认后：

```bash
python3 scripts/template_cli.py confirm CANDIDATE_ID \
  --candidate-sha256 SHA256 \
  --user-confirmation-text "确认解析结果"
```

确认成功后：

```bash
python3 scripts/template_cli.py save CANDIDATE_ID \
  --name "单位年度报告"
```

保存后，本轮导入结束。下一次排版任务重新运行 `list`，再选择“单位年度报告”。

## 6. 导入母版模板

推荐在模板正文位置放一个唯一的：

```text
{{BODY}}
```

还可以放：

```text
{{TITLE}} / {{标题}}
{{RECIPIENT}} / {{主送机关}}
{{SIGNATURE}} / {{落款}}
{{DATE}} / {{日期}}
```

未知锚点、重复锚点或页眉表格里的锚点都会阻止确认。

如果模板没有锚点，可以从摘要的 `blocks` 中选择连续正文样例范围：

```bash
python3 scripts/template_cli.py amend CANDIDATE_ID \
  --mode shell --body-start 3 --body-end 12
```

修正会清空旧确认并生成新哈希。

## 7. 管理本地模板

```bash
python3 scripts/template_cli.py show "单位年度报告"
python3 scripts/template_cli.py rename "单位年度报告" "年度报告正式版"
python3 scripts/template_cli.py archive "旧版模板"
python3 scripts/template_cli.py restore "旧版模板"
```

有效自定义模板最多 3 套。归档模板不占槽位，但不能直接用于排版。

## 8. 视觉预览

```bash
python3 scripts/render_preview.py 输出.docx preview/
```

脚本优先使用带本机字体目录的 LibreOffice + poppler；macOS 没有这些工具时会退回 Quick Look 首页预览。

如果渲染器不能正确显示中文字体，保留 DOCX 中的正确字体名称，并请用户在装有相应字体的 Word/WPS 中检查。不要为了预览好看而改掉正式字体。

## 9. 失败时怎么处理

| 现象 | 处理 |
|---|---|
| 文件损坏、加密或打不开 | 解锁或重新导出 DOCX，不制造替代文件 |
| `.doc`、`.docm`、PDF 或扫描件 | 先转换为可编辑 DOCX，并核对转换结果 |
| 三套模板已满 | 先归档一套旧模板 |
| 模板名称重复 | 换名，不要静默覆盖 |
| 锚点阻断 | 修正模板或显式选择正文范围 |
| 内容完整性失败 | 停止交付，定位段落、表格或富对象差异 |
| 字体缺失 | 保留配置字体名并报告替换风险 |

## 10. 内容边界

Skill 只改变格式。不会：

- 改写、缩短、润色、翻译或补全文字；
- 统一数字、日期或标点；
- 修改组织名称或术语；
- 新增收件人、落款、红头、页码或附件；
- 自动修复正文里的逻辑问题。

文种、结构或标题层级仍不清楚时，应先询问用户。
