# 贡献指南

## 核心边界

这个 Skill 的第一原则是：

> 先选择并确认模板，再改变格式；正文内容保持不变。

任何改动都必须优先保证段落文字、数字、标点、表格、页眉页脚、图片、公式和 OLE 对象不变。

## 开发环境

需要 Python 3.10 或更高版本。

```bash
python3 -m pip install -r requirements.txt
python3 -m unittest discover -s tests -v
```

涉及模板生命周期时，还要用一套匿名 DOCX 完成：

```bash
python3 scripts/template_cli.py inspect 匿名模板.docx --mode auto
python3 scripts/format_docx.py 输入.docx 输出.docx --template "匿名模板"
python3 scripts/verify_docx.py 输出.docx --reference 输入.docx --template "匿名模板"
```

## 不可破坏的状态机

- 未确认候选不能保存；
- 任何 `amend` 都必须使旧确认失效；
- 确认必须绑定当前 `candidate_sha256`；
- 保存成功前不能询问或使用模板名称；
- 保存成功后，本轮导入流程立即结束；
- 每次新排版任务都必须先展示并选择模板。

## 提交内容

请勿提交：

- 未脱敏的单位模板或正文；
- 个人身份信息；
- 商业字体文件；
- 作者、修订记录、批注或隐藏元数据未清理的真实材料；
- 本地模板库、运行缓存和 release 压缩包。

新增示例必须使用虚构内容，并保留源文件、输出文件、排版报告和核验报告。

## 版本规则

- 修复 bug：补丁版本号；
- 向后兼容的新能力：次版本号；
- 技术名、模板状态机或核心行为变化：主版本号。

每次发布同步更新 `VERSION`、`CHANGELOG.md`、`docs/iterations.md` 和对应发布说明。
