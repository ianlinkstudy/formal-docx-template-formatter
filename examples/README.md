# Examples

所有示例均为虚构、脱敏材料，不包含客户、单位或个人数据。

| 文件 | 用途 |
|---|---|
| `before/01-project-report.docx` | 带标题、层级标题、合并表格、落款和日期的源正文 |
| `templates/01-style-template.docx` | 没有锚点的普通样式模板 |
| `templates/02-shell-template.docx` | 带固定封面、页眉页脚和正文锚点的母版 |
| `after/01-gb9704-project-report.docx` | 使用内置 GB 参考版生成的结果 |
| `after/02-company-shell-project-report.docx` | 使用虚构公司母版生成的结果 |
| `reports/*.format-report.json` | 输出生成的语义完整性报告 |
| `reports/*.verify-report.json` | 独立排版核验报告 |

## 重建

在仓库根目录运行：

```bash
python3 examples/build_examples.py
```

脚本会在临时目录里完成“解析 → 确认 → 命名 → 保存”后再排版，不会写入本机默认模板库。
