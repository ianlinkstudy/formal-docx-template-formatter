# 模板生命周期

## 选择规则

每次新的排版任务都从 `template_cli.py list` 开始。有效选择只有：

1. 内置 `gb9704`；
2. `custom` 中的已保存模板名称。

`pending` 候选不是模板，不能传入 `format_docx.py`。没有合适模板时，先完成导入流程。

## 导入状态

### 1. 解析

```bash
python3 scripts/template_cli.py inspect TEMPLATE.docx --mode auto
```

`auto` 发现已知锚点时进入母版模式，否则进入样式模式。也可显式使用 `--mode style` 或 `--mode shell`。

返回的 `candidate.summary` 是给用户确认的唯一摘要来源。不要自行猜测低置信度规则，也不要把完整 JSON 原样倾倒给普通用户。

### 2. 修正

修改 profile：

```bash
python3 scripts/template_cli.py amend CANDIDATE_ID \
  --changes-json '{"profile":{"styles":{"body":{"size_pt":12,"line_rule":"multiple","line_value":1.5}}}}'
```

修改母版正文范围：

```bash
python3 scripts/template_cli.py amend CANDIDATE_ID \
  --mode shell --body-start 3 --body-end 12
```

`body-start` 和 `body-end` 使用 `summary` 中 `blocks` 的零基索引，范围包含结束块。任何 `amend` 都会清空旧确认并生成新哈希。

### 3. 明确确认

用户明确确认后：

```bash
python3 scripts/template_cli.py confirm CANDIDATE_ID \
  --candidate-sha256 SHA256 \
  --user-confirmation-text "确认解析结果"
```

如果 CLI 报告哈希不匹配或存在阻断项，重新展示最新摘要，不能确认旧版本。

### 4. 命名并保存

确认成功后才询问名称：

```bash
python3 scripts/template_cli.py save CANDIDATE_ID --name "集团年度报告"
```

名称去除首尾空格，最多 32 个字符，不能与当前有效模板重名。

## 三槽与归档

```bash
python3 scripts/template_cli.py archive "旧模板"
python3 scripts/template_cli.py restore "旧模板"
python3 scripts/template_cli.py rename "旧名称" "新名称"
python3 scripts/template_cli.py show "模板名称"
```

三套有效模板之外还可以保存归档记录；归档不占槽位，但不能用于排版。更新模板时先归档旧版，再从原始 DOCX 重新解析、确认和命名。
