# tools/：模型手工具脚本（仅 modeler 分支）

| 脚本 | 作用 |
|---|---|
| `fix_template_labels.py` | 把官方结果模板的 144 个 10 分钟标签重写为右端点版本（只改文字、不动结构），用于生成 `templates/`。 |

```bash
python tools/fix_template_labels.py --src "D:/数模" --dst templates
```

参照解脚本在 `src/q1/reference_check.py`（数据基准 + 独立求解 + 锚点核对）。
