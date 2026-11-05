# excel-recalc-gate

[English](README.md) | 简体中文

**含公式 Excel 工作簿的交付闸门。** 在你把生成的 `.xlsx` 发出去之前，这个工具先做公式 lint，再用无头 LibreOffice 在临时副本上真实重算一遍，报告所有算出错误值的公式单元格——最终只输出一个 JSON 裁决。**只有 `success` 才许交付。**

## 它解决的问题

openpyxl 这类库写出的公式**没有缓存值**。在真正的表格引擎打开重算之前，公式格在 pandas、预览窗、PDF 转换里全是空值——有时在收件人那边也是。更麻烦的是有些公式会**静默失败**：

- 动态数组函数（`XLOOKUP`、`SORT`、`FILTER` 等）LibreOffice 根本无法求值，重算可能报 0 错误、只填左上角一格；
- 2007 年后的新函数（`TEXTJOIN`、`IFS`、`MAXIFS` 等）不写 `_xlfn.` 前缀就会失效；
- 外部链接公式（`'[1]Other'!$B$2`）一旦被重存，链接和缓存值一起丢失；
- openpyxl 写出的 `Print_Area` 打印区域定义名会让 Excel 弹出「内容有问题、是否修复」——这是把 Excel 修复后自存的副本与原文件逐项比对定位到的真凶。

用脚本或 AI Agent 生成工作簿的人，这些坑都会先坑到读者、再坑到你。这道闸门把它们拦在交付之前。

## 谁在用、用它做什么

| 角色 | 典型场景 | 价值点 |
|---|---|---|
| **财务与经营分析人员** | 用脚本拼装的测算模型、预算、汇报数据包 | 出门的工作簿不再有未计算或错误值的公式；JSON 裁决可直接进 QA 留痕 |
| **数据分析与工程师** | 定时管道产出的 `.xlsx` 报表 | 对 CI 友好的退出码＋机器可读报告：出错单元格被点名（`Sheet!A3`、`#DIV/0!`），不用开 GUI 找 |
| **AI Agent 开发者** | Agent 用 openpyxl/xlsxwriter 生成表格 | 一道确定性交付前检查，Agent 能自己跑、自己按报告修复；lint 规则正是 Agent 最容易踩的函数兼容坑 |
| **任何要发 Excel 的人** | 发给用 Excel/WPS/手机预览打开的导出文件 | 读者打开看到的是数字不是空白——因为已经有真引擎替他算过一遍了 |

## 闸门如何判定

电子表格软件会把公式结果缓存在文件里。于是工作簿可能出现「显示值与公式对不上」的情况——被其它工具改过、外链断了、保存不完整——而屏幕上看一切正常。

闸门对工作簿的副本依次做三步：

1. **预检（lint）**：先做结构检查。若工作簿带外部链接，重存会把链接及其缓存值一并剥掉，这类文件直接报告、不重写。
2. **重算**：用无头 LibreOffice 强制全部公式重新计算。
3. **回读**：读取重算后文件里的缓存值，扫描错误值（#VALUE!、#REF! 等）。

判定结果：`success`（可交付）、`errors_found`（重算后仍有公式错误）、`lint_failed`（该文件根本不该被重存）。交付流程应把非 `success` 一律当作停机信号，而不是警告。

## 安装

```bash
pip install git+https://github.com/ryangucheng2022-sudo/excel-recalc-gate.git   # 提供 recalc-check 命令
```

完整重算需要 PATH 里有 LibreOffice（soffice）；没有时预检照常运行、重算步会自报不可用。也可以直接 clone 后运行 `bin/recalc_check.py`。

## 用法

```bash
python3 bin/recalc_check.py report.xlsx
python3 bin/recalc_check.py report.xlsx --force     # 明知有外部链接仍重算（接受丢链）
python3 bin/recalc_check.py report.xlsx --timeout 120
```

输出为单个 JSON 对象：

```json
{"status": "success", "total_formulas": 128, "formulas_checked": 128, "total_errors": 0}
```

| 状态 | 含义 | 退出码 |
|---|---|---|
| `success` | lint 干净且重算后错误值为 0 | 0 |
| `errors_found` | 重算出现错误值，`error_summary` 点名单元格 | 0 |
| `lint_failed` | 禁用函数、缺 `_xlfn.` 前缀，或存在 `Print_Area`/`Print_Titles` 定义名 | 0 |
| `refused_external_links` | 有外部链接公式且未给 `--force` | 2 |
| `recalc_failed` | 检查本身跑不起来（文件缺失、LibreOffice 失败） | 2 |

源文件**永不修改**——重算在临时副本上进行。

## 配套规范

[docs/delivery-standards.md](docs/delivery-standards.md) 是与闸门配套的建模与编辑规范提炼版：函数该用哪些、假设参数怎么放、为什么绝不能保存 `data_only=True` 打开的工作簿、为什么生成的工作簿不该带打印区域。

## 持续集成

`ci/` 里有一套备好的 GitHub Actions 工具包：两个 job（单元测试；以及安装 LibreOffice 的真实重算 job，并把固定样例——干净公式、错误公式、打印区域、被禁的动态数组函数——逐一过闸门断言判定）。把 `ci/ci.yml` 复制到 `.github/workflows/ci.yml` 即可启用。

## 环境

- Python 3.9+ 与 `openpyxl`
- PATH 上有 LibreOffice（`soffice`）用于重算步骤（只有 lint 时不需要）

## 测试

```bash
python3 -m unittest discover -s tests -v
# 9 个用例：7 个 lint/裁决用例 + 2 个真实重算用例（无 soffice 时自动跳过）
```

## 版本

v1.0.0 —— 见 [CHANGELOG.md](CHANGELOG.md)。

## 许可证

MIT —— 见 [LICENSE](LICENSE)。
