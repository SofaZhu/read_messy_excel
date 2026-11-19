# read-messy-excel

把有标题行、多级表头和合并单元格的 `.xlsx` 报表整理成 pandas DataFrame。自动识别表头和末尾汇总行，返回带原始 Excel 坐标的转换报告，也可以在本地浏览器预览结果。

Flatten messy Excel sheets into pandas DataFrames with explainable parsing and a local preview.

## 安装

需要 Python 3.10 或更新版本。当前开发版可从仓库安装。

```sh
git clone https://github.com/SofaZhu/read_messy_excel.git
cd read_messy_excel
python -m pip install .
```

如果系统用 `python3` 启动 Python 3，请在下文中用 `python3` 替换 `python`。

## 使用

### Python

```python
from messy_excel import flatten

df, report = flatten("sales.xlsx")
print(df.head())
print(report.header_rows)          # 例如 (2, 3)
print(report.filled_merge_ranges)  # 例如 ('B2:C2', 'A4:A5')
print(report.dropped_tail_rows)    # 例如 (7,)
```

表头证据不足时，`flatten()` 抛出 `HeaderDetectionError`。可以从异常的 `report.candidates` 和 `report.warnings` 查看候选和原因，再指定表头重试。

```python
df, report = flatten(
    "sales.xlsx",
    sheet_name="销售明细",
    header_row=2,
    header_depth=2,
    drop_tail=0,
)
```

| 参数 | 含义 |
| --- | --- |
| `sheet_name` | 工作表名称或从 `0` 开始的索引。默认读取第一个工作表。 |
| `header_row` | Excel 中从 `1` 开始的表头起始行。省略时自动识别。 |
| `header_depth` | 连续表头行数。指定 `header_row` 后默认用一行。 |
| `drop_tail` | 默认 `None`，自动识别连续的末尾汇总行。`0` 保留非空尾行，正整数明确删除指定数量的尾行。 |

[快速示例](examples/quickstart.py)会生成一份临时销售周报并打印结果，运行 `python examples/quickstart.py` 即可复现。

### 命令行

`inspect` 输出候选表头、评分、合并范围和尾行处理。表头无法可靠判定时会输出原因并以非零状态退出。

```sh
messy-excel inspect sales.xlsx
messy-excel inspect sales.xlsx --sheet-name "销售明细" --header-row 2 --header-depth 2 --drop-tail 0
```

如果 `messy-excel` 没有加入 `PATH`，可用 `python -m messy_excel inspect sales.xlsx`。

### 本地工作台

```sh
python -m messy_excel gui
```

命令会启动本地工作台并打开浏览器。可以先用内置销售周报试用，也可以选择或拖入自己的 `.xlsx`。界面支持切换工作表、按原 Excel 行号预览、查看解析报告，以及调整参数后重试。工作簿由本机处理，原文件不会被修改。用 `Ctrl+C` 停止工作台。

## 读取范围

- 每次读取一个 `.xlsx` 工作表。自动表头在前 30 行中寻找连续 1–3 行。短表、纯文本表和结构不明显的多级表头可能需要手动指定。
- 表头按 `父级 / 子级` 拼成列名。数据区只展开单列纵向合并，普通空单元格保留为空值。
- 自动去尾只考虑连续的末尾候选行，综合汇总标签、数字和明细字段判断。不确定的候选会保留并报告原因。可用 `drop_tail=0` 保留全部非空尾行。
- 公式读取已有缓存值。没有缓存时读为空。读取范围上限为 100 万格，本地工作台接受 32 MB 以内的文件，预览前 50 行与前 30 列。

## 开发

```sh
python -m pip install -e .
python -m unittest discover -s tests -v
```

反馈识别问题时，可以附上一份可公开分享的脱敏工作簿、预期表头和 `messy-excel inspect` 的输出。

## 许可证

[MIT](LICENSE)
