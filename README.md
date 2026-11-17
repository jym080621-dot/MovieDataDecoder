# MovieDataDecoder v1.0.0

MovieDataDecoder 将 MovieData 的导出结果转换成容易查阅的 Excel 工作簿。它独立于采集器运行，全程离线。

## 三步使用

1. 解压 Windows 免安装包，双击 `MovieDataDecoder.exe`。
2. 点「选择文件」或「选择文件夹」，选择 MovieData 导出的 CSV、包含结果的 ZIP、逐日汇总 JSON，或 `output` 文件夹。发现多份结果时，从下拉菜单选一份。
3. 点「生成 Excel」，选择保存位置。完成后点「打开 Excel」。

要看逐影院的票房，在 MovieData 的全国历史页面先点「导出当前 CSV / 审计」，再选择 `*_cinema_days.csv`。`*_daily_audit.csv` 和 `*_summary.json` 适合只看每日核查进度。普通 MovieData 结果 CSV 也可直接转成便于筛选的数据表。

## 怎样读 Excel

| 工作表 | 内容 |
| --- | --- |
| 结果概览 | 影院日数量、已经核查的数量与合计金额、当前数据范围和来源 |
| 每日汇总 | 每天已知候选数、已核查数、有片、无片、204、未采、异常的数量及当前可确认票房 |
| 影院明细 | 每家影院每日的核查结果和票房。可按日期、影院编码或状态筛选 |
| 审计对照 | 同目录或 ZIP 中有同名前缀的逐日审计 CSV 时，按日对比。发现差异可检查两份导出的时间 |

影院编码和电影编码按文本保留前导零。有效片目中没有目标片时金额为 0。HTTP 204、未采集、异常和其他未知状态的金额留空。已核查票房只汇总有效片目，历史全国完整性仍以来源证据为准。逐日审计和汇总 JSON 本身没有票房明细，金额栏留空。

勾选「明细只显示含目标电影的影院日」可缩短明细工作表，每日汇总仍统计全部状态。长度超过 Excel 单表上限时，程序自动分出续表。

## 源码运行

需要 Python 3.10+ 和 Tkinter。在源码目录执行：

```text
python -m pip install -r requirements.txt
python app.py
python decoder.py --help
```

命令行可用 `python decoder.py 结果.zip --list` 查看可转换项，再用 `--select 2 -o report.xlsx` 生成指定结果。离线测试需安装 `requirements-test.txt` 后运行 `python -m unittest discover -s tests -v`。

在 Windows 安装 PyInstaller 后可构建免安装版：

```text
python -m PyInstaller --clean --noconfirm --onefile --windowed --name MovieDataDecoder app.py
```
