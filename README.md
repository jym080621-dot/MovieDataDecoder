# MovieDataDecoder v1.1.0

把 MovieData 导出的一部电影的影院日明细，转成**逐影院票房 Excel**。每家已知影院一行，按已确认的该片票房从高到低排列，金额单位为元。

## 怎么用

1. 在 MovieData 的「全国全影院历史票房采集」中选好电影与日期，采集后点「导出当前 CSV / 审计」。
2. 解压 Decoder Windows 免安装包，双击 `MovieDataDecoder.exe`。选择 `*_cinema_days.csv`，也可以选择包含该文件的 ZIP 或结果文件夹。
3. 点「导出逐影院票房」，保存并打开 Excel。

## 两列金额怎么读

| 列 | 含义 |
| --- | --- |
| 已确认票房（元） | 该影院所有已核查日期的该片票房相加。至少有一天拿到有效片目时才填 |
| 完整区间票房（元） | 仅当该影院在文件名声明的区间内每天都已核查，才写整段区间的金额 |
| 待核查天数 | 区间内尚未确认的天数，包括 204、缺行、未采集及异常。204 另有一列计数 |

例如，两天里某家影院第 1 天已确认 350 元，第 2 天返回 204，第一列金额显示 350 元，完整区间票房留空。每天都取得有效片目且确认没有该片时，完整区间票房为 0 元。重命名过的明细文件如果缺少原始日期区间，会用文件内日期显示已确认金额，完整区间票房留空。

所有源文件里出现的影院都会保留，影院编码按文本保存以保留前导零。同一影院日的重复结果只计一次，金额或状态冲突时会报错。源数据的全国名单覆盖仍待核实。

## 源码运行

需要 Python 3.10+ 和 Tkinter。

```text
python -m pip install -r requirements.txt
python app.py
python decoder.py 结果.zip --list
python decoder.py 结果.zip --select 1 -o 逐影院票房.xlsx
```

离线测试需安装 `requirements-test.txt`，然后运行 `python -m unittest discover -s tests -v`。Windows 构建需安装 PyInstaller，运行 `python -m PyInstaller --clean --noconfirm --onefile --windowed --name MovieDataDecoder app.py`。
