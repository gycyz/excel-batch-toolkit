# Excel 批量工具箱

把重复的 Excel 手工活变成点几下鼠标。适合：多个表格要合并、一张表要按部门/班级拆开、批量转 CSV。

## 两种用法

| 方式 | 怎么用 | 适合谁 |
|---|---|---|
| **窗口版**（推荐） | 双击 `run_gui.bat`，或运行 `python gui.py` | 甲方、不懂命令的人 |
| 命令行版 | `python main.py --help` | 自己批量处理、写脚本 |

窗口版有三个页签：**合并多个文件 / 按列拆分 / 批量转 CSV**，选完目录点一下就行，底部有运行日志。
拆分的列名会自动从表头读取并填进下拉框，不用手打。

## 三个功能

| 命令 | 作用 | 对应甲方常见需求 |
|---|---|---|
| `merge` | 合并目录下所有 xlsx | "把这 30 个门店的月报合成一张总表" |
| `split` | 按某列拆成多个文件 | "按部门拆开，发给各部门负责人" |
| `to-csv` | 批量转 CSV | "系统只认 csv，帮我转一下" |

## 用法

```powershell
# 合并
python main.py merge .\样例数据

# 合并（含子目录，自定义输出名）
python main.py merge .\样例数据 --recursive --out 总表.xlsx

# 按「部门」列拆分
python main.py split 总表.xlsx --by 部门

# 批量转 CSV
python main.py to-csv .\样例数据
```

## 打包成 exe 交付

**本机实测结论（2026-11-17）**：`--onefile` 打出来的 exe 运行时报 `Failed to create temporary directory!`，改用 `--onedir` 正常。

```powershell
# 推荐：onedir（本机可用，启动快，不依赖临时目录可写）
pyinstaller --onedir --name ExcelTool main.py
# 产物：dist\ExcelTool\ExcelTool.exe（整个 ExcelTool 文件夹一起发给甲方）

# onefile（本机打包成功但运行失败，换电脑可再试）
pyinstaller --onefile --name ExcelTool main.py
```

**注意**：
- exe 可能被 Windows Defender / 360 误报为病毒，交付时提前说明一句，或者同时给源码包。
- 交付前一定在**普通桌面 PowerShell 里双击/运行一次**，不要只在开发环境里测。
- CLI 工具对不懂电脑的甲方不友好，真要交付可以再加个 tkinter 小窗口（不增加依赖）。

## 接单话术

> 我可以帮您把这个流程做成一个双击就能用的小工具：把文件丢进文件夹，点一下，自动合并/拆分，以后每个月都能用，不用再手动复制粘贴。价格 XXX 元，含一次修改。

## 下一步可以加的功能（加价点）

- 支持 `.xls` 老格式（需装 `xlrd`）
- 按多个条件拆分
- 合并时自动加一列「来源文件名」
- 做成带界面的窗口（`tkinter`，不增加依赖）
