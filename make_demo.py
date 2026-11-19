"""生成演示数据：3 个门店的销售表，供 merge / split 测试与向甲方演示用。"""
from pathlib import Path

import openpyxl

OUT = Path(__file__).parent / "样例数据"
OUT.mkdir(exist_ok=True)

DATA = {
    "门店A": [
        ("部门", "姓名", "金额"),
        ("销售部", "张三", 1200),
        ("销售部", "李四", 980),
        ("技术部", "王五", 2400),
        ("技术部", "赵六", 1750),
    ],
    "门店B": [
        ("部门", "姓名", "金额"),
        ("销售部", "钱七", 3300),
        ("技术部", "孙八", 2100),
        ("市场部", "周九", 860),
    ],
    "门店C": [
        ("部门", "姓名", "金额"),
        ("销售部", "吴十", 1500),
        ("市场部", "郑一", 4400),
        ("市场部", "王二", 1120),
    ],
}

for name, rows in DATA.items():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"
    for row in rows:
        ws.append(row)
    path = OUT / f"{name}.xlsx"
    wb.save(path)
    print(f"生成 {path.name}  ({len(rows) - 1} 行)")

print(f"\n共 {len(DATA)} 个文件 -> {OUT}")
