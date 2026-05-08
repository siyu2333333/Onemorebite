#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
========================================================
  朴朴商品数据解析 - 本地运行入口
  run_parser.py
========================================================

【目录结构说明】
  你的项目文件夹建议如下组织：

  my_project/
  ├── run_parser.py         ← 本文件（运行入口）
  ├── pupu_parser.py        ← 解析器核心（从myGPT下载）
  ├── input/                ← 放原始 JSON/txt 数据文件
  │   ├── 朴朴-牛肉片源数据.txt
  │   ├── 朴朴-猪肉片数据.txt
  │   └── ...
  └── output/               ← 解析结果 Excel 输出到这里
      └── 朴朴_解析结果.xlsx

【使用前准备】
  1. 安装依赖（在终端执行一次）：
     pip install pandas openpyxl

  2. 把 pupu_parser.py 和本文件放在同一个文件夹

  3. 修改下方"配置区"的路径，然后运行：
     python run_parser.py
========================================================
"""

import sys
import os
from pathlib import Path
import pandas as pd

# ────────────────────────────────────────────────────────
# ★ 配置区：修改这里的路径即可，其他不用动
# ────────────────────────────────────────────────────────

# 【模式1】单文件解析 —— 填入你的 JSON/txt 文件路径
SINGLE_FILE = r"C:\Users\你的用户名\Desktop\my_project\input\朴朴-牛肉片源数据.txt"
# macOS / Linux 示例：
# SINGLE_FILE = "/Users/yourname/Desktop/my_project/input/朴朴-牛肉片源数据.txt"

# 【模式2】批量文件夹解析 —— 填入存放所有 JSON 文件的文件夹路径
BATCH_FOLDER = r"C:\Users\你的用户名\Desktop\my_project\input"
# macOS / Linux 示例：
# BATCH_FOLDER = "/Users/yourname/Desktop/my_project/input"

# 输出 Excel 保存路径（文件夹必须存在）
OUTPUT_FILE = r"C:\Users\你的用户名\Desktop\my_project\output\朴朴_解析结果.xlsx"
# macOS / Linux 示例：
# OUTPUT_FILE = "/Users/yourname/Desktop/my_project/output/朴朴_解析结果.xlsx"

# 平台名称标签（会写入"竞品渠道"列）
PLATFORM = "朴朴"

# 运行模式：填 "single"（单文件）或 "batch"（批量文件夹）
RUN_MODE = "single"

# ────────────────────────────────────────────────────────
# 以下代码不需要修改
# ────────────────────────────────────────────────────────

def check_env():
    """运行前检查：文件是否存在、依赖是否安装"""
    print("\n🔍 环境检查...")

    # 检查 pupu_parser.py 是否在同目录
    parser_path = Path(__file__).parent / "pupu_parser.py"
    if not parser_path.exists():
        print(f"  ❌ 找不到 pupu_parser.py，请把它放到：{parser_path.parent}")
        sys.exit(1)
    print(f"  ✅ pupu_parser.py 已找到")

    # 检查依赖
    try:
        import pandas
        import openpyxl
        print(f"  ✅ pandas {pandas.__version__} 已安装")
    except ImportError as e:
        print(f"  ❌ 缺少依赖：{e}")
        print("  请在终端运行：pip install pandas openpyxl")
        sys.exit(1)

    # 检查输出目录是否存在
    output_dir = Path(OUTPUT_FILE).parent
    if not output_dir.exists():
        print(f"  📁 输出目录不存在，自动创建：{output_dir}")
        output_dir.mkdir(parents=True, exist_ok=True)
    print(f"  ✅ 输出目录：{output_dir}")


def run_single():
    """单文件解析模式"""
    from pupu_parser import parse_file

    fp = Path(SINGLE_FILE)
    if not fp.exists():
        print(f"\n❌ 文件不存在：{SINGLE_FILE}")
        print("   请检查路径是否正确，注意 Windows 路径用 r\"...\" 或双反斜杠")
        sys.exit(1)

    print(f"\n📂 正在解析：{fp.name}")
    result = parse_file(str(fp), platform=PLATFORM)

    # 打印关键字段预览
    print("\n" + "=" * 55)
    print("✅ 解析结果预览")
    print("=" * 55)
    preview_fields = [
        "竞品渠道", "类似产品品名", "品牌", "规格",
        "售价", "折合每kg(元)", "冷冻/冰鲜", "国别",
        "产地", "保质期", "营销卖点", "评论数",
    ]
    for k in preview_fields:
        v = str(result.get(k, ""))
        v_display = v[:50] + "…" if len(v) > 50 else v
        print(f"  {k:<15}: {v_display}")

    # 保存 Excel
    df = pd.DataFrame([result])
    df.to_excel(_sanitize_filename(OUTPUT_FILE), index=False)
    print(f"\n📁 Excel 已保存到：{OUTPUT_FILE}")
    return df


def run_batch():
    """批量文件夹解析模式"""
    from pupu_parser import parse_folder

    folder = Path(BATCH_FOLDER)
    if not folder.exists():
        print(f"\n❌ 文件夹不存在：{BATCH_FOLDER}")
        sys.exit(1)

    # 统计文件数量
    txt_files  = list(folder.glob("*.txt"))
    json_files = list(folder.glob("*.json"))
    total = len(txt_files) + len(json_files)

    if total == 0:
        print(f"\n⚠️  文件夹中没有 .txt 或 .json 文件：{BATCH_FOLDER}")
        sys.exit(1)

    print(f"\n📂 批量解析文件夹：{folder}")
    print(f"   找到 {len(txt_files)} 个 .txt 文件，{len(json_files)} 个 .json 文件，共 {total} 个")

    df = parse_folder(str(folder), platform=PLATFORM)

    if df.empty:
        print("\n⚠️  没有成功解析任何文件，请检查文件格式")
        sys.exit(1)

    print(f"\n✅ 成功解析 {len(df)} 条记录")
    print(df[["竞品渠道", "类似产品品名", "规格", "售价", "折合每kg(元)", "冷冻/冰鲜", "国别"]].to_string(index=False))

    # 保存 Excel
    df.to_excel(_sanitize_filename(OUTPUT_FILE), index=False)
    print(f"\n📁 Excel 已保存到：{OUTPUT_FILE}")
    return df


# ────────────────────────────────────────────────────────
# 主入口
# ────────────────────────────────────────────────────────
if __name__ == "__main__":
    check_env()

    if RUN_MODE == "single":
        print("\n🚀 运行模式：单文件解析")
        run_single()
    elif RUN_MODE == "batch":
        print("\n🚀 运行模式：批量文件夹解析")
        run_batch()
    else:
        print(f"\n❌ 未知的 RUN_MODE: {RUN_MODE}，请填写 \'single\' 或 \'batch\'")
        sys.exit(1)

    print("\n🎉 完成！")
