# -*- coding: utf-8 -*-
"""蚊媒监测数据处理工具 —— 入口。

用法：
    python main.py                      # 启动图形界面
    python main.py --cli -i 源文件.xlsx -d 2026-08-20 [-e 荔湾区] [-o 输出目录]
"""
import argparse
import os
import sys


def run_cli(args):
    from app import excel_writer, processor, word_writer

    def log(msg, step=None):
        prefix = f"[{step}/{7}] " if step else ""
        print(prefix + msg, flush=True)

    outdir = args.outdir or os.path.dirname(os.path.abspath(args.input))
    os.makedirs(outdir, exist_ok=True)

    result = processor.process_all(
        args.input, args.year, args.month, args.day,
        exclude_field=args.exclude, log=log,
    )

    p_path = excel_writer.process_excel_path(outdir, result.target)
    w_path = excel_writer.word_path(outdir, result.target)
    s_path = excel_writer.summary_excel_path(outdir, result.target)

    log("正在生成计算过程Excel（BI/ADI 全部中间表）…", 6)
    excel_writer.build_process_excel(result, p_path)

    log("正在生成日报Word…", 6)
    word_writer.build_word(result, w_path)

    log("正在生成监测点汇总Excel…", 6)
    excel_writer.build_summary_excel(result, s_path)

    log("全部处理完成！", 7)
    print("\n生成文件：")
    print(f"  1. {p_path}")
    print(f"  2. {w_path}")
    print(f"  3. {s_path}")
    print(f"\n基础数据集记录数：{len(result.base)}；"
          f"BI最终记录数：{len(result.bi_final)}；ADI最终记录数：{len(result.adi_final)}；"
          f"删除记录数：{len(result.deletions)}")


def main():
    parser = argparse.ArgumentParser(description="蚊媒监测数据处理工具")
    parser.add_argument("--cli", action="store_true", help="命令行模式（不启动图形界面）")
    parser.add_argument("-i", "--input", help="源Excel文件路径")
    parser.add_argument("-y", "--year", type=int, help="目标年份")
    parser.add_argument("-m", "--month", type=int, help="目标月份")
    parser.add_argument("-d", "--day", type=int, help="目标日期")
    parser.add_argument("-e", "--exclude", help="排除字段（可选）")
    parser.add_argument("-o", "--outdir", help="输出目录（默认与源文件相同）")
    args = parser.parse_args()

    if args.cli:
        if not args.input or args.year is None or args.month is None or args.day is None:
            parser.error("CLI模式必须提供 -i/--input、-y/--year、-m/--month、-d/--day")
        run_cli(args)
        return

    # 图形界面模式
    try:
        from app.gui import run
        run()
    except Exception as exc:  # noqa: BLE001
        print(f"无法启动图形界面（可能缺少Tk支持）：{exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
