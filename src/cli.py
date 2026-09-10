from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .pipeline import run_pipeline

ROOT = Path(__file__).resolve().parents[1]


def _collect_excels(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(f"不是 Excel 文件: {input_path}")
        return [input_path]
    if input_path.is_dir():
        files = sorted(
            p
            for p in input_path.iterdir()
            if p.is_file()
            and p.suffix.lower() in {".xlsx", ".xlsm"}
            and not p.name.startswith("~$")
        )
        return files
    raise FileNotFoundError(f"输入路径不存在: {input_path}")


def _print_one(result: dict) -> None:
    print("==== 运行结果 ====")
    print(f"账单: {result['bill_name']}")
    print(f"渲染模式: {result.get('render_mode')}")
    print(f"批次放行: {result['batch_ok']}")
    print(f"说明: {result['message']}")
    print(f"系统合计: {result['system_total']} / 汇总合计: {result['summary_total']}")
    print(f"生成 PDF 数: {result['container_count']}")
    print(f"报告: {result['report']}")
    if result["generated"]:
        print("输出目录:", Path(result["generated"][0]).parent)
        for p in result["generated"][:5]:
            print(" -", Path(p).name)
        if len(result["generated"]) > 5:
            print(f" - ... 共 {len(result['generated'])} 个")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="SMART 单票 Invoice 生成（Excel → 按柜 PDF）"
    )
    parser.add_argument(
        "--input",
        "-i",
        default=str(ROOT / "input"),
        help="已确认账单 Excel 文件，或包含多个 Excel 的文件夹（默认 ./input）",
    )
    parser.add_argument("--output", "-o", default=str(ROOT / "output"), help="PDF 输出根目录")
    parser.add_argument("--config", "-c", default=None, help="config.yaml 路径")
    parser.add_argument("--logs", default=str(ROOT / "logs"), help="日志/报告根目录")
    parser.add_argument(
        "--mode",
        choices=["template", "reportlab"],
        default=None,
        help="template=填Excel模板再导出PDF（默认）；reportlab=代码绘制",
    )
    parser.add_argument(
        "--flat",
        action="store_true",
        help="多文件时也把 PDF 直接丢进 output 根目录（默认按账单名分子目录，避免重名覆盖）",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    input_path = Path(args.input)
    try:
        files = _collect_excels(input_path)
    except (FileNotFoundError, ValueError) as e:
        print(str(e), file=sys.stderr)
        return 2

    if not files:
        print(f"未找到 Excel：{input_path}", file=sys.stderr)
        return 2

    multi = len(files) > 1 or input_path.is_dir()
    print(f"待处理 Excel：{len(files)} 个")
    for f in files:
        print(f" - {f.name}")

    exit_code = 0
    for f in files:
        # 多账单时默认分子目录，避免不同账单同柜号 PDF 互相覆盖
        if multi and not args.flat:
            out_dir = Path(args.output) / f.stem
            log_dir = Path(args.logs) / f.stem
        else:
            out_dir = Path(args.output)
            log_dir = Path(args.logs)

        print()
        print(f">>> 处理: {f.name}")
        result = run_pipeline(
            input_xlsx=f,
            output_dir=out_dir,
            config_path=args.config,
            logs_dir=log_dir,
            render_mode=args.mode,
        )
        _print_one(result)
        if not result["batch_ok"]:
            exit_code = 1

    if multi:
        print()
        print(f"全部完成。PDF 根目录: {Path(args.output).resolve()}")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
