from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .config_loader import load_config
from .send_service import build_send_plan, send_plan


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="发送单票 Invoice 邮件（企业邮 SMTP）")
    parser.add_argument(
        "--bill",
        "-b",
        required=True,
        help="账单 PDF 文件夹，如 output/#DS-06052026-SMART已确认-67788.41",
    )
    parser.add_argument(
        "--yes",
        "-y",
        action="store_true",
        help="跳过确认直接发送",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    cfg = load_config()
    plan = build_send_plan(Path(args.bill), cfg)
    print("账单:", plan.bill_name)
    print("主题:", plan.subject)
    print("To:", ", ".join(plan.to))
    print("CC:", ", ".join(plan.cc))
    print(f"附件: {len(plan.pdfs)} 个, {plan.attachment_mb:.1f} MB")
    print("校验:", plan.message)
    if not plan.batch_ok:
        print("发送被阻断", file=sys.stderr)
        return 1
    if not args.yes:
        ans = input("确认发送？[y/N] ").strip().lower()
        if ans not in {"y", "yes"}:
            print("已取消")
            return 0
    send_plan(plan, cfg)
    print("发送成功")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
