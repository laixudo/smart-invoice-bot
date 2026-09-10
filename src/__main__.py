"""入口：无参数开 GUI；有参数走 CLI。"""

from __future__ import annotations

import sys


def main() -> None:
    if len(sys.argv) <= 1 or sys.argv[1] in {"gui", "--gui"}:
        from .gui import main as gui_main

        gui_main()
        return
    from .cli import main as cli_main

    raise SystemExit(cli_main())


if __name__ == "__main__":
    main()
