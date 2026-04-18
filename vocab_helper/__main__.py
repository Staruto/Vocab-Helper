from __future__ import annotations

import argparse


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="vocab_helper")
    parser.add_argument(
        "mode",
        nargs="?",
        choices=["cli", "gui"],
        default="cli",
        help="Runtime mode. Default is cli.",
    )
    parser.add_argument(
        "--command",
        dest="command",
        default=None,
        help="Run a single CLI command and exit (cli mode only).",
    )
    parser.add_argument(
        "--no-banner",
        action="store_true",
        help="Hide the CLI banner.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.mode == "gui":
        from .app import main as app_main

        app_main()
        return 0

    from .cli import main as cli_main

    return cli_main(command=args.command, show_banner=not args.no_banner)


if __name__ == "__main__":
    raise SystemExit(main())
