from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _quote(value: str) -> str:
    return f'"{value}"'


def _headless_flag(args: argparse.Namespace) -> str:
    return " `\n  --headless" if getattr(args, "headless", True) else ""


def _common_rules() -> str:
    return """\
硬规则：
1. 先阅读 AGENTS.md、skills/quote-update/SKILL.md、skills/quote-update/references/rules.md、skills/quote-update/references/workflow.md、README.md 第 5.0 节。
2. 禁止直接修改 Excel，禁止手工调用 openpyxl 写单元格，只能使用正式入口脚本。
3. 禁止自动确认厂家映射；发现待确认、新厂家、冲突、价格偏差、OCR异常、库存颜色异常、登录失败时必须停止并报告。
4. 禁止编造价格、库存、厂家映射和数量统计；最终结果必须来自 JSON/Markdown 报告或实际文件。
5. confirm-write 必须使用同一次 dry-run 生成的 Manifest 路径。
6. 如果 manifest 校验提示项目或源文件已变化，必须重新 dry-run，禁止绕过。
"""


def _single_dry_run(args: argparse.Namespace) -> str:
    command = "\n".join(
        [
            "python skills/quote-update/scripts/run_single.py `",
            f"  --project {_quote(args.project)} `",
            f"  --mode {args.mode} `",
            "  --dry-run" + _headless_flag(args),
        ]
    )
    return f"""\
请在项目根目录执行一次单文件 dry-run，不要写入 Excel。

项目根目录：{Path.cwd()}
项目文件：{args.project}
更新模式：{args.mode}

{_common_rules()}
执行命令：
```powershell
{command}
```

执行后请报告：
- Status
- Manifest 路径
- Report / MarkdownReport / Events 路径
- 是否有待确认厂家或新厂家
- 是否有价格偏差、OCR异常、库存颜色异常、登录失败
- 是否可以进入 confirm-write

注意：dry-run 阶段禁止执行确认写入。
"""


def _single_confirm(args: argparse.Namespace) -> str:
    command = "\n".join(
        [
            "python skills/quote-update/scripts/run_single.py `",
            f"  --project {_quote(args.project)} `",
            f"  --mode {args.mode} `",
            "  --confirm-write `",
            f"  --manifest {_quote(args.manifest)}" + _headless_flag(args),
        ]
    )
    return f"""\
我已确认同一次 dry-run 报告无异常。请执行单文件 confirm-write。

项目根目录：{Path.cwd()}
项目文件：{args.project}
更新模式：{args.mode}
Manifest：{args.manifest}

{_common_rules()}
执行命令：
```powershell
{command}
```

执行后必须按 README.md 第 5.0 节格式汇报完整结果，并列出 Report、MarkdownReport、Events 路径。
"""


def _batch_dry_run(args: argparse.Namespace) -> str:
    command = "\n".join(
        [
            "python skills/quote-update/scripts/run_batch.py `",
            f"  --project-dir {_quote(args.project_dir)} `",
            f"  --glob {_quote(args.glob)} `",
            f"  --mode {args.mode} `",
            "  --dry-run" + _headless_flag(args),
        ]
    )
    return f"""\
请在项目根目录执行一次批量 dry-run，不要写入 Excel。

项目根目录：{Path.cwd()}
项目目录：{args.project_dir}
匹配规则：{args.glob}
更新模式：{args.mode}

{_common_rules()}
执行命令：
```powershell
{command}
```

执行后请按项目分组报告：
- Status
- Report / Events 路径
- 每个项目是否有 Manifest 路径
- 待确认厂家、新厂家、价格偏差、OCR异常、库存颜色异常、登录失败
- 哪些项目可以进入 confirm-write，哪些必须先人工处理

注意：dry-run 阶段禁止执行确认写入。
"""


def _batch_confirm(args: argparse.Namespace) -> str:
    command = "\n".join(
        [
            "python skills/quote-update/scripts/run_batch.py `",
            f"  --project-dir {_quote(args.project_dir)} `",
            f"  --glob {_quote(args.glob)} `",
            f"  --mode {args.mode} `",
            "  --confirm-write `",
            f"  --manifest {_quote(args.manifest)}" + _headless_flag(args),
        ]
    )
    return f"""\
我已确认同一次批量 dry-run 报告无异常。请执行批量 confirm-write。

项目根目录：{Path.cwd()}
项目目录：{args.project_dir}
匹配规则：{args.glob}
更新模式：{args.mode}
Manifest：{args.manifest}

{_common_rules()}
执行命令：
```powershell
{command}
```

执行后必须按 README.md 第 5.0 节格式逐项目汇报完整结果，并列出 Report、Events 路径。
"""


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate safe prompts for other agents.")
    sub = parser.add_subparsers(dest="command", required=True)

    for name in ("single-dry-run", "single-confirm"):
        p = sub.add_parser(name)
        p.add_argument("--project", required=True, help="项目Excel路径")
        p.add_argument("--mode", choices=("web", "image_doc", "both"), default="both")
        p.add_argument("--headless", dest="headless", action="store_true", default=True)
        p.add_argument("--no-headless", dest="headless", action="store_false")
        if name == "single-confirm":
            p.add_argument("--manifest", required=True, help="dry-run生成的manifest路径")

    for name in ("batch-dry-run", "batch-confirm"):
        p = sub.add_parser(name)
        p.add_argument("--project-dir", default="项目报价", help="项目Excel目录")
        p.add_argument("--glob", default="*.xlsx", help="项目文件匹配规则")
        p.add_argument("--mode", choices=("web", "image_doc", "both"), default="both")
        p.add_argument("--headless", dest="headless", action="store_true", default=True)
        p.add_argument("--no-headless", dest="headless", action="store_false")
        if name == "batch-confirm":
            p.add_argument("--manifest", required=True, help="dry-run生成的manifest路径")

    return parser


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args = _build_parser().parse_args()
    if args.command == "single-dry-run":
        print(_single_dry_run(args))
    elif args.command == "single-confirm":
        print(_single_confirm(args))
    elif args.command == "batch-dry-run":
        print(_batch_dry_run(args))
    elif args.command == "batch-confirm":
        print(_batch_confirm(args))
    else:
        raise SystemExit(f"unsupported command: {args.command}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
