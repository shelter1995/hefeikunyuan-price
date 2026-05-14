from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[3]


# Ensure repo root is on path so `ocr_price` can be imported
_repo_root().resolve()


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Run quote-update single pipeline.")
    p.add_argument("--project", required=True, help="项目Excel路径")
    p.add_argument("--mode", choices=("web", "image_doc", "both"), default="both", help="执行模式")
    p.add_argument("--list-url", help="网价列表页URL")
    p.add_argument("--detail-url", help="网价详情页URL")
    p.add_argument("--account-file", default="网站账号密码.txt", help="账号密码文件")
    p.add_argument("--username", help="账号")
    p.add_argument("--password", help="密码")
    p.add_argument("--headless", action="store_true", help="网价抓取启用无头")
    p.add_argument("--manual-login-timeout", type=int, default=180, help="有头模式下等待人工登录的秒数")
    p.add_argument("--image-inputs", nargs="*", default=[], help="图片/文档原始文件")
    p.add_argument("--image-jsons", nargs="*", default=[], help="已提取OCR json")
    p.add_argument("--artifact-dir", default="运行产物", help="产物目录")
    p.add_argument("--dry-run", action="store_true", help="预演流程，只生成报告，不修改项目Excel")
    p.add_argument("--confirm-write", action="store_true", help="明确允许写入项目Excel")
    p.add_argument("--refresh-web-artifacts", action="store_true", help="confirm-write时强制重跑网价抓取与对照生成")
    p.add_argument("--refresh-image-artifacts", action="store_true", help="confirm-write时强制重跑OCR与图片文档对照生成")
    p.add_argument("--manifest", help="dry-run生成的manifest路径")
    p.add_argument("--report-out", help="总结报告路径")
    return p


def main() -> int:
    args = _build_parser().parse_args()
    root = _repo_root()
    old_cwd = Path.cwd()
    old_argv = sys.argv[:]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    try:
        os.chdir(root)
        sys.argv = [
            "ocr_price.pipeline",
            "single",
            "--project",
            args.project,
            "--mode",
            args.mode,
            "--account-file",
            args.account_file,
            "--artifact-dir",
            args.artifact_dir,
        ]
        if args.list_url:
            sys.argv.extend(["--list-url", args.list_url])
        if args.detail_url:
            sys.argv.extend(["--detail-url", args.detail_url])
        if args.username:
            sys.argv.extend(["--username", args.username])
        if args.password:
            sys.argv.extend(["--password", args.password])
        if args.headless:
            sys.argv.append("--headless")
        if args.manual_login_timeout != 180:
            sys.argv.extend(["--manual-login-timeout", str(args.manual_login_timeout)])
        if args.dry_run:
            sys.argv.append("--dry-run")
        if args.confirm_write:
            sys.argv.append("--confirm-write")
        if args.refresh_web_artifacts:
            sys.argv.append("--refresh-web-artifacts")
        if args.refresh_image_artifacts:
            sys.argv.append("--refresh-image-artifacts")
        if args.manifest:
            sys.argv.extend(["--manifest", args.manifest])
        if args.image_inputs:
            sys.argv.append("--image-inputs")
            sys.argv.extend(args.image_inputs)
        if args.image_jsons:
            sys.argv.append("--image-jsons")
            sys.argv.extend(args.image_jsons)
        if args.report_out:
            sys.argv.extend(["--report-out", args.report_out])
        from ocr_price import pipeline

        return pipeline.main()
    finally:
        sys.argv = old_argv
        os.chdir(old_cwd)


if __name__ == "__main__":
    raise SystemExit(main())
