from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[3]


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Run quote-update batch pipeline.")
    p.add_argument("--project-dir", required=True, help="项目Excel目录")
    p.add_argument("--glob", default="*.xlsx", help="项目文件匹配模式")
    p.add_argument("--mode", choices=("web", "image_doc", "both"), default="both", help="执行模式")
    p.add_argument("--list-url", help="网价列表页URL")
    p.add_argument("--detail-url", help="网价详情页URL")
    p.add_argument("--account-file", default="网站账号密码.txt", help="账号密码文件")
    p.add_argument("--username", help="账号")
    p.add_argument("--password", help="密码")
    p.add_argument("--headless", action="store_true", help="网价抓取启用无头")
    p.add_argument("--image-source-map", help="批量图片源映射json")
    p.add_argument("--artifact-dir", default="运行产物", help="产物目录")
    p.add_argument("--dry-run", action="store_true", help="预演流程，只生成报告，不修改项目Excel")
    p.add_argument("--confirm-write", action="store_true", help="明确允许写入项目Excel")
    p.add_argument("--refresh-web-artifacts", action="store_true", help="confirm-write时强制重跑网价抓取与对照生成")
    p.add_argument("--refresh-image-artifacts", action="store_true", help="confirm-write时强制重跑OCR与图片文档对照生成")
    p.add_argument("--report-out", help="总结报告路径")
    return p


def main() -> int:
    args = _build_parser().parse_args()
    root = _repo_root()
    cmd = [
        sys.executable,
        "-m",
        "ocr_price.pipeline",
        "batch",
        "--project-dir",
        args.project_dir,
        "--glob",
        args.glob,
        "--mode",
        args.mode,
        "--account-file",
        args.account_file,
        "--artifact-dir",
        args.artifact_dir,
    ]
    if args.list_url:
        cmd.extend(["--list-url", args.list_url])
    if args.detail_url:
        cmd.extend(["--detail-url", args.detail_url])
    if args.username:
        cmd.extend(["--username", args.username])
    if args.password:
        cmd.extend(["--password", args.password])
    if args.headless:
        cmd.append("--headless")
    if args.dry_run:
        cmd.append("--dry-run")
    if args.confirm_write:
        cmd.append("--confirm-write")
    if args.refresh_web_artifacts:
        cmd.append("--refresh-web-artifacts")
    if args.refresh_image_artifacts:
        cmd.append("--refresh-image-artifacts")
    if args.image_source_map:
        cmd.extend(["--image-source-map", args.image_source_map])
    if args.report_out:
        cmd.extend(["--report-out", args.report_out])
    return subprocess.run(cmd, cwd=root).returncode


if __name__ == "__main__":
    raise SystemExit(main())
