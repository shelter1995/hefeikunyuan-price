from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[3]


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
    p.add_argument("--report-out", help="总结报告路径")
    return p


def main() -> int:
    args = _build_parser().parse_args()
    root = _repo_root()
    cmd = [
        sys.executable,
        "-m",
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
        cmd.extend(["--list-url", args.list_url])
    if args.detail_url:
        cmd.extend(["--detail-url", args.detail_url])
    if args.username:
        cmd.extend(["--username", args.username])
    if args.password:
        cmd.extend(["--password", args.password])
    if args.headless:
        cmd.append("--headless")
    if args.manual_login_timeout != 180:
        cmd.extend(["--manual-login-timeout", str(args.manual_login_timeout)])
    if args.dry_run:
        cmd.append("--dry-run")
    if args.confirm_write:
        cmd.append("--confirm-write")
    if args.refresh_web_artifacts:
        cmd.append("--refresh-web-artifacts")
    if args.refresh_image_artifacts:
        cmd.append("--refresh-image-artifacts")
    if args.image_inputs:
        cmd.append("--image-inputs")
        cmd.extend(args.image_inputs)
    if args.image_jsons:
        cmd.append("--image-jsons")
        cmd.extend(args.image_jsons)
    if args.report_out:
        cmd.extend(["--report-out", args.report_out])
    return subprocess.run(cmd, cwd=root).returncode


if __name__ == "__main__":
    raise SystemExit(main())
