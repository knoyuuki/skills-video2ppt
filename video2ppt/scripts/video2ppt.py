#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一键流水线：视频 → 画面帧 → PPTX / PDF。

示例：
  python scripts/video2ppt.py -i lecture.mp4 -o out --formats pptx,pdf
  python scripts/video2ppt.py -i ./videos -o out --start 00:05:00 --end 00:40:00 \\
      --similarity 0.82 --fit fit --json
"""

import argparse
import json
import os
import sys
from pathlib import Path

from _common import (ThrottledProgress, make_logger, parse_crop, parse_rgb,
                     parse_time, setup_runtime)
from images_to_pdf import images_to_pdf
from images_to_pptx import images_to_pptx
from video_to_frames import (VideoFrameExtractor, dependency_errors,
                             expand_inputs, missing_dependencies)


def build_parser():
    p = argparse.ArgumentParser(
        description="一键流水线：从视频提取幻灯片画面，并生成 PPTX / PDF。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("-i", "--input", nargs="+", required=True,
                   help="视频文件、目录或通配符（可多个）")
    p.add_argument("-o", "--outdir", default="video2ppt_output",
                   help="输出目录（帧保存在其 frames/ 子目录）")
    p.add_argument("--name", default=None,
                   help="输出文件基名（默认取第一个视频的文件名）")
    p.add_argument("--formats", default="pptx,pdf",
                   help="要生成的格式，逗号分隔：pptx,pdf；留空则只抽帧")

    g = p.add_argument_group("抽帧参数")
    g.add_argument("--similarity", type=float, default=0.8,
                   help="相似度阈值(0-1)，越低提取的页越少")
    g.add_argument("--min-interval", type=int, default=5, dest="min_interval",
                   help="两次保存之间的最小帧数间隔")
    g.add_argument("--delay-frames", type=int, default=5, dest="delay_frames",
                   help="检测到切换后延迟的帧数")
    g.add_argument("--no-stable-check", action="store_false", dest="stable_check",
                   help="关闭画面稳定性检查")
    g.add_argument("--stable-threshold", type=float, default=0.95,
                   dest="stable_threshold", help="稳定性阈值(0-1)")
    g.add_argument("--skip", type=int, default=15,
                   help="跳帧数：每 (n+1) 帧处理 1 帧，0 表示逐帧")
    g.add_argument("--resize", type=int, default=480,
                   help="相似度计算的缩放宽度，0 表示不缩放")
    g.add_argument("--start", default="0",
                   help="开始时间：秒 / MM:SS / HH:MM:SS")
    g.add_argument("--end", default="0",
                   help="结束时间：秒 / MM:SS / HH:MM:SS")
    g.add_argument("--crop", default="none",
                   help="裁剪方向：none / all / top,bottom,left,right")
    g.add_argument("--crop-mode", choices=["auto", "fixed"], default="auto",
                   dest="crop_mode")
    g.add_argument("--black-threshold", type=int, default=30,
                   dest="black_threshold")
    g.add_argument("--frame-format", choices=["png", "jpg", "jpeg"], default="png",
                   dest="frame_format", help="截图格式")
    g.add_argument("--quality", type=int, default=100, help="JPEG 质量(1-100)")

    g = p.add_argument_group("PPTX 参数")
    g.add_argument("--slide-width", type=float, default=16.0, dest="slide_width",
                   help="页面宽度（英寸）")
    g.add_argument("--slide-height", type=float, default=9.0, dest="slide_height",
                   help="页面高度（英寸）")
    g.add_argument("--bg", default="255,255,255", help="PPT 背景色 R,G,B")
    g.add_argument("--title", default="", help="每页标题前缀（留空则无）")
    g.add_argument("--fit", choices=["stretch", "fit"], default="stretch",
                   help="stretch=全屏拉伸；fit=保持比例居中留白")

    g = p.add_argument_group("PDF 参数")
    g.add_argument("--dpi", type=int, default=150, help="PDF 页面 DPI")
    g.add_argument("--no-compress", action="store_false", dest="compress",
                   help="生成 PDF 时保留原画质（默认压缩）")
    g.add_argument("--jpeg-quality", type=int, default=85, dest="jpeg_quality",
                   help="PDF 图片 JPEG 质量(10-100)")
    g.add_argument("--max-width", type=int, default=1920, dest="max_width",
                   help="PDF 图片最大宽度，0 表示不缩放")

    g = p.add_argument_group("其他")
    g.add_argument("--sort", choices=["name", "date"], default="name",
                   help="图片排序方式")
    g.add_argument("--order", choices=["asc", "desc"], default="asc",
                   help="排序方向")
    g.add_argument("--verbose", action="store_true", help="输出调试信息")
    g.add_argument("--quiet", action="store_true", help="仅输出错误与警告")
    g.add_argument("--json", action="store_true", help="在 stdout 输出 JSON 摘要")
    return p


def main(argv=None):
    setup_runtime()
    parser = build_parser()
    args = parser.parse_args(argv)

    missing = missing_dependencies()
    if missing:
        print(f"缺少或无法加载必需依赖：{', '.join(missing)}。请执行："
              f"python -m pip install -r requirements.txt", file=sys.stderr)
        for pkg, err in dependency_errors().items():
            last_line = err.strip().splitlines()[-1] if err.strip() else ""
            print(f"  - {pkg}: {last_line}", file=sys.stderr)
        return 3

    formats = [f.strip().lower() for f in args.formats.split(",") if f.strip()]
    unknown = [f for f in formats if f not in ("pptx", "pdf")]
    if unknown:
        parser.error(f"不支持输出格式：{', '.join(unknown)}（可选 pptx,pdf）")

    if not 0.0 <= args.similarity <= 1.0:
        parser.error("--similarity 必须在 0-1 之间")
    if not 0.0 <= args.stable_threshold <= 1.0:
        parser.error("--stable-threshold 必须在 0-1 之间")
    if args.slide_width <= 0 or args.slide_height <= 0:
        parser.error("页面宽高必须大于 0")
    try:
        start_seconds = parse_time(args.start)
        end_seconds = parse_time(args.end)
        crop_sides = parse_crop(args.crop)
        bg_color = parse_rgb(args.bg)
    except ValueError as e:
        parser.error(str(e))
    if end_seconds > 0 and end_seconds <= start_seconds:
        parser.error("--end 必须大于 --start")

    try:
        inputs = expand_inputs(args.input)
    except FileNotFoundError as e:
        parser.error(str(e))
    if not inputs:
        print("未找到任何视频文件。", file=sys.stderr)
        return 1

    log = make_logger(args.quiet)
    progress = ThrottledProgress(enabled=not args.quiet, prefix="抽帧")

    def progress_cb(current, total, video_name, overall_pct=None):
        pct = overall_pct if overall_pct is not None else (
            int(current / total * 100) if total else 0)
        progress(pct, f"{video_name} {current:,}/{total:,} 帧")

    outdir = Path(args.outdir).absolute()
    frames_dir = outdir / "frames"
    base_name = args.name or Path(inputs[0]).stem
    # 去掉文件名中可能影响路径的字符
    base_name = "".join(ch for ch in base_name if ch not in '<>:"/\\|?*').strip() or "output"

    config = {
        "input": inputs,
        "output": str(frames_dir),
        "similarity": args.similarity,
        "min_interval": args.min_interval,
        "delay_frames": args.delay_frames,
        "stable_check": args.stable_check,
        "stable_threshold": args.stable_threshold,
        "format": args.frame_format,
        "quality": args.quality,
        "skip": args.skip,
        "resize": args.resize,
        "verbose": args.verbose,
        "crop_sides": crop_sides,
        "crop_mode": args.crop_mode,
        "black_threshold": args.black_threshold,
        "crop_margins": {"top": 10, "bottom": 10, "left": 8, "right": 8},
        "start_seconds": start_seconds,
        "end_seconds": end_seconds,
    }

    extractor = VideoFrameExtractor(config, log_callback=log,
                                    progress_callback=progress_cb)
    try:
        count = extractor.run()
    except KeyboardInterrupt:
        extractor.stop()
        print("已中断。", file=sys.stderr)
        return 130

    summary = {
        "status": "ok",
        "outdir": str(outdir),
        "frames_dir": str(frames_dir),
        "frames": count,
        "outputs": {},
        "errors": extractor.errors,
    }

    if count == 0:
        summary["status"] = "empty"
        if args.json:
            print(json.dumps(summary, ensure_ascii=False))
        else:
            print("未提取到任何画面，请调整 --similarity/--skip 等参数后重试。",
                  file=sys.stderr)
        return 1

    if "pptx" in formats:
        pptx_path = str(outdir / f"{base_name}.pptx")
        slides = images_to_pptx(
            str(frames_dir), pptx_path,
            page_size=(args.slide_width, args.slide_height),
            bg_color=bg_color, slide_title=args.title.strip(),
            sort_mode=args.sort, sort_order=args.order, fit_mode=args.fit,
            log=log,
            progress=ThrottledProgress(enabled=not args.quiet, prefix="PPT"),
        )
        summary["outputs"]["pptx"] = {"path": pptx_path, "slides": slides}
        if slides == 0:
            summary["errors"].append("PPTX 生成失败")

    if "pdf" in formats:
        pdf_path = str(outdir / f"{base_name}.pdf")
        pages = images_to_pdf(
            str(frames_dir), pdf_path,
            sort_mode=args.sort, sort_order=args.order,
            dpi=args.dpi, enable_compress=args.compress,
            jpeg_quality=args.jpeg_quality, max_width=args.max_width,
            log=log,
            progress=ThrottledProgress(enabled=not args.quiet, prefix="PDF"),
        )
        summary["outputs"]["pdf"] = {"path": pdf_path, "pages": pages}
        if pages == 0:
            summary["errors"].append("PDF 生成失败")

    ok = all(v.get("slides", v.get("pages", 0)) > 0
             for v in summary["outputs"].values()) if summary["outputs"] else True
    if not ok:
        summary["status"] = "partial"

    if args.json:
        print(json.dumps(summary, ensure_ascii=False))
    else:
        print(f"完成：提取 {count} 页画面 → {outdir}")
        for kind, info in summary["outputs"].items():
            label = "页幻灯片" if kind == "pptx" else "页PDF"
            amount = info.get("slides", info.get("pages", 0))
            print(f"  {kind.upper()}: {info['path']}（{amount}{label}）")
    return 0 if summary["status"] == "ok" else 1


if __name__ == "__main__":
    sys.exit(main())
