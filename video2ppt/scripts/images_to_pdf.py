#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将图片文件夹批量生成为 PDF（每张图片一页，无需 Office / LibreOffice）。

移植自 integrated_tool_v4.py 的 images_to_pdf（GUI 版），改为无界面调用。

压缩模式（默认开启）：
  - 图片按 --max-width 等比缩放，再以 --jpeg-quality 重编码为 JPEG
  - 优先使用 img2pdf 流式合并（不二次重编码），否则回退 Pillow 逐页写入
关闭压缩（--no-compress）：原画质嵌入，等同 v3 行为。
"""

import argparse
import io
import json
import os
import sys

from _common import (ThrottledProgress, collect_images, human_size,
                     make_logger, setup_runtime)

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    import img2pdf
    _HAS_IMG2PDF = True
except ImportError:
    _HAS_IMG2PDF = False


def images_to_pdf(image_dir, pdf_path, sort_mode="name", sort_order="asc",
                  dpi=150, enable_compress=True, jpeg_quality=85, max_width=1920,
                  log=None, progress=None):
    """将目录中的图片逐页写入 PDF，返回成功写入的页数。"""
    image_paths = collect_images(image_dir, sort_mode, sort_order)
    if not image_paths:
        if log:
            log(f"❌ 在目录 {image_dir} 中未找到任何图片文件")
        return 0

    total = len(image_paths)
    save_dir = os.path.dirname(os.path.abspath(pdf_path))
    if save_dir:
        os.makedirs(save_dir, exist_ok=True)

    # ============================================================
    #  路线一：不压缩 —— 原画质直接嵌入
    # ============================================================
    if not enable_compress:
        engine = "img2pdf(流式原画质)" if _HAS_IMG2PDF else "Pillow(逐页原画质)"
        if log:
            log(f"✅ 找到 {total} 张图片（{sort_mode}排序，{sort_order}方向），"
                f"压缩已关闭，使用 {engine} 生成PDF...")
        if _HAS_IMG2PDF:
            try:
                with open(pdf_path, "wb") as outf:
                    outf.write(img2pdf.convert(image_paths, dpi=dpi))
                if progress:
                    progress(100, "PDF生成完成")
                if log:
                    log(f"\n🎉 PDF生成成功！共 {total} 页（原画质），"
                        f"路径：{os.path.abspath(pdf_path)}")
                    log(f"📄 文件大小：{human_size(os.path.getsize(pdf_path))}")
                return total
            except Exception as e:
                if log:
                    log(f"⚠️ img2pdf 失败({e})，回退到 Pillow 逐页模式...")

        written = 0
        for idx, img_path in enumerate(image_paths, 1):
            try:
                with Image.open(img_path) as img:
                    img.load()
                    if img.mode != "RGB":
                        img = img.convert("RGB")
                    img.save(pdf_path, "PDF", resolution=float(dpi),
                             save_all=(idx == 1), append=(idx > 1))
                written += 1
                if log:
                    log(f"✅ 已写入第{idx}/{total}页：{os.path.basename(img_path)}")
                if progress:
                    progress(int(idx / total * 100), f"写入PDF [{idx}/{total}]")
            except Exception as e:
                if log:
                    log(f"⚠️ 第{idx}张图片 {os.path.basename(img_path)} 处理失败：{e}")
                continue
        if written == 0:
            if log:
                log("❌ 没有成功写入任何图片，PDF生成失败")
            return 0
        if log:
            log(f"\n🎉 PDF生成成功！共 {written} 页（原画质），"
                f"路径：{os.path.abspath(pdf_path)}")
            log(f"📄 文件大小：{human_size(os.path.getsize(pdf_path))}")
        return written

    # ============================================================
    #  路线二：压缩 —— 缩放 + JPEG 重编码
    # ============================================================
    if log:
        w_desc = f"{max_width}px" if max_width > 0 else "不限"
        log(f"✅ 找到 {total} 张图片，启用压缩（JPEG质量 {jpeg_quality}，"
            f"最大宽 {w_desc}，DPI {dpi}）...")

    quality = max(10, min(100, int(jpeg_quality)))
    compressed = []   # [(BytesIO, basename, orig_bytes, comp_bytes), ...]
    orig_total = 0
    comp_total = 0
    for idx, img_path in enumerate(image_paths, 1):
        try:
            orig_size = os.path.getsize(img_path)
            with Image.open(img_path) as img:
                img = img.convert("RGB")
                if max_width > 0 and img.width > max_width:
                    ratio = max_width / img.width
                    new_h = max(1, int(img.height * ratio))
                    img = img.resize((max_width, new_h), Image.LANCZOS)
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=quality, optimize=True)
                comp_size = buf.tell()
                buf.seek(0)
                compressed.append((buf, os.path.basename(img_path), orig_size, comp_size))
                orig_total += orig_size
                comp_total += comp_size
                if log:
                    log(f"✅ 压缩 [{idx}/{total}] {os.path.basename(img_path)}："
                        f"{human_size(orig_size)} → {human_size(comp_size)} "
                        f"({img.width}x{img.height})")
                if progress:
                    progress(int(idx / total * 80), f"压缩图片 [{idx}/{total}]")
        except Exception as e:
            if log:
                log(f"⚠️ 第{idx}张图片 {os.path.basename(img_path)} 压缩失败：{e}")
            continue

    if not compressed:
        if log:
            log("❌ 没有成功处理任何图片，PDF生成失败")
        return 0

    if log and orig_total > 0:
        save_pct = (1 - comp_total / orig_total) * 100
        log(f"📊 压缩汇总：源图 {human_size(orig_total)} → 压缩后 {human_size(comp_total)}"
            f"（节省 {save_pct:.1f}%）")

    written = 0
    # 优先 img2pdf 合并压缩后的 JPEG（直接嵌入、不重编码）
    if _HAS_IMG2PDF:
        try:
            pdf_bytes = img2pdf.convert(*[buf for buf, _, _, _ in compressed], dpi=dpi)
            with open(pdf_path, "wb") as outf:
                outf.write(pdf_bytes)
            written = len(compressed)
            if progress:
                progress(100, "PDF生成完成")
        except Exception as e:
            if log:
                log(f"⚠️ img2pdf 合并失败({e})，回退 Pillow 逐页写入...")
            written = 0

    # 回退：Pillow 逐页 append
    if written == 0:
        for idx, (buf, name, _, _) in enumerate(compressed, 1):
            buf.seek(0)
            try:
                with Image.open(buf) as img:
                    img.load()
                    if img.mode != "RGB":
                        img = img.convert("RGB")
                    img.save(pdf_path, "PDF", resolution=float(dpi),
                             save_all=(idx == 1), append=(idx > 1))
                written += 1
                if progress:
                    progress(80 + int(idx / len(compressed) * 20),
                             f"写入PDF [{idx}/{len(compressed)}]")
            except Exception as e:
                if log:
                    log(f"⚠️ 第{idx}页 {name} 写入PDF失败：{e}")
                continue

    if written == 0:
        if log:
            log("❌ PDF写入失败，未生成有效文件")
        return 0

    if log:
        final_size = os.path.getsize(pdf_path) if os.path.exists(pdf_path) else 0
        log(f"\n🎉 PDF生成成功！共 {written} 页（压缩模式），"
            f"路径：{os.path.abspath(pdf_path)}")
        log(f"📄 最终PDF大小：{human_size(final_size)}")
    return written


def build_parser():
    p = argparse.ArgumentParser(
        description="将图片文件夹批量生成为 PDF（每张图片一页）。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("-i", "--input", required=True, help="图片文件夹")
    p.add_argument("-o", "--output", default="output_presentation.pdf",
                   help="PDF 保存路径")
    p.add_argument("--sort", choices=["name", "date"], default="name",
                   help="按文件名（自然排序）或修改时间")
    p.add_argument("--order", choices=["asc", "desc"], default="asc",
                   help="排序方向")
    p.add_argument("--dpi", type=int, default=150, help="PDF 页面 DPI 声明")

    g = p.add_argument_group("压缩")
    g.add_argument("--compress", dest="compress", action="store_true", default=True,
                   help="压缩图片以减小 PDF 体积（默认开启）")
    g.add_argument("--no-compress", dest="compress", action="store_false",
                   help="保留原画质，不压缩")
    g.add_argument("--jpeg-quality", type=int, default=85, dest="jpeg_quality",
                   help="JPEG 质量(10-100)，越小文件越小")
    g.add_argument("--max-width", type=int, default=1920, dest="max_width",
                   help="最大像素宽度，超出则等比缩小；0 表示不缩放")

    g = p.add_argument_group("其他")
    g.add_argument("--quiet", action="store_true", help="仅输出错误与警告")
    g.add_argument("--json", action="store_true", help="在 stdout 输出 JSON 摘要")
    return p


def main(argv=None):
    setup_runtime()
    parser = build_parser()
    args = parser.parse_args(argv)

    if Image is None:
        print("缺少依赖 Pillow，请执行：python -m pip install Pillow", file=sys.stderr)
        return 3
    if not os.path.isdir(args.input):
        parser.error(f"图片文件夹不存在：{args.input}")
    if not 1 <= args.dpi <= 1200:
        parser.error("--dpi 必须在 1-1200 之间")
    if not 10 <= args.jpeg_quality <= 100:
        parser.error("--jpeg-quality 必须在 10-100 之间")
    if args.max_width < 0:
        parser.error("--max-width 不能为负数")

    image_paths = collect_images(args.input, args.sort, args.order)
    if not image_paths:
        print(f"在目录 {args.input} 中未找到任何图片文件。", file=sys.stderr)
        return 1

    log = make_logger(args.quiet)
    progress = ThrottledProgress(enabled=not args.quiet, prefix="进度")

    written = images_to_pdf(
        args.input, args.output,
        sort_mode=args.sort, sort_order=args.order,
        dpi=args.dpi, enable_compress=args.compress,
        jpeg_quality=args.jpeg_quality, max_width=args.max_width,
        log=log, progress=progress,
    )

    final_size = os.path.getsize(args.output) if os.path.exists(args.output) else 0
    summary = {
        "status": "ok" if written > 0 else "failed",
        "output": os.path.abspath(args.output),
        "images_found": len(image_paths),
        "pages": written,
        "size_bytes": final_size,
        "compress": args.compress,
        "engine": ("img2pdf" if _HAS_IMG2PDF else "Pillow"),
    }
    if args.json:
        print(json.dumps(summary, ensure_ascii=False))
    else:
        print(f"完成：{written} 页 → {summary['output']}（{human_size(final_size)}）")
    return 0 if written > 0 else 1


if __name__ == "__main__":
    sys.exit(main())
