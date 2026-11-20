#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将图片文件夹批量生成为 PPTX 演示文稿（每张图片一页）。

移植自 integrated_tool_v4.py 的 images_to_pptx（GUI 版），改为无界面调用。
"""

import argparse
import json
import os
import sys

from _common import (ThrottledProgress, collect_images, make_logger,
                     parse_rgb, setup_runtime)

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt
except ImportError:
    Presentation = None

try:
    from PIL import Image
except ImportError:
    Image = None


def images_to_pptx(image_dir, pptx_path, page_size=(16, 9),
                   bg_color=(255, 255, 255), slide_title="",
                   title_size=24, title_color=(0, 0, 0),
                   sort_mode="name", sort_order="asc", fit_mode="stretch",
                   log=None, progress=None, stop_check=None):
    """将目录中的图片逐张生成 PPT 页面，返回成功添加的页数。

    fit_mode:
      stretch - 全屏拉伸（非页面比例的图片会变形）
      fit     - 保持比例缩放并居中留白（不变形）
    """
    image_paths = collect_images(image_dir, sort_mode, sort_order)
    if not image_paths:
        if log:
            log(f"❌ 在目录 {image_dir} 中未找到任何图片文件")
        return 0
    if log:
        log(f"✅ 找到 {len(image_paths)} 张图片（{sort_mode}排序，{sort_order}方向，"
            f"{fit_mode}适应），开始生成PPT...")

    prs = Presentation()
    slide_width_emu = Inches(page_size[0])
    slide_height_emu = Inches(page_size[1])
    prs.slide_width = slide_width_emu
    prs.slide_height = slide_height_emu

    total = len(image_paths)
    added = 0
    for idx, img_path in enumerate(image_paths, 1):
        if stop_check and stop_check():
            if log:
                log("⚠️ 用户取消，PPT生成中止")
            break
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*bg_color)

        if slide_title:
            title_height = Inches(1.0)
            img_top = title_height
            avail_w = slide_width_emu
            avail_h = slide_height_emu - title_height
            title_box = slide.shapes.add_textbox(left=Inches(0), top=Inches(0),
                                                 width=slide_width_emu,
                                                 height=title_height)
            title_frame = title_box.text_frame
            title_frame.text = f"{slide_title} - 第{idx}页"
            title_para = title_frame.paragraphs[0]
            title_para.alignment = PP_ALIGN.CENTER
            title_font = title_para.font
            title_font.size = Pt(title_size)
            title_font.color.rgb = RGBColor(*title_color)
            title_font.name = "微软雅黑"
        else:
            img_top = Inches(0)
            avail_w = slide_width_emu
            avail_h = slide_height_emu

        try:
            if fit_mode == "fit":
                with Image.open(img_path) as im:
                    iw, ih = im.size
                if iw > 0 and ih > 0:
                    scale = min(avail_w / iw, avail_h / ih)
                    draw_w = int(iw * scale)
                    draw_h = int(ih * scale)
                    left = int((avail_w - draw_w) / 2)
                    top = int(img_top + (avail_h - draw_h) / 2)
                    slide.shapes.add_picture(img_path, left=left, top=top,
                                             width=draw_w, height=draw_h)
                else:
                    slide.shapes.add_picture(img_path, left=Inches(0), top=img_top,
                                             width=avail_w, height=avail_h)
                desc = f"fit模式 {iw}x{ih}"
            else:
                pic = slide.shapes.add_picture(img_path, left=Inches(0), top=img_top,
                                               width=avail_w, height=avail_h)
                pic.left = round(pic.left)
                pic.top = round(pic.top)
                pic.width = round(pic.width)
                pic.height = round(pic.height)
                desc = "全屏拉伸"
        except Exception as e:
            if log:
                log(f"⚠️ 第{idx}张图片 {os.path.basename(img_path)} 导入失败：{e}")
            continue

        added += 1
        if log:
            log(f"✅ 已添加第{idx}/{total}页：{os.path.basename(img_path)}（{desc}）")
        if progress:
            progress(int(idx / total * 95), f"生成PPT [{idx}/{total}]")

    try:
        save_dir = os.path.dirname(os.path.abspath(pptx_path))
        if save_dir:
            os.makedirs(save_dir, exist_ok=True)
        prs.save(pptx_path)
    except Exception as e:
        if log:
            log(f"❌ PPT保存失败：{e}")
        return 0

    if progress:
        progress(100, "PPT生成完成")
    if log:
        log(f"\n🎉 PPT生成成功！共 {added} 页，保存路径：{os.path.abspath(pptx_path)}")
    return added


def build_parser():
    p = argparse.ArgumentParser(
        description="将图片文件夹批量生成为 PPTX（每张图片一页）。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("-i", "--input", required=True, help="图片文件夹")
    p.add_argument("-o", "--output", default="output_presentation.pptx",
                   help="PPTX 保存路径")

    g = p.add_argument_group("页面设置")
    g.add_argument("--slide-width", type=float, default=16.0, dest="slide_width",
                   help="页面宽度（英寸）")
    g.add_argument("--slide-height", type=float, default=9.0, dest="slide_height",
                   help="页面高度（英寸）")
    g.add_argument("--bg", default="255,255,255",
                   help="背景颜色 R,G,B")
    g.add_argument("--title", default="",
                   help="每页标题前缀（留空则无标题），形如“标题 - 第N页”")
    g.add_argument("--title-size", type=int, default=24, dest="title_size",
                   help="标题字号")
    g.add_argument("--title-color", default="0,0,0", dest="title_color",
                   help="标题颜色 R,G,B")
    g.add_argument("--fit", choices=["stretch", "fit"], default="stretch",
                   help="stretch=全屏拉伸；fit=保持比例居中留白")

    g = p.add_argument_group("排序")
    g.add_argument("--sort", choices=["name", "date"], default="name",
                   help="按文件名（自然排序）或修改时间")
    g.add_argument("--order", choices=["asc", "desc"], default="asc",
                   help="排序方向")

    g = p.add_argument_group("其他")
    g.add_argument("--quiet", action="store_true", help="仅输出错误与警告")
    g.add_argument("--json", action="store_true", help="在 stdout 输出 JSON 摘要")
    return p


def main(argv=None):
    setup_runtime()
    parser = build_parser()
    args = parser.parse_args(argv)

    if Presentation is None:
        print("缺少依赖 python-pptx，请执行：python -m pip install python-pptx",
              file=sys.stderr)
        return 3
    if Image is None:
        print("缺少依赖 Pillow，请执行：python -m pip install Pillow", file=sys.stderr)
        return 3
    if not os.path.isdir(args.input):
        parser.error(f"图片文件夹不存在：{args.input}")
    if args.slide_width <= 0 or args.slide_height <= 0:
        parser.error("页面宽高必须大于 0")
    try:
        bg_color = parse_rgb(args.bg)
        title_color = parse_rgb(args.title_color)
    except ValueError as e:
        parser.error(str(e))

    images = collect_images(args.input, args.sort, args.order)
    if not images:
        print(f"在目录 {args.input} 中未找到任何图片文件。", file=sys.stderr)
        return 1

    log = make_logger(args.quiet)
    progress = ThrottledProgress(enabled=not args.quiet, prefix="进度")

    added = images_to_pptx(
        args.input, args.output,
        page_size=(args.slide_width, args.slide_height),
        bg_color=bg_color, slide_title=args.title.strip(),
        title_size=args.title_size, title_color=title_color,
        sort_mode=args.sort, sort_order=args.order, fit_mode=args.fit,
        log=log, progress=progress,
    )

    summary = {
        "status": "ok" if added > 0 else "failed",
        "output": os.path.abspath(args.output),
        "images_found": len(images),
        "slides_added": added,
        "page_size_in": [args.slide_width, args.slide_height],
        "fit_mode": args.fit,
    }
    if args.json:
        print(json.dumps(summary, ensure_ascii=False))
    else:
        print(f"完成：{added} 页 → {summary['output']}")
    return 0 if added > 0 else 1


if __name__ == "__main__":
    sys.exit(main())
