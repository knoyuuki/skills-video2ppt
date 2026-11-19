#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从视频/录屏中智能提取 PPT 画面帧（CLI 版）。

移植自 integrated_tool_v4.py 的 VideoFrameExtractor（GUI 版），改为无界面调用。

核心算法：
  1. 按 --skip 跳帧采样，在缩小后的灰度图上计算 SSIM 相似度
  2. 相似度低于阈值 → 判定画面切换，进入延迟缓冲（避开转场动画）
  3. 画面稳定或缓冲满后，按拉普拉斯方差（清晰度）选出最佳帧保存
  4. 可选自动检测/固定像素黑边裁剪
"""

import argparse
import glob
import json
import os
import subprocess
import sys
from pathlib import Path

from _common import (ThrottledProgress, make_logger, parse_crop, parse_time,
                     setup_utf8_stdio)

VIDEO_EXTENSIONS = {".mp4", ".avi", ".mov", ".mkv", ".wmv", ".flv",
                    ".webm", ".m4v", ".ts", ".mpg", ".mpeg"}

# 延迟加载：cv2/numpy 先在子进程中探测，避免损坏的库（如 OpenBLAS
# 初始化失败）在导入阶段直接把本进程中止，导致无法给出可读的错误信息。
cv2 = None
np = None
ssim = None
_DEP_ERRORS = {}


def _probe_import(module):
    """在子进程中试探 import，返回 (是否成功, 错误摘要)。"""
    try:
        proc = subprocess.run(
            [sys.executable, "-c", f"import {module}"],
            capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=120,
        )
    except subprocess.TimeoutExpired:
        return False, "import 超时（120s）"
    if proc.returncode == 0:
        return True, None
    err = (proc.stderr or proc.stdout or "").strip()
    return False, err[-300:] or f"退出码 {proc.returncode}"


def missing_dependencies():
    """加载并返回缺失的必需依赖包名；返回空列表时 cv2/numpy 已就绪。"""
    global cv2, np, ssim
    if cv2 is not None and np is not None:
        return []
    missing = []
    for module, pkg in (("cv2", "opencv-python"), ("numpy", "numpy")):
        ok, err = _probe_import(module)
        if not ok:
            missing.append(pkg)
            _DEP_ERRORS[pkg] = err
    if missing:
        return missing
    import cv2 as _cv2
    import numpy as _np
    cv2, np = _cv2, _np
    try:
        from skimage.metrics import structural_similarity as _ssim
    except ImportError:
        _ssim = None
    ssim = _ssim
    return []


def dependency_errors():
    """返回依赖探测失败的错误摘要（包名 -> 信息）。"""
    return dict(_DEP_ERRORS)


def expand_inputs(patterns):
    """展开视频文件 / 目录 / 通配符，返回去重后的文件列表。"""
    found = []
    for p in patterns:
        if os.path.isdir(p):
            for name in sorted(os.listdir(p)):
                fp = os.path.join(p, name)
                if os.path.isfile(fp) and os.path.splitext(name)[1].lower() in VIDEO_EXTENSIONS:
                    found.append(fp)
        elif any(ch in p for ch in "*?["):
            for m in sorted(glob.glob(p)):
                if os.path.isfile(m):
                    found.append(m)
        elif os.path.isfile(p):
            found.append(p)
        else:
            raise FileNotFoundError(f"找不到视频文件或目录：{p}")
    # 去重并保持顺序
    seen = set()
    result = []
    for f in found:
        key = os.path.abspath(f)
        if key not in seen:
            seen.add(key)
            result.append(f)
    return result


class VideoFrameExtractor:
    """从视频中智能提取 PPT 画面帧。"""

    def __init__(self, config, log_callback=None, progress_callback=None):
        self.input_paths = config["input"]
        self.output_dir = Path(config["output"]).absolute()
        self.sim_threshold = config["similarity"]
        self.min_frame_interval = config["min_interval"]
        self.delay_frames = config["delay_frames"]
        self.stable_check = config["stable_check"]
        self.stable_threshold = config["stable_threshold"]
        self.output_format = config["format"]
        self.image_quality = config["quality"]
        self.skip_frames = config["skip"]
        self.resize_width = config["resize"]
        self.verbose = config["verbose"]
        self.crop_sides = config.get("crop_sides", None)
        self.crop_mode = config.get("crop_mode", "auto")
        self.black_threshold = config.get("black_threshold", 30)
        self.crop_margins = config.get("crop_margins",
                                       {"top": 0, "bottom": 0, "left": 0, "right": 0})
        self.start_seconds = config.get("start_seconds", 0)
        self.end_seconds = config.get("end_seconds", 0)
        self.log_callback = log_callback
        self.progress_callback = progress_callback
        self._stop_flag = False
        self.prev_frame = None
        self.frame_count = 0
        self.ppt_count = 0
        self.last_save_frame = -1
        self.delay_buffer = []
        self.max_delay_buffer = self.delay_frames + 5
        self.errors = []

    def _log(self, msg):
        if self.log_callback:
            self.log_callback(msg)

    def stop(self):
        self._stop_flag = True

    # ------------------------------------------------------------------
    #  黑边裁剪
    # ------------------------------------------------------------------
    def _auto_crop_black_borders(self, frame):
        crop_sides = self.crop_sides
        if not crop_sides:
            return frame
        h, w = frame.shape[:2]

        if self.crop_mode == "fixed":
            margins = self.crop_margins
            top = margins.get("top", 0)
            bottom = h - margins.get("bottom", 0)
            left = margins.get("left", 0)
            right = w - margins.get("right", 0)
            if not crop_sides.get("top", False):
                top = 0
            if not crop_sides.get("bottom", False):
                bottom = h
            if not crop_sides.get("left", False):
                left = 0
            if not crop_sides.get("right", False):
                right = w
            top = max(0, min(top, h - 1))
            bottom = max(top + 1, min(bottom, h))
            left = max(0, min(left, w - 1))
            right = max(left + 1, min(right, w))
            cropped = frame[top:bottom, left:right]
            if self.verbose:
                self._log(f"固定像素裁剪：上{margins.get('top', 0)}下{margins.get('bottom', 0)}"
                          f"左{margins.get('left', 0)}右{margins.get('right', 0)}，"
                          f"裁剪后尺寸：{cropped.shape[1]}x{cropped.shape[0]}")
            return cropped

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY) if len(frame.shape) == 3 else frame
        threshold = self.black_threshold
        row_means = np.mean(gray, axis=1)  # shape (h,)
        top = 0
        bottom = h
        left = 0
        right = w
        if crop_sides.get("top", False):
            hits = np.where(row_means > threshold)[0]
            top = int(hits[0]) if hits.size > 0 else h
        if top >= h:
            self._log("【警告】画面全黑，跳过裁剪")
            return frame
        if crop_sides.get("bottom", False):
            hits = np.where(row_means[top:] > threshold)[0]
            bottom = int(hits[-1]) + 1 + top if hits.size > 0 else top
        if crop_sides.get("left", False) or crop_sides.get("right", False):
            col_means = np.mean(gray[top:bottom, :], axis=0)  # shape (w,)
            if crop_sides.get("left", False):
                hits = np.where(col_means > threshold)[0]
                left = int(hits[0]) if hits.size > 0 else w
            if crop_sides.get("right", False):
                hits = np.where(col_means[left:] > threshold)[0]
                right = int(hits[-1]) + 1 + left if hits.size > 0 else left
        if top >= bottom or left >= right:
            self._log("【警告】黑边检测结果异常（全黑画面），跳过裁剪")
            return frame
        cropped = frame[top:bottom, left:right]
        if self.verbose:
            self._log(f"自动裁剪：上{top}下{h - bottom}左{left}右{w - right}，"
                      f"裁剪后尺寸：{cropped.shape[1]}x{cropped.shape[0]}")
        return cropped

    # ------------------------------------------------------------------
    #  相似度计算
    # ------------------------------------------------------------------
    def _preprocess_frame(self, frame):
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        if self.resize_width > 0 and gray.shape[1] > self.resize_width:
            ratio = self.resize_width / gray.shape[1]
            new_height = int(gray.shape[0] * ratio)
            gray = cv2.resize(gray, (self.resize_width, new_height),
                              interpolation=cv2.INTER_AREA)
        return gray

    def _calculate_similarity(self, frame1, frame2):
        if frame1.shape != frame2.shape:
            frame2 = cv2.resize(frame2, (frame1.shape[1], frame1.shape[0]),
                                interpolation=cv2.INTER_AREA)
        diff = cv2.absdiff(frame1, frame2)
        mean_diff = np.mean(diff)
        if mean_diff < 1.0:
            return 1.0 - mean_diff / 255.0
        if ssim is not None:
            try:
                return float(ssim(frame1, frame2, data_range=255))
            except Exception as e:
                self._log(f"SSIM计算失败：{e}，使用备用方法")
        # 备用方法：统计变化像素占比
        changed_pixels = np.count_nonzero(diff > 30)
        return 1.0 - changed_pixels / frame1.size

    def _is_frame_stable(self, recent_frames):
        if len(recent_frames) < 2:
            return True
        for i in range(1, len(recent_frames)):
            if self._calculate_similarity(recent_frames[i - 1], recent_frames[i]) < self.stable_threshold:
                return False
        return True

    # ------------------------------------------------------------------
    #  保存与延迟缓冲
    # ------------------------------------------------------------------
    def _save_frame(self, frame, video_name, current_frame_num):
        self.ppt_count += 1
        video_base = Path(video_name).stem
        output_filename = (f"{video_base}_ppt_{self.ppt_count:03d}_"
                           f"frame_{current_frame_num:06d}.{self.output_format.lower()}")
        output_path = self.output_dir / output_filename
        try:
            save_frame = self._auto_crop_black_borders(frame)
            output_ext = self.output_format.lower()
            if output_ext in ("jpg", "jpeg"):
                ext_str = ".jpg"
                params = [cv2.IMWRITE_JPEG_QUALITY, self.image_quality,
                          cv2.IMWRITE_JPEG_OPTIMIZE, 1]
            else:
                ext_str = ".png"
                params = [cv2.IMWRITE_PNG_COMPRESSION, 0]
            success, encoded_img = cv2.imencode(ext_str, save_frame, params)
            if success:
                with open(str(output_path), "wb") as f:
                    f.write(encoded_img.tobytes())
                self._log(f"保存PPT页面 {self.ppt_count} | 对应视频帧数：{current_frame_num:,} "
                          f"| 文件：{output_path} "
                          f"({save_frame.shape[1]}x{save_frame.shape[0]})")
            else:
                self._log(f"【错误】编码图片失败！帧数：{current_frame_num:,}")
        except Exception as e:
            self.errors.append(f"保存第 {self.ppt_count} 页失败：{e}")
            self._log(f"【错误】保存PPT页面 {self.ppt_count} 失败！路径：{output_path}，错误：{e}")

    def _process_delayed_frame(self, video_name):
        if not self.delay_buffer:
            return
        if self.verbose:
            self._log(f"处理延迟缓冲区，共有 {len(self.delay_buffer)} 帧待处理")
        if self.stable_check and len(self.delay_buffer) >= 3:
            sharpness_scores = []
            for frame_data in self.delay_buffer:
                gray = cv2.cvtColor(frame_data["frame"], cv2.COLOR_BGR2GRAY)
                sharpness_scores.append(cv2.Laplacian(gray, cv2.CV_64F).var())
            best_index = int(np.argmax(sharpness_scores))
            best_frame_data = self.delay_buffer[best_index]
            if self.verbose:
                self._log(f"清晰度评分：{[f'{s:.1f}' for s in sharpness_scores]}，"
                          f"选择第 {best_index + 1} 帧")
        else:
            best_frame_data = self.delay_buffer[-1]
        self._save_frame(best_frame_data["frame"], video_name, best_frame_data["frame_num"])
        self.prev_frame = self._preprocess_frame(best_frame_data["frame"])
        self.last_save_frame = best_frame_data["frame_num"]
        self.delay_buffer = []

    def _should_skip_frame(self):
        if self.skip_frames <= 0:
            return False
        local = getattr(self, "_local_frame_idx", self.frame_count)
        return local % (self.skip_frames + 1) != 0

    # ------------------------------------------------------------------
    #  单个视频处理
    # ------------------------------------------------------------------
    def process_video(self, video_path):
        self._log(f"\n开始处理视频：{video_path}")
        start_ppt = self.ppt_count
        self.frame_count = 0
        self.prev_frame = None
        self.last_save_frame = -1
        self.delay_buffer = []
        video_name = os.path.basename(video_path)
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.errors.append(f"无法创建输出目录 {self.output_dir}：{e}")
            self._log(f"【错误】无法创建输出目录 {self.output_dir}：{e}")
            return

        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            self.errors.append(f"无法打开视频：{video_path}")
            self._log(f"【错误】无法打开视频：{video_path}")
            return

        try:
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            fps = cap.get(cv2.CAP_PROP_FPS)
            frames_known = total_frames > 0
            fps_ok = bool(fps and fps > 0)

            # === 时间范围处理 ===
            start_frame = 0
            end_frame = total_frames if frames_known else None
            if self.start_seconds > 0:
                if fps_ok:
                    start_frame = int(self.start_seconds * fps)
                    cap.set(cv2.CAP_PROP_POS_FRAMES, start_frame)
                    self._log(f"从第 {start_frame:,} 帧开始（{self.start_seconds:.1f}秒）")
                else:
                    self._log("【警告】无法获取帧率，忽略开始时间设置")
            if self.end_seconds > 0:
                if fps_ok:
                    end_req = int(self.end_seconds * fps)
                    end_frame = min(end_req, total_frames) if frames_known else end_req
                    self._log(f"到第 {end_frame:,} 帧结束（{self.end_seconds:.1f}秒）")
                else:
                    self._log("【警告】无法获取帧率，忽略结束时间设置")
            if not frames_known:
                self._log("【警告】无法获取视频总帧数（可能为损坏文件），将读取到末尾")

            range_frames = max(1, (end_frame - start_frame) if end_frame is not None else 1)
            self.frame_count = start_frame  # 保持帧号连续
            self._local_frame_idx = 0        # 跳帧使用的局部帧号

            fps_desc = f"{fps:.2f}" if fps_ok else "未知"
            if start_frame > 0 or (end_frame is not None and (not frames_known or end_frame < total_frames)):
                ef = end_frame if end_frame is not None else "末尾"
                self._log(f"处理帧范围：{start_frame:,} ~ {ef}，"
                          f"总帧数：{total_frames if frames_known else '未知'}，帧率：{fps_desc}")
            else:
                self._log(f"视频总帧数：{total_frames:,}，帧率：{fps_desc}（全程处理）")

            while cap.isOpened() and not self._stop_flag:
                if end_frame is not None and self.frame_count >= end_frame:
                    break
                ret, frame = cap.read()
                if not ret:
                    break
                if frame is None or frame.size == 0:
                    self._log(f"帧 {self.frame_count + 1} 读取异常（空帧），跳过")
                    self.frame_count += 1
                    self._local_frame_idx += 1
                    continue
                self.frame_count += 1
                self._local_frame_idx += 1

                if self.progress_callback:
                    done = self.frame_count - start_frame
                    if getattr(self, "_total_videos", 1) > 1:
                        overall = int((self._current_video_idx + done / range_frames)
                                      / self._total_videos * 100)
                    else:
                        overall = int(done / range_frames * 100) if range_frames > 0 else 0
                    self.progress_callback(done, range_frames, video_name, overall)

                if self._should_skip_frame():
                    continue

                current_frame_gray = self._preprocess_frame(frame)
                if self.prev_frame is None:
                    self._save_frame(frame, video_name, self.frame_count)
                    self.prev_frame = current_frame_gray
                    self.last_save_frame = self.frame_count
                    continue
                if self.frame_count - self.last_save_frame < self.min_frame_interval:
                    continue

                similarity = self._calculate_similarity(self.prev_frame, current_frame_gray)
                if self.verbose:
                    self._log(f"帧 {self.frame_count:,} - 相似度: {similarity:.3f}")

                if similarity < self.sim_threshold:
                    self.delay_buffer.append({"frame": frame.copy(),
                                              "frame_num": self.frame_count})
                    if len(self.delay_buffer) >= self.max_delay_buffer:
                        if self.verbose:
                            self._log("缓冲区满，立即处理")
                        self._process_delayed_frame(video_name)
                    elif self.verbose:
                        self._log(f"检测到变化，加入延迟缓冲区（当前大小：{len(self.delay_buffer)}）")
                else:
                    if self.delay_buffer and len(self.delay_buffer) >= self.delay_frames:
                        if self.stable_check:
                            gray_frames = [self._preprocess_frame(item["frame"])
                                           for item in self.delay_buffer[-3:]]
                            if self._is_frame_stable(gray_frames):
                                if self.verbose:
                                    self._log("画面稳定，处理延迟缓冲区")
                                self._process_delayed_frame(video_name)
                            elif self.verbose:
                                self._log("画面不稳定，继续等待")
                        else:
                            self._process_delayed_frame(video_name)
                    else:
                        self.prev_frame = current_frame_gray
        finally:
            cap.release()

        if self.delay_buffer and not self._stop_flag:
            self._log("视频结束，处理剩余的延迟缓冲区")
            self._process_delayed_frame(video_name)
        self._log(f"视频处理完成！本视频提取 {self.ppt_count - start_ppt} 页"
                  f"（累计 {self.ppt_count} 页）")

    def run(self):
        """处理所有输入视频，返回本次新增的帧数。"""
        self._log("=" * 60)
        self._log("视频抽帧启动（延迟截图 + 清晰度选帧 + 可选黑边裁剪）")
        self._log("=" * 60)
        start_count = self.ppt_count
        self._total_videos = len(self.input_paths)
        for i, video_path in enumerate(self.input_paths):
            if self._stop_flag:
                break
            self._current_video_idx = i
            try:
                self.process_video(video_path)
            except Exception as e:
                self.errors.append(f"处理 {video_path} 出错：{e}")
                self._log(f"【错误】处理 {video_path} 出错：{e}")
        count = self.ppt_count - start_count
        if self._stop_flag:
            self._log("\n已按请求停止，部分结果可能已保存。")
        else:
            self._log("\n" + "=" * 60)
            self._log(f"所有视频处理完成！共提取 {count} 页")
            self._log(f"文件保存位置：{self.output_dir}")
            self._log("=" * 60)
        return count


# ----------------------------------------------------------------------
#  CLI
# ----------------------------------------------------------------------
def build_parser():
    p = argparse.ArgumentParser(
        description="从视频/录屏中智能提取 PPT 画面帧（每页一张图）。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("-i", "--input", nargs="+", required=True,
                   help="视频文件、目录或通配符（可多个）")
    p.add_argument("-o", "--output", default="frames", help="输出目录")

    g = p.add_argument_group("识别参数")
    g.add_argument("--similarity", type=float, default=0.8,
                   help="相似度阈值(0-1)：低于该值判定画面切换，越低提取的页越少")
    g.add_argument("--min-interval", type=int, default=5, dest="min_interval",
                   help="两次保存之间的最小帧数间隔")
    g.add_argument("--delay-frames", type=int, default=5, dest="delay_frames",
                   help="检测到切换后延迟的帧数（等待转场结束）")
    g.add_argument("--no-stable-check", action="store_false", dest="stable_check",
                   help="关闭画面稳定性检查，缓冲满后直接保存")
    g.add_argument("--stable-threshold", type=float, default=0.95, dest="stable_threshold",
                   help="稳定性阈值(0-1)，越高要求画面越静止")
    g.add_argument("--skip", type=int, default=15,
                   help="跳帧数：每 (n+1) 帧处理 1 帧，0 表示逐帧处理")
    g.add_argument("--resize", type=int, default=480,
                   help="相似度计算前缩放到该宽度，0 表示不缩放（更快）")

    g = p.add_argument_group("时间范围")
    g.add_argument("--start", default="0",
                   help="开始时间：秒 / MM:SS / HH:MM:SS，0 表示从头")
    g.add_argument("--end", default="0",
                   help="结束时间：秒 / MM:SS / HH:MM:SS，0 表示到结尾")

    g = p.add_argument_group("黑边裁剪")
    g.add_argument("--crop", default="none",
                   help="裁剪方向：none / all / 逗号分隔的 top,bottom,left,right")
    g.add_argument("--crop-mode", choices=["auto", "fixed"], default="auto",
                   dest="crop_mode", help="auto=按像素亮度自动检测，fixed=固定边距")
    g.add_argument("--black-threshold", type=int, default=30, dest="black_threshold",
                   help="自动裁剪的亮度阈值（判定黑边的行/列均值上限）")
    g.add_argument("--margin-top", type=int, default=10, dest="margin_top")
    g.add_argument("--margin-bottom", type=int, default=10, dest="margin_bottom")
    g.add_argument("--margin-left", type=int, default=8, dest="margin_left")
    g.add_argument("--margin-right", type=int, default=8, dest="margin_right")

    g = p.add_argument_group("输出")
    g.add_argument("--frame-format", choices=["png", "jpg", "jpeg"], default="png",
                   dest="frame_format", help="截图保存格式")
    g.add_argument("--quality", type=int, default=100, help="JPEG 质量(1-100)")

    g = p.add_argument_group("其他")
    g.add_argument("--verbose", action="store_true", help="输出逐帧相似度等调试信息")
    g.add_argument("--quiet", action="store_true", help="仅输出错误与警告")
    g.add_argument("--json", action="store_true",
                   help="在 stdout 输出 JSON 摘要（日志仍走 stderr）")
    return p


def main(argv=None):
    setup_utf8_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)

    missing = missing_dependencies()
    if missing:
        print(f"缺少或无法加载必需依赖：{', '.join(missing)}。"
              f"请执行：python -m pip install -r requirements.txt", file=sys.stderr)
        for pkg, err in dependency_errors().items():
            last_line = err.strip().splitlines()[-1] if err.strip() else ""
            print(f"  - {pkg}: {last_line}", file=sys.stderr)
        return 3

    if not 0.0 <= args.similarity <= 1.0:
        parser.error("--similarity 必须在 0-1 之间")
    if not 0.0 <= args.stable_threshold <= 1.0:
        parser.error("--stable-threshold 必须在 0-1 之间")
    if not 1 <= args.quality <= 100:
        parser.error("--quality 必须在 1-100 之间")
    if args.skip < 0 or args.resize < 0 or args.min_interval < 0 or args.delay_frames < 0:
        parser.error("--skip/--resize/--min-interval/--delay-frames 不能为负数")
    try:
        start_seconds = parse_time(args.start)
        end_seconds = parse_time(args.end)
        crop_sides = parse_crop(args.crop)
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
    progress = ThrottledProgress(enabled=not args.quiet, prefix="进度")

    def progress_cb(current, total, video_name, overall_pct=None):
        pct = overall_pct if overall_pct is not None else (
            int(current / total * 100) if total else 0)
        progress(pct, f"{video_name} {current:,}/{total:,} 帧")

    config = {
        "input": inputs,
        "output": args.output,
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
        "crop_margins": {"top": args.margin_top, "bottom": args.margin_bottom,
                         "left": args.margin_left, "right": args.margin_right},
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
        "status": "ok" if count > 0 else "empty",
        "output_dir": str(Path(args.output).absolute()),
        "videos": len(inputs),
        "frames_saved": count,
        "frame_format": args.frame_format,
        "time_range": {"start": start_seconds, "end": end_seconds},
        "errors": extractor.errors,
    }
    if args.json:
        print(json.dumps(summary, ensure_ascii=False))
    else:
        print(f"完成：共提取 {count} 张画面，保存于 {summary['output_dir']}")
        if extractor.errors:
            print(f"有 {len(extractor.errors)} 条错误，详见日志。", file=sys.stderr)
    return 0 if count > 0 else 1


if __name__ == "__main__":
    sys.exit(main())
