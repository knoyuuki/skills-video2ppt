#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
视频截图 → PPT → PDF 一体化工具 v4.0
整合三个独立程序的功能，通过页签模式切换：
  Tab 1 - 视频截图：从视频中智能提取PPT画面帧（支持时间范围定位）
  Tab 2 - 转PPT：将图片批量生成PPT演示文稿
  Tab 3 - 转PDF：将图片批量生成PDF文件（直接图片→PDF，无需经过PPT）
功能：
  - 页签模式切换三大功能
  - 视频截图支持指定开始/结束时间（默认全程）
  - 转PDF改为图片直接生成PDF（PIL实现，无需LibreOffice/Office）
  - 每Tab左侧参数设置 + 右侧日志输出
  - 自动保存/恢复所有设置参数
v4.0 优化：
  - 转PDF新增图片压缩（缩放+JPEG重编码），大幅减小PDF体积
  - 可调JPEG质量、最大像素宽度，提供高压缩/均衡/高清三种预设
  - 保留原画质开关，关闭压缩即等同 v3 行为
  - 日志输出压缩前后体积对比与最终PDF大小
"""

import os
import sys
import re
import glob
import time
import json
import io
import threading
import queue
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ===================== 应用路径（兼容 PyInstaller 打包）=====================
def _get_app_dir():
    """获取应用程序所在目录，兼容 PyInstaller 单文件打包"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

APP_DIR = _get_app_dir()
CONFIG_FILE = os.path.join(APP_DIR, "integrated_tool_v4_config.json")

# ===================== 尝试导入依赖 =====================
missing_deps = []

try:
    import cv2
except ImportError:
    missing_deps.append("opencv-python")

try:
    import numpy as np
except ImportError:
    missing_deps.append("numpy")

try:
    from skimage.metrics import structural_similarity as ssim
except ImportError:
    missing_deps.append("scikit-image")

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt
except ImportError:
    missing_deps.append("python-pptx")

try:
    from PIL import Image
except ImportError:
    missing_deps.append("Pillow")

# img2pdf 可选依赖：有则使用流式零内存合并，无则回退 PIL 分批
try:
    import img2pdf
    _HAS_IMG2PDF = True
except ImportError:
    _HAS_IMG2PDF = False

# 图片扩展名统一常量
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tiff"}

if missing_deps:
    if getattr(sys, 'frozen', False):
        # 打包环境：无法自动安装，弹窗提示后退出
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("缺少依赖库",
            "程序运行缺少以下依赖库，请联系开发者获取完整安装包：\n\n" +
            "\n".join(f"  - {dep}" for dep in missing_deps))
        root.destroy()
        sys.exit(1)
    else:
        # 源码环境：尝试自动安装
        print("=" * 60)
        print("检测到缺少以下依赖库：")
        for dep in missing_deps:
            print(f"  - {dep}")
        print("=" * 60)
        print("正在尝试自动安装...")
        for dep in missing_deps:
            os.system(f"pip install {dep} -i https://pypi.tuna.tsinghua.edu.cn/simple")
        print("安装完成，请重新启动程序。")
        input("按回车键退出...")
        sys.exit(0)


# =====================================================================
#  配置管理器（自动保存/恢复设置）
# =====================================================================
class ConfigManager:
    """管理所有设置的保存与加载"""

    DEFAULTS = {
        # 视频截图 Tab
        "v_input_files": [],
        "v_output_dir": "ppt_output_GUI",
        "v_similarity": 0.8,
        "v_min_interval": 5,
        "v_delay_frames": 5,
        "v_stable_check": True,
        "v_stable_threshold": 0.95,
        "v_format": "png",
        "v_quality": 100,
        "v_skip": 15,
        "v_resize": 480,
        "v_verbose": False,
        "v_crop_enable": False,
        "v_crop_mode": "auto",
        "v_crop_top": True,
        "v_crop_bottom": True,
        "v_crop_left": True,
        "v_crop_right": True,
        "v_black_threshold": 30,
        "v_crop_margin_top": 10,
        "v_crop_margin_bottom": 10,
        "v_crop_margin_left": 8,
        "v_crop_margin_right": 8,
        # 视频截图 - 时间范围
        "v_time_enable": False,
        "v_start_h": 0,
        "v_start_m": 0,
        "v_start_s": 0,
        "v_end_h": 0,
        "v_end_m": 30,
        "v_end_s": 0,
        # 转PPT Tab
        "p_image_dir": "ppt_output_GUI",
        "p_ppt_path": "output_presentation.pptx",
        "p_page_w": 16,
        "p_page_h": 9,
        "p_bg_r": 255,
        "p_bg_g": 255,
        "p_bg_b": 255,
        "p_title": "",
        "p_title_size": 24,
        "p_title_r": 0,
        "p_title_g": 0,
        "p_title_b": 0,
        "p_sort_mode": "name",
        "p_sort_order": "asc",
        "p_fit_mode": "stretch",
        # 转PDF Tab（图片→PDF）
        "d_image_dir": "ppt_output_GUI",
        "d_pdf_path": "output_presentation.pdf",
        "d_sort_mode": "name",
        "d_sort_order": "asc",
        "d_dpi": 150,
        "d_enable_compress": True,
        "d_jpeg_quality": 85,
        "d_max_width": 1920,
    }

    def __init__(self, filepath=CONFIG_FILE):
        self.filepath = filepath
        self.data = {}
        self.load()

    def load(self):
        """从文件加载配置"""
        try:
            if os.path.exists(self.filepath):
                with open(self.filepath, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                # 合并默认值（保留旧配置中已有的，补充新增的默认值）
                self.data = {**self.DEFAULTS, **loaded}
            else:
                self.data = dict(self.DEFAULTS)
        except Exception:
            self.data = dict(self.DEFAULTS)

    def save(self):
        """保存配置到文件"""
        try:
            with open(self.filepath, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置文件失败：{e}")

    def get(self, key, default=None):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value

    def apply_from_ui(self, ui_vars_dict):
        """从UI变量字典收集设置"""
        for key, var in ui_vars_dict.items():
            if isinstance(var, (tk.StringVar, tk.IntVar, tk.DoubleVar, tk.BooleanVar)):
                self.data[key] = var.get()
            elif isinstance(var, list):
                self.data[key] = var[:]  # 复制列表

    def apply_to_ui(self, ui_vars_dict, list_widgets=None):
        """将设置应用到UI变量"""
        for key, var in ui_vars_dict.items():
            if key in self.data:
                try:
                    if isinstance(var, tk.StringVar):
                        var.set(str(self.data[key]))
                    elif isinstance(var, tk.IntVar):
                        var.set(int(self.data[key]))
                    elif isinstance(var, tk.DoubleVar):
                        var.set(float(self.data[key]))
                    elif isinstance(var, tk.BooleanVar):
                        var.set(bool(self.data[key]))
                except Exception:
                    pass  # 类型转换失败则跳过
        # 恢复列表控件
        if list_widgets:
            for key, listbox in list_widgets.items():
                if key in self.data and isinstance(self.data[key], list):
                    for item in self.data[key]:
                        listbox.insert(tk.END, os.path.basename(item) if os.path.isfile(item) else item)


# =====================================================================
#  模块一：视频截图（来自 带图形化界面V1 - 副本.py）
# =====================================================================
class VideoFrameExtractor:
    """从视频中智能提取PPT画面帧"""

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
        self.crop_margins = config.get("crop_margins", {"top": 0, "bottom": 0, "left": 0, "right": 0})
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

    def stop(self):
        self._stop_flag = True

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
            if self.verbose and self.log_callback:
                self.log_callback(f"固定像素裁剪：上{margins.get('top',0)}下{margins.get('bottom',0)}左{margins.get('left',0)}右{margins.get('right',0)}，裁剪后尺寸：{cropped.shape[1]}x{cropped.shape[0]}")
            return cropped

        if len(frame.shape) == 3:
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        else:
            gray = frame
        threshold = self.black_threshold
        # 向量化：一次求出所有行/列均值，避免逐行循环
        row_means = np.mean(gray, axis=1)  # shape (h,)
        top = 0
        bottom = h
        left = 0
        right = w
        if crop_sides.get("top", False):
            hits = np.where(row_means > threshold)[0]
            top = int(hits[0]) if hits.size > 0 else h
        if top >= h:
            if self.log_callback:
                self.log_callback("【警告】画面全黑，跳过裁剪")
            return frame
        if crop_sides.get("bottom", False):
            hits = np.where(row_means[top:] > threshold)[0]
            if hits.size > 0:
                bottom = int(hits[-1]) + 1 + top
            else:
                bottom = top
        if crop_sides.get("left", False) or crop_sides.get("right", False):
            col_means = np.mean(gray[top:bottom, :], axis=0)  # shape (w,)
            if crop_sides.get("left", False):
                hits = np.where(col_means > threshold)[0]
                left = int(hits[0]) if hits.size > 0 else w
            if crop_sides.get("right", False):
                hits = np.where(col_means[left:] > threshold)[0]
                if hits.size > 0:
                    right = int(hits[-1]) + 1 + left
                else:
                    right = left
        if top >= bottom or left >= right:
            if self.log_callback:
                self.log_callback("【警告】黑边检测结果异常（全黑画面），跳过裁剪")
            return frame
        cropped = frame[top:bottom, left:right]
        if self.verbose and self.log_callback:
            self.log_callback(f"自动裁剪：上{top}下{h-bottom}左{left}右{w-right}，裁剪后尺寸：{cropped.shape[1]}x{cropped.shape[0]}")
        return cropped

    def _preprocess_frame(self, frame):
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        if self.resize_width > 0 and gray.shape[1] > self.resize_width:
            ratio = self.resize_width / gray.shape[1]
            new_height = int(gray.shape[0] * ratio)
            gray = cv2.resize(gray, (self.resize_width, new_height), interpolation=cv2.INTER_AREA)
        return gray

    def _calculate_similarity(self, frame1, frame2):
        if frame1.shape != frame2.shape:
            frame2 = cv2.resize(frame2, (frame1.shape[1], frame1.shape[0]), interpolation=cv2.INTER_AREA)
        diff = cv2.absdiff(frame1, frame2)
        mean_diff = np.mean(diff)
        if mean_diff < 1.0:
            return 1.0 - mean_diff / 255.0
        try:
            score = ssim(frame1, frame2, data_range=255)
            return score
        except Exception as e:
            if self.log_callback:
                self.log_callback(f"SSIM计算失败：{e}，使用备用方法")
            changed_pixels = np.count_nonzero(diff > 30)
            total_pixels = frame1.size
            change_ratio = changed_pixels / total_pixels
            return 1.0 - change_ratio

    def _is_frame_stable(self, recent_frames):
        if len(recent_frames) < 2:
            return True
        for i in range(1, len(recent_frames)):
            similarity = self._calculate_similarity(recent_frames[i - 1], recent_frames[i])
            if similarity < self.stable_threshold:
                return False
        return True

    def _save_frame(self, frame, video_name, current_frame_num):
        self.ppt_count += 1
        video_base = Path(video_name).stem
        output_filename = f"{video_base}_ppt_{self.ppt_count:03d}_frame_{current_frame_num:06d}.{self.output_format.lower()}"
        output_path = self.output_dir / output_filename
        try:
            save_frame = self._auto_crop_black_borders(frame)
            output_ext = self.output_format.lower()
            if output_ext in ['jpg', 'jpeg']:
                ext_str = '.jpg'
                params = [cv2.IMWRITE_JPEG_QUALITY, self.image_quality, cv2.IMWRITE_JPEG_OPTIMIZE, 1]
            else:
                ext_str = '.png'
                params = [cv2.IMWRITE_PNG_COMPRESSION, 0]
            success, encoded_img = cv2.imencode(ext_str, save_frame, params)
            if success:
                with open(str(output_path), 'wb') as f:
                    f.write(encoded_img.tobytes())
                if self.log_callback:
                    original_size = f"{save_frame.shape[1]}x{save_frame.shape[0]}"
                    self.log_callback(f"保存PPT页面 {self.ppt_count} | 对应视频帧数：{current_frame_num:,} | 文件：{output_path} ({original_size})")
            else:
                if self.log_callback:
                    self.log_callback(f"【错误】编码图片失败！帧数：{current_frame_num:,}")
        except Exception as e:
            if self.log_callback:
                self.log_callback(f"【错误】保存PPT页面 {self.ppt_count} 失败！路径：{output_path}，错误：{str(e)}")

    def _process_delayed_frame(self, video_name):
        if not self.delay_buffer:
            return
        if self.verbose and self.log_callback:
            self.log_callback(f"处理延迟缓冲区，共有 {len(self.delay_buffer)} 帧待处理")
        if self.stable_check and len(self.delay_buffer) >= 3:
            sharpness_scores = []
            for frame_data in self.delay_buffer:
                gray = cv2.cvtColor(frame_data['frame'], cv2.COLOR_BGR2GRAY)
                laplacian_var = cv2.Laplacian(gray, cv2.CV_64F).var()
                sharpness_scores.append(laplacian_var)
            best_index = np.argmax(sharpness_scores)
            best_frame_data = self.delay_buffer[best_index]
            if self.verbose and self.log_callback:
                self.log_callback(f"清晰度评分：{[f'{s:.1f}' for s in sharpness_scores]}，选择第 {best_index + 1} 帧")
        else:
            best_frame_data = self.delay_buffer[-1]
        self._save_frame(best_frame_data['frame'], video_name, best_frame_data['frame_num'])
        self.prev_frame = self._preprocess_frame(best_frame_data['frame'])
        self.last_save_frame = best_frame_data['frame_num']
        self.delay_buffer = []

    def _should_skip_frame(self):
        if self.skip_frames <= 0:
            return False
        # 使用局部帧号，避免时间范围起点不为0时跳帧周期错乱
        local = getattr(self, '_local_frame_idx', self.frame_count)
        return local % (self.skip_frames + 1) != 0

    def process_video(self, video_path):
        if self.log_callback:
            self.log_callback(f"\n开始处理视频：{video_path}")
        start_ppt = self.ppt_count
        self.frame_count = 0
        self.prev_frame = None
        self.last_save_frame = -1
        self.delay_buffer = []
        video_name = os.path.basename(video_path)
        # 输出目录预创建
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            if self.log_callback:
                self.log_callback(f"【错误】无法创建输出目录 {self.output_dir}：{e}")
            return
        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            if self.log_callback:
                self.log_callback(f"无法打开视频：{video_path}")
            return
        try:
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            fps = cap.get(cv2.CAP_PROP_FPS)
            frames_known = total_frames > 0
            fps_ok = fps and fps > 0

            # === 时间范围处理 ===
            start_frame = 0
            end_frame = total_frames if frames_known else None
            if self.start_seconds > 0:
                if fps_ok:
                    start_frame = int(self.start_seconds * fps)
                    cap.set(cv2.CAP_PROP_POS_FRAMES, start_frame)
                    if self.log_callback:
                        self.log_callback(f"从第 {start_frame:,} 帧开始（{self.start_seconds:.1f}秒）")
                elif self.log_callback:
                    self.log_callback("【警告】无法获取帧率，忽略开始时间设置")
            if self.end_seconds > 0:
                if fps_ok:
                    end_req = int(self.end_seconds * fps)
                    if frames_known:
                        end_frame = min(end_req, total_frames)
                    else:
                        end_frame = end_req
                    if self.log_callback:
                        self.log_callback(f"到第 {end_frame:,} 帧结束（{self.end_seconds:.1f}秒）")
                elif self.log_callback:
                    self.log_callback("【警告】无法获取帧率，忽略结束时间设置")
            if not frames_known:
                if self.log_callback:
                    self.log_callback("【警告】无法获取视频总帧数（可能为损坏文件），将读取到末尾")
            range_frames = max(1, (end_frame - start_frame) if end_frame is not None else 1)
            self.frame_count = start_frame  # 保持帧号连续
            self._local_frame_idx = 0  # 用于跳帧的局部帧号

            if self.log_callback:
                if start_frame > 0 or (end_frame is not None and (not frames_known or end_frame < total_frames)):
                    ef = end_frame if end_frame is not None else "末尾"
                    self.log_callback(f"处理帧范围：{start_frame:,} ~ {ef}，总帧数：{total_frames if frames_known else '未知'}，帧率：{fps:.2f if fps_ok else 0:.2f}")
                else:
                    self.log_callback(f"视频总帧数：{total_frames:,}，帧率：{fps:.2f}（全程处理）")
            while cap.isOpened() and not self._stop_flag:
                if end_frame is not None and self.frame_count >= end_frame:
                    break
                ret, frame = cap.read()
                if not ret:
                    break
                if frame is None or frame.size == 0:
                    if self.log_callback:
                        self.log_callback(f"帧 {self.frame_count + 1} 读取异常（空帧），跳过")
                    self.frame_count += 1
                    self._local_frame_idx += 1
                    continue
                self.frame_count += 1
                self._local_frame_idx += 1
                if self.progress_callback:
                    done = self.frame_count - start_frame
                    if hasattr(self, '_total_videos') and self._total_videos > 1:
                        overall_pct = int((self._current_video_idx + done / range_frames) / self._total_videos * 100)
                    else:
                        overall_pct = int(done / range_frames * 100) if range_frames > 0 else 0
                    self.progress_callback(done, range_frames, video_name, overall_pct)
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
                if self.verbose and self.log_callback:
                    self.log_callback(f"帧 {self.frame_count:,} - 相似度: {similarity:.3f}")
                if similarity < self.sim_threshold:
                    self.delay_buffer.append({'frame': frame.copy(), 'frame_num': self.frame_count})
                    if len(self.delay_buffer) >= self.max_delay_buffer:
                        if self.verbose and self.log_callback:
                            self.log_callback("缓冲区满，立即处理")
                        self._process_delayed_frame(video_name)
                    elif self.verbose and self.log_callback:
                        self.log_callback(f"检测到变化，加入延迟缓冲区（当前大小：{len(self.delay_buffer)}）")
                else:
                    if self.delay_buffer and len(self.delay_buffer) >= self.delay_frames:
                        if self.stable_check:
                            gray_frames = [self._preprocess_frame(item['frame']) for item in self.delay_buffer[-3:]]
                            if self._is_frame_stable(gray_frames):
                                if self.verbose and self.log_callback:
                                    self.log_callback("画面稳定，处理延迟缓冲区")
                                self._process_delayed_frame(video_name)
                            elif self.verbose and self.log_callback:
                                self.log_callback("画面不稳定，继续等待")
                        else:
                            self._process_delayed_frame(video_name)
                    elif self.delay_buffer:
                        self.prev_frame = current_frame_gray
                    else:
                        self.prev_frame = current_frame_gray
        finally:
            cap.release()
        if self.delay_buffer and not self._stop_flag:
            if self.log_callback:
                self.log_callback("视频结束，处理剩余的延迟缓冲区")
            self._process_delayed_frame(video_name)
        if self.log_callback:
            self.log_callback(f"视频处理完成！本视频提取 {self.ppt_count - start_ppt} 页（累计 {self.ppt_count} 页）")

    def run(self):
        if self.log_callback:
            self.log_callback("=" * 60)
            self.log_callback("PPT自动截取工具启动（增强版：延迟截图+清晰度优化+黑边裁剪）")
            self.log_callback("=" * 60)
        try:
            total_ppt_count = 0
            self._total_videos = len(self.input_paths)
            for i, video_path in enumerate(self.input_paths):
                if self._stop_flag:
                    break
                self._current_video_idx = i
                ppt_before = self.ppt_count
                self.process_video(video_path)
                new_pages = self.ppt_count - ppt_before
                total_ppt_count += new_pages
            if not self._stop_flag and self.log_callback:
                self.log_callback("\n" + "=" * 60)
                self.log_callback(f"所有视频处理完成！总提取PPT页数：{total_ppt_count}")
                self.log_callback(f"文件保存位置：{self.output_dir.absolute()}")
                self.log_callback("=" * 60)
        except Exception as e:
            if self.log_callback:
                self.log_callback(f"处理过程中出错：{str(e)}")
            raise


# =====================================================================
#  模块二：图片转PPT（来自 topptx.py）
# =====================================================================
def images_to_pptx(image_dir, ppt_save_path, page_size=(16, 9),
                   bg_color=(255, 255, 255), slide_title="",
                   title_font_size=24, title_color=(0, 0, 0),
                   sort_mode="name", sort_order="asc",
                   fit_mode="stretch", stop_check=None,
                   log_callback=None, progress_callback=None):
    """将指定目录下的图片批量生成PPT。
    fit_mode:
      stretch  - 全屏拉伸（原行为，非16:9图片会变形）
      fit      - 保持比例缩放居中留白（不变形）
    stop_check: 可选的无参回调，返回 True 时中止生成
    """
    image_paths = []
    for f in os.listdir(image_dir):
        if os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS:
            image_paths.append(os.path.join(image_dir, f))
    if not image_paths:
        if log_callback:
            log_callback(f"❌ 在目录 {image_dir} 中未找到任何图片文件")
        return 0
    if sort_mode == "name":
        image_paths.sort(key=_natural_key)
    elif sort_mode == "date":
        image_paths.sort(key=lambda x: os.path.getmtime(x))
    if sort_order.lower() == "desc":
        image_paths = image_paths[::-1]
    if log_callback:
        log_callback(f"✅ 成功找到 {len(image_paths)} 张图片（{sort_mode}排序，{sort_order}方向，{fit_mode}适应），开始生成PPT...")
    prs = Presentation()
    slide_width_emu = Inches(page_size[0])
    slide_height_emu = Inches(page_size[1])
    prs.slide_width = slide_width_emu
    prs.slide_height = slide_height_emu
    total = len(image_paths)
    added = 0
    for idx, img_path in enumerate(image_paths, 1):
        if stop_check and stop_check():
            if log_callback:
                log_callback("⚠️ 用户取消，PPT生成中止")
            break
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*bg_color)
        if slide_title:
            title_height = Inches(1.0)
            img_top = title_height
            avail_w = slide_width_emu
            avail_h = slide_height_emu - title_height
            title_box = slide.shapes.add_textbox(left=Inches(0), top=Inches(0), width=slide_width_emu, height=title_height)
            title_frame = title_box.text_frame
            title_frame.text = f"{slide_title} - 第{idx}页"
            title_para = title_frame.paragraphs[0]
            title_para.alignment = PP_ALIGN.CENTER
            title_font = title_para.font
            title_font.size = Pt(title_font_size)
            title_font.color.rgb = RGBColor(*title_color)
            title_font.name = "微软雅黑"
        else:
            img_top = Inches(0)
            avail_w = slide_width_emu
            avail_h = slide_height_emu
        try:
            if fit_mode == "fit":
                # 读取图片原始像素尺寸，按比例缩放居中
                from PIL import Image as _PILImg
                with _PILImg.open(img_path) as _im:
                    iw, ih = _im.size
                box_w_emu = avail_w
                box_h_emu = avail_h
                if iw > 0 and ih > 0:
                    scale = min(box_w_emu / iw, box_h_emu / ih)
                    draw_w = int(iw * scale)
                    draw_h = int(ih * scale)
                    left = int((box_w_emu - draw_w) / 2)
                    top = int(img_top + (box_h_emu - draw_h) / 2)
                    pic = slide.shapes.add_picture(img_path, left=left, top=top, width=draw_w, height=draw_h)
                else:
                    pic = slide.shapes.add_picture(img_path, left=Inches(0), top=img_top, width=avail_w, height=avail_h)
                desc = f"fit模式 {iw}x{ih}"
            else:
                pic = slide.shapes.add_picture(img_path, left=Inches(0), top=img_top, width=avail_w, height=avail_h)
                pic.left = round(pic.left)
                pic.top = round(pic.top)
                pic.width = round(pic.width)
                pic.height = round(pic.height)
                desc = "全屏拉伸"
        except Exception as e:
            if log_callback:
                log_callback(f"⚠️  第{idx}张图片 {os.path.basename(img_path)} 导入失败：{str(e)}")
            continue
        added += 1
        if log_callback:
            log_callback(f"✅ 已添加第{idx}/{total}页：{os.path.basename(img_path)}（{desc}）")
        if progress_callback:
            progress_callback(int(idx / total * 95), f"生成PPT [{idx}/{total}]")
    try:
        save_dir = os.path.dirname(os.path.abspath(ppt_save_path))
        if save_dir:
            os.makedirs(save_dir, exist_ok=True)
        prs.save(ppt_save_path)
        if progress_callback:
            progress_callback(100, "PPT生成完成")
        if log_callback:
            log_callback(f"\n🎉 PPT生成成功！保存路径：{os.path.abspath(ppt_save_path)}")
        return added
    except Exception as e:
        if log_callback:
            log_callback(f"❌ PPT保存失败：{str(e)}")
        return 0


# =====================================================================
#  模块三：图片转PDF（PIL 实现，无需 LibreOffice/Office）
# =====================================================================
def _natural_key(path):
    """自然排序键：page_2 < page_10，而非字典序 page_10 < page_2"""
    return [int(t) if t.isdigit() else t.lower()
            for t in re.split(r'(\d+)', os.path.basename(path))]


def _human_size(num):
    """将字节数转为人类可读大小字符串"""
    num = float(num)
    for unit in ['B', 'KB', 'MB', 'GB']:
        if num < 1024.0:
            return f"{num:.1f}{unit}"
        num /= 1024.0
    return f"{num:.1f}TB"


def images_to_pdf(image_dir, pdf_save_path, sort_mode="name", sort_order="asc",
                  dpi=150, enable_compress=True, jpeg_quality=85, max_width=1920,
                  log_callback=None, progress_callback=None):
    """将指定目录下的图片批量生成PDF（每张图片一页）。
    v4.0 优化要点：
      - enable_compress=True：图片统一缩放(max_width)+JPEG重编码(jpeg_quality)后嵌入，体积大幅缩小
        · 优先 img2pdf 流式合并压缩后的JPEG（零重编码、无损嵌入）
        · 无 img2pdf 时回退 Pillow 逐页 append
      - enable_compress=False：原画质路线（img2pdf 或 Pillow），等同 v3 行为
      - dpi 控制 PDF 页面物理尺寸声明，不影响压缩后图片的像素
    """
    image_paths = []
    for f in os.listdir(image_dir):
        if os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS:
            image_paths.append(os.path.join(image_dir, f))
    if not image_paths:
        if log_callback:
            log_callback(f"❌ 在目录 {image_dir} 中未找到任何图片文件")
        return 0
    if sort_mode == "name":
        image_paths.sort(key=_natural_key)
    elif sort_mode == "date":
        image_paths.sort(key=lambda x: os.path.getmtime(x))
    if sort_order.lower() == "desc":
        image_paths = image_paths[::-1]
    total = len(image_paths)
    save_dir = os.path.dirname(os.path.abspath(pdf_save_path))
    if save_dir:
        os.makedirs(save_dir, exist_ok=True)

    # ============================================================
    #  路线一：不启用压缩 —— 原画质直接嵌入（等同 v3）
    # ============================================================
    if not enable_compress:
        engine = "img2pdf(流式原画质)" if _HAS_IMG2PDF else "Pillow(逐页原画质)"
        if log_callback:
            log_callback(f"✅ 找到 {total} 张图片（{sort_mode}排序，{sort_order}方向），压缩已关闭，使用 {engine} 生成PDF...")
        if _HAS_IMG2PDF:
            try:
                with open(pdf_save_path, "wb") as outf:
                    outf.write(img2pdf.convert(image_paths, dpi=dpi))
                if log_callback:
                    log_callback(f"\n🎉 PDF生成成功！共 {total} 页（原画质），路径：{os.path.abspath(pdf_save_path)}")
                    log_callback(f"📄 文件大小：{_human_size(os.path.getsize(pdf_save_path))}")
                if progress_callback:
                    progress_callback(100, "PDF生成完成")
                return total
            except Exception as e:
                if log_callback:
                    log_callback(f"⚠️ img2pdf 失败({str(e)})，回退到 Pillow 逐页模式...")
        written = 0
        for idx, img_path in enumerate(image_paths, 1):
            try:
                with Image.open(img_path) as img:
                    img.load()
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    img.save(pdf_save_path, "PDF", resolution=float(dpi),
                             save_all=(idx == 1), append=(idx > 1))
                written += 1
                if log_callback:
                    log_callback(f"✅ 已写入第{idx}/{total}页：{os.path.basename(img_path)}")
                if progress_callback:
                    progress_callback(int(idx / total * 100), f"写入PDF [{idx}/{total}]")
            except Exception as e:
                if log_callback:
                    log_callback(f"⚠️  第{idx}张图片 {os.path.basename(img_path)} 处理失败：{str(e)}")
                continue
        if written == 0:
            if log_callback:
                log_callback("❌ 没有成功写入任何图片，PDF生成失败")
            return 0
        if log_callback:
            log_callback(f"\n🎉 PDF生成成功！共 {written} 页（原画质），路径：{os.path.abspath(pdf_save_path)}")
            log_callback(f"📄 文件大小：{_human_size(os.path.getsize(pdf_save_path))}")
        return written

    # ============================================================
    #  路线二：启用压缩 —— 缩放 + JPEG 重编码
    # ============================================================
    if log_callback:
        w_desc = f"{max_width}px" if max_width > 0 else "不限"
        log_callback(f"✅ 找到 {total} 张图片，启用压缩（JPEG质量 {jpeg_quality}，最大宽 {w_desc}，DPI {dpi}）...")

    quality = max(10, min(100, int(jpeg_quality)))
    compressed = []   # [(BytesIO, basename, orig_bytes, comp_bytes), ...]
    orig_total = 0
    comp_total = 0
    for idx, img_path in enumerate(image_paths, 1):
        try:
            orig_size = os.path.getsize(img_path)
            with Image.open(img_path) as img:
                img = img.convert('RGB')
                # 按最大宽度等比缩放
                if max_width > 0 and img.width > max_width:
                    ratio = max_width / img.width
                    new_h = max(1, int(img.height * ratio))
                    img = img.resize((max_width, new_h), Image.LANCZOS)
                # JPEG 重编码到内存
                buf = io.BytesIO()
                img.save(buf, format='JPEG', quality=quality, optimize=True)
                comp_size = buf.tell()
                buf.seek(0)
                compressed.append((buf, os.path.basename(img_path), orig_size, comp_size))
                orig_total += orig_size
                comp_total += comp_size
                if log_callback:
                    log_callback(f"✅ 压缩 [{idx}/{total}] {os.path.basename(img_path)}："
                                 f"{_human_size(orig_size)} → {_human_size(comp_size)} "
                                 f"({img.width}x{img.height})")
                if progress_callback:
                    progress_callback(int(idx / total * 80), f"压缩图片 [{idx}/{total}]")
        except Exception as e:
            if log_callback:
                log_callback(f"⚠️  第{idx}张图片 {os.path.basename(img_path)} 压缩失败：{str(e)}")
            continue

    if not compressed:
        if log_callback:
            log_callback("❌ 没有成功处理任何图片，PDF生成失败")
        return 0

    if log_callback and orig_total > 0:
        save_pct = (1 - comp_total / orig_total) * 100
        log_callback(f"📊 压缩汇总：源图 {_human_size(orig_total)} → 压缩后 {_human_size(comp_total)}（节省 {save_pct:.1f}%）")

    written = 0
    # 优先 img2pdf 合并压缩后的 JPEG（对 JPEG 直接嵌入、不重编码）
    if _HAS_IMG2PDF:
        try:
            pdf_bytes = img2pdf.convert(*[buf for buf, _, _, _ in compressed], dpi=dpi)
            with open(pdf_save_path, "wb") as outf:
                outf.write(pdf_bytes)
            written = len(compressed)
            if progress_callback:
                progress_callback(100, "PDF生成完成")
        except Exception as e:
            if log_callback:
                log_callback(f"⚠️ img2pdf 合并失败({str(e)})，回退 Pillow 逐页写入...")
            written = 0

    # 回退：Pillow 逐页 append
    if written == 0:
        for idx, (buf, name, _, _) in enumerate(compressed, 1):
            buf.seek(0)
            try:
                with Image.open(buf) as img:
                    img.load()
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    img.save(pdf_save_path, "PDF", resolution=float(dpi),
                             save_all=(idx == 1), append=(idx > 1))
                written += 1
                if progress_callback:
                    progress_callback(80 + int(idx / len(compressed) * 20), f"写入PDF [{idx}/{len(compressed)}]")
            except Exception as e:
                if log_callback:
                    log_callback(f"⚠️  第{idx}页 {name} 写入PDF失败：{str(e)}")
                continue

    if written == 0:
        if log_callback:
            log_callback("❌ PDF写入失败，未生成有效文件")
        return 0

    if log_callback:
        final_size = os.path.getsize(pdf_save_path) if os.path.exists(pdf_save_path) else 0
        log_callback(f"\n🎉 PDF生成成功！共 {written} 页（压缩模式），路径：{os.path.abspath(pdf_save_path)}")
        log_callback(f"📄 最终PDF大小：{_human_size(final_size)}")
    return written


# =====================================================================
#  主GUI应用程序
# =====================================================================
class IntegratedApp:
    def __init__(self, root):
        self.root = root
        self.root.title("视频截图 → PPT → PDF 一体化工具 v4.0")
        self.root.geometry("1300x850")
        self.root.minsize(1100, 700)

        # ---- 配置管理器 ----
        self.cfg = ConfigManager()

        # ---- 公用变量 ----
        self.running = False
        self._stop_flag = False
        self.worker_thread = None
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.after_id = None
        self._drain_count = 0
        self._max_drain = 5

        # ---- 构建 UI ----
        self._build_menu()
        self._build_main_ui()

        # 启动日志轮询
        self.after_id = self.root.after(100, self._process_queues)

        # 绑定窗口关闭
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="保存当前设置", command=self._manual_save_config)
        file_menu.add_command(label="关于", command=self._show_about)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self._on_closing)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self._show_help)

    def _show_about(self):
        messagebox.showinfo("关于",
            "视频截图 → PPT → PDF 一体化工具 v4.0\n\n"
            "整合三大功能：\n"
            "1. 视频截图 - 智能提取视频中的PPT画面帧（支持时间范围定位）\n"
            "2. 转PPT - 图片批量生成PPT（支持拉伸/留白两种适应模式）\n"
            "3. 转PDF - 图片直接生成PDF（支持图片压缩，可调DPI，无需Office/LibreOffice）\n\n"
            "v3.0 优化：流式PDF内存优化、多视频全局进度、线程安全退出\n"
            "v4.0 优化：转PDF新增图片压缩（缩放+JPEG），大幅减小PDF体积\n"
            "支持自动保存/恢复所有设置参数\n"
            "配置文件：integrated_tool_v4_config.json")

    def _show_help(self):
        messagebox.showinfo("使用说明",
            "【使用说明】\n\n"
            "=== Tab 1: 视频截图 ===\n"
            "1. 点击「选择视频文件」添加视频\n"
            "2. 设置输出目录和识别参数\n"
            "3. 如需只截取部分时间段，勾选「仅截取部分时间段」并设置起止时间\n"
            "4. 点击「开始处理」提取PPT画面帧\n\n"
            "=== Tab 2: 转PPT ===\n"
            "1. 点击「选择图片文件夹」\n"
            "2. 设置PPT保存路径和页面参数\n"
            "3. 点击「开始生成」创建PPT\n\n"
            "=== Tab 3: 转PDF ===\n"
            "1. 点击「选择图片文件夹」（含图片）\n"
            "2. 设置PDF保存路径\n"
            "3. 勾选「启用图片压缩」可大幅减小PDF体积（可调JPEG质量与最大宽度）\n"
            "4. 点击「开始生成PDF」直接生成PDF（每张图片一页）\n\n"
            "所有设置会在退出时自动保存，下次启动自动恢复。")

    def _manual_save_config(self):
        """手动保存当前设置"""
        self._save_all_config()
        messagebox.showinfo("提示", "设置已保存！")

    def _save_all_config(self):
        """收集所有UI设置并保存"""
        # 收集视频截图Tab设置
        v_vars = {
            "v_input_files": self.v_input_files,
            "v_output_dir": self.v_output_dir,
            "v_similarity": self.v_similarity,
            "v_min_interval": self.v_min_interval,
            "v_delay_frames": self.v_delay_frames,
            "v_stable_check": self.v_stable_check,
            "v_stable_threshold": self.v_stable_threshold,
            "v_format": self.v_format,
            "v_quality": self.v_quality,
            "v_skip": self.v_skip,
            "v_resize": self.v_resize,
            "v_verbose": self.v_verbose,
            "v_crop_enable": self.v_crop_enable,
            "v_crop_mode": self.v_crop_mode,
            "v_crop_top": self.v_crop_top,
            "v_crop_bottom": self.v_crop_bottom,
            "v_crop_left": self.v_crop_left,
            "v_crop_right": self.v_crop_right,
            "v_black_threshold": self.v_black_threshold,
            "v_crop_margin_top": self.v_crop_margin_top,
            "v_crop_margin_bottom": self.v_crop_margin_bottom,
            "v_crop_margin_left": self.v_crop_margin_left,
            "v_crop_margin_right": self.v_crop_margin_right,
            "v_time_enable": self.v_time_enable,
            "v_start_h": self.v_start_h,
            "v_start_m": self.v_start_m,
            "v_start_s": self.v_start_s,
            "v_end_h": self.v_end_h,
            "v_end_m": self.v_end_m,
            "v_end_s": self.v_end_s,
        }
        self.cfg.apply_from_ui(v_vars)

        # 收集转PPT Tab设置
        p_vars = {
            "p_image_dir": self.p_image_dir,
            "p_ppt_path": self.p_ppt_path,
            "p_page_w": self.p_page_w,
            "p_page_h": self.p_page_h,
            "p_bg_r": self.p_bg_r,
            "p_bg_g": self.p_bg_g,
            "p_bg_b": self.p_bg_b,
            "p_title": self.p_title,
            "p_title_size": self.p_title_size,
            "p_title_r": self.p_title_r,
            "p_title_g": self.p_title_g,
            "p_title_b": self.p_title_b,
            "p_sort_mode": self.p_sort_mode,
            "p_sort_order": self.p_sort_order,
            "p_fit_mode": self.p_fit_mode,
        }
        self.cfg.apply_from_ui(p_vars)

        # 收集转PDF Tab设置
        d_vars = {
            "d_image_dir": self.d_image_dir,
            "d_pdf_path": self.d_pdf_path,
            "d_sort_mode": self.d_sort_mode,
            "d_sort_order": self.d_sort_order,
            "d_dpi": self.d_dpi,
            "d_enable_compress": self.d_enable_compress,
            "d_jpeg_quality": self.d_jpeg_quality,
            "d_max_width": self.d_max_width,
        }
        self.cfg.apply_from_ui(d_vars)

        self.cfg.save()

    def _load_all_config(self):
        """从配置文件恢复所有UI设置"""
        # 恢复视频截图设置
        v_vars = {
            "v_output_dir": self.v_output_dir,
            "v_similarity": self.v_similarity,
            "v_min_interval": self.v_min_interval,
            "v_delay_frames": self.v_delay_frames,
            "v_stable_check": self.v_stable_check,
            "v_stable_threshold": self.v_stable_threshold,
            "v_format": self.v_format,
            "v_quality": self.v_quality,
            "v_skip": self.v_skip,
            "v_resize": self.v_resize,
            "v_verbose": self.v_verbose,
            "v_crop_enable": self.v_crop_enable,
            "v_crop_mode": self.v_crop_mode,
            "v_crop_top": self.v_crop_top,
            "v_crop_bottom": self.v_crop_bottom,
            "v_crop_left": self.v_crop_left,
            "v_crop_right": self.v_crop_right,
            "v_black_threshold": self.v_black_threshold,
            "v_crop_margin_top": self.v_crop_margin_top,
            "v_crop_margin_bottom": self.v_crop_margin_bottom,
            "v_crop_margin_left": self.v_crop_margin_left,
            "v_crop_margin_right": self.v_crop_margin_right,
            "v_time_enable": self.v_time_enable,
            "v_start_h": self.v_start_h,
            "v_start_m": self.v_start_m,
            "v_start_s": self.v_start_s,
            "v_end_h": self.v_end_h,
            "v_end_m": self.v_end_m,
            "v_end_s": self.v_end_s,
        }
        self.cfg.apply_to_ui(v_vars)
        # 恢复视频文件列表
        saved_files = self.cfg.get("v_input_files", [])
        for f in saved_files:
            if os.path.isfile(f):
                self.v_input_files.append(f)
                self.v_listbox.insert(tk.END, os.path.basename(f))

        # 恢复转PPT设置
        p_vars = {
            "p_image_dir": self.p_image_dir,
            "p_ppt_path": self.p_ppt_path,
            "p_page_w": self.p_page_w,
            "p_page_h": self.p_page_h,
            "p_bg_r": self.p_bg_r,
            "p_bg_g": self.p_bg_g,
            "p_bg_b": self.p_bg_b,
            "p_title": self.p_title,
            "p_title_size": self.p_title_size,
            "p_title_r": self.p_title_r,
            "p_title_g": self.p_title_g,
            "p_title_b": self.p_title_b,
            "p_sort_mode": self.p_sort_mode,
            "p_sort_order": self.p_sort_order,
            "p_fit_mode": self.p_fit_mode,
        }
        self.cfg.apply_to_ui(p_vars)

        # 恢复转PDF设置
        d_vars = {
            "d_image_dir": self.d_image_dir,
            "d_pdf_path": self.d_pdf_path,
            "d_sort_mode": self.d_sort_mode,
            "d_sort_order": self.d_sort_order,
            "d_dpi": self.d_dpi,
            "d_enable_compress": self.d_enable_compress,
            "d_jpeg_quality": self.d_jpeg_quality,
            "d_max_width": self.d_max_width,
        }
        self.cfg.apply_to_ui(d_vars)

        # 更新UI依赖状态
        self._v_update_stable()
        self._v_update_crop()
        self._v_update_crop_mode()
        self._v_update_time()
        self._d_update_compress()

    def _build_main_ui(self):
        # 主容器
        main_frame = ttk.Frame(self.root, padding=6)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ===== 页签 =====
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 创建三个页签
        self._build_tab_video(self.notebook)
        self._build_tab_pptx(self.notebook)
        self._build_tab_pdf(self.notebook)

        # ===== 底部：全局进度条 + 状态栏 =====
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(4, 0))

        progress_frame = ttk.Frame(bottom_frame)
        progress_frame.pack(fill=tk.X)
        ttk.Label(progress_frame, text="全局进度：").pack(side=tk.LEFT)
        self.global_progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.global_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.global_progress_label = ttk.Label(progress_frame, text="0%")
        self.global_progress_label.pack(side=tk.RIGHT)

        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(bottom_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(2, 0))

        # 加载保存的设置
        self._load_all_config()

    # ==================================================================
    #  通用：创建左右分栏的页签
    # ==================================================================
    def _create_tab_layout(self, notebook, text):
        """创建一个左右分栏的页签：左侧参数区，右侧日志区"""
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=text)

        # 使用 PanedWindow 实现左右分栏
        pw = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        pw.pack(fill=tk.BOTH, expand=True)

        # 左：参数设置区（可滚动）
        left_frame = ttk.Frame(pw)
        pw.add(left_frame, weight=2)

        canvas = tk.Canvas(left_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas)
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        def _on_mw(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        def _bind_mw(event):
            canvas.bind_all("<MouseWheel>", _on_mw)
        def _unbind_mw(event):
            canvas.unbind_all("<MouseWheel>")
        canvas.bind("<Enter>", _bind_mw)
        canvas.bind("<Leave>", _unbind_mw)

        # 右：日志区
        right_frame = ttk.Frame(pw)
        pw.add(right_frame, weight=1)

        log_frame = ttk.LabelFrame(right_frame, text="📋 处理日志", padding=4)
        log_frame.pack(fill=tk.BOTH, expand=True)

        log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=('Consolas', 9),
                                              bg='#1e1e1e', fg='#d4d4d4', insertbackground='white')
        log_text.pack(fill=tk.BOTH, expand=True)
        log_text.tag_config('info', foreground='#d4d4d4')
        log_text.tag_config('success', foreground='#4ec9b0')
        log_text.tag_config('error', foreground='#f44747')
        log_text.tag_config('warning', foreground='#dcdcaa')
        log_text.tag_config('title', foreground='#569cd6', font=('Consolas', 10, 'bold'))
        log_text.tag_config('progress', foreground='#c586c0')

        return tab, scrollable, log_text

    def _log_to_widget(self, log_widget, msg):
        """将消息写入日志控件，自动判断颜色标签"""
        tag = 'info'
        if msg.startswith("=") or "开始" in msg or "结束" in msg or "完成" in msg:
            tag = 'title'
        elif "错误" in msg or "失败" in msg or "出错" in msg:
            tag = 'error'
        elif "保存" in msg or "成功" in msg or "OK" in msg or "🎉" in msg or "✅" in msg:
            tag = 'success'
        elif "警告" in msg or "⚠" in msg:
            tag = 'warning'
        elif "进度" in msg or "%" in msg:
            tag = 'progress'
        log_widget.insert(tk.END, msg + "\n", tag)
        log_widget.see(tk.END)

    # ==================================================================
    #  Tab 1: 视频截图
    # ==================================================================
    def _build_tab_video(self, notebook):
        tab, scrollable, self.v_log = self._create_tab_layout(notebook, "  视频截图  ")

        # ---- 变量 ----
        self.v_input_files = []
        self.v_output_dir = tk.StringVar(value="ppt_output_GUI")
        self.v_similarity = tk.DoubleVar(value=0.8)
        self.v_min_interval = tk.IntVar(value=5)
        self.v_delay_frames = tk.IntVar(value=5)
        self.v_stable_check = tk.BooleanVar(value=True)
        self.v_stable_threshold = tk.DoubleVar(value=0.95)
        self.v_format = tk.StringVar(value="png")
        self.v_quality = tk.IntVar(value=100)
        self.v_skip = tk.IntVar(value=15)
        self.v_resize = tk.IntVar(value=480)
        self.v_verbose = tk.BooleanVar(value=False)
        self.v_crop_enable = tk.BooleanVar(value=False)
        self.v_crop_mode = tk.StringVar(value="auto")
        self.v_crop_top = tk.BooleanVar(value=True)
        self.v_crop_bottom = tk.BooleanVar(value=True)
        self.v_crop_left = tk.BooleanVar(value=True)
        self.v_crop_right = tk.BooleanVar(value=True)
        self.v_black_threshold = tk.IntVar(value=30)
        self.v_crop_margin_top = tk.IntVar(value=10)
        self.v_crop_margin_bottom = tk.IntVar(value=10)
        self.v_crop_margin_left = tk.IntVar(value=8)
        self.v_crop_margin_right = tk.IntVar(value=8)
        # 时间范围
        self.v_time_enable = tk.BooleanVar(value=False)
        self.v_start_h = tk.IntVar(value=0)
        self.v_start_m = tk.IntVar(value=0)
        self.v_start_s = tk.IntVar(value=0)
        self.v_end_h = tk.IntVar(value=0)
        self.v_end_m = tk.IntVar(value=30)
        self.v_end_s = tk.IntVar(value=0)

        # ---- 文件选择 ----
        f1 = ttk.LabelFrame(scrollable, text="📁 文件选择", padding=8)
        f1.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f1, text="输入视频：").grid(row=0, column=0, sticky="w", pady=2)
        bf = ttk.Frame(f1)
        bf.grid(row=0, column=1, sticky="w", pady=2)
        self.v_select_btn = ttk.Button(bf, text="选择视频文件", command=self._v_select_files)
        self.v_select_btn.pack(side=tk.LEFT, padx=2)
        self.v_clear_btn = ttk.Button(bf, text="清空列表", command=self._v_clear_files)
        self.v_clear_btn.pack(side=tk.LEFT, padx=2)
        self.v_listbox = tk.Listbox(f1, height=4)
        self.v_listbox.grid(row=1, column=0, columnspan=2, sticky="ew", pady=2)
        f1.columnconfigure(1, weight=1)
        ttk.Label(f1, text="输出目录：").grid(row=2, column=0, sticky="w", pady=2)
        df = ttk.Frame(f1)
        df.grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Entry(df, textvariable=self.v_output_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(df, text="浏览...", command=lambda: self.v_output_dir.set(filedialog.askdirectory(title="选择输出目录") or self.v_output_dir.get())).pack(side=tk.RIGHT)

        # ---- 识别参数 ----
        f2 = ttk.LabelFrame(scrollable, text="🔍 识别参数", padding=8)
        f2.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f2, text="相似度阈值：").grid(row=0, column=0, sticky="w", pady=3)
        sf = ttk.Frame(f2)
        sf.grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Scale(sf, from_=0.5, to=1.0, variable=self.v_similarity, orient=tk.HORIZONTAL,
                  command=lambda v: self.v_sim_label.config(text=f"{float(v):.2f}")).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.v_sim_label = ttk.Label(sf, text="0.80", width=5)
        self.v_sim_label.pack(side=tk.RIGHT, padx=5)
        ttk.Label(f2, text="最小帧间隔：").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Spinbox(f2, from_=0, to=100, textvariable=self.v_min_interval, width=8).grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(f2, text="跳帧数：").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Spinbox(f2, from_=0, to=100, textvariable=self.v_skip, width=8).grid(row=2, column=1, sticky="w", pady=3)
        ttk.Label(f2, text="缩放宽度：").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Spinbox(f2, from_=0, to=1920, textvariable=self.v_resize, width=8).grid(row=3, column=1, sticky="w", pady=3)
        f2.columnconfigure(1, weight=1)

        # ---- 时间范围 ----
        f2b = ttk.LabelFrame(scrollable, text="⏰ 时间范围（默认全程，勾选后生效）", padding=8)
        f2b.pack(fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(f2b, text="仅截取部分时间段", variable=self.v_time_enable, command=self._v_update_time).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 5))
        self.v_time_start_label = ttk.Label(f2b, text="开始时间：")
        self.v_time_start_label.grid(row=1, column=0, sticky="w", pady=3)
        sb = ttk.Frame(f2b)
        sb.grid(row=1, column=1, columnspan=3, sticky="w", pady=3)
        self.v_start_h_sb = ttk.Spinbox(sb, from_=0, to=23, textvariable=self.v_start_h, width=4)
        self.v_start_h_sb.pack(side=tk.LEFT, padx=2)
        ttk.Label(sb, text="时").pack(side=tk.LEFT)
        self.v_start_m_sb = ttk.Spinbox(sb, from_=0, to=59, textvariable=self.v_start_m, width=4)
        self.v_start_m_sb.pack(side=tk.LEFT, padx=2)
        ttk.Label(sb, text="分").pack(side=tk.LEFT)
        self.v_start_s_sb = ttk.Spinbox(sb, from_=0, to=59, textvariable=self.v_start_s, width=4)
        self.v_start_s_sb.pack(side=tk.LEFT, padx=2)
        ttk.Label(sb, text="秒").pack(side=tk.LEFT)
        self.v_time_end_label = ttk.Label(f2b, text="结束时间：")
        self.v_time_end_label.grid(row=2, column=0, sticky="w", pady=3)
        eb = ttk.Frame(f2b)
        eb.grid(row=2, column=1, columnspan=3, sticky="w", pady=3)
        self.v_end_h_sb = ttk.Spinbox(eb, from_=0, to=23, textvariable=self.v_end_h, width=4)
        self.v_end_h_sb.pack(side=tk.LEFT, padx=2)
        ttk.Label(eb, text="时").pack(side=tk.LEFT)
        self.v_end_m_sb = ttk.Spinbox(eb, from_=0, to=59, textvariable=self.v_end_m, width=4)
        self.v_end_m_sb.pack(side=tk.LEFT, padx=2)
        ttk.Label(eb, text="分").pack(side=tk.LEFT)
        self.v_end_s_sb = ttk.Spinbox(eb, from_=0, to=59, textvariable=self.v_end_s, width=4)
        self.v_end_s_sb.pack(side=tk.LEFT, padx=2)
        ttk.Label(eb, text="秒").pack(side=tk.LEFT)
        ttk.Label(f2b, text="提示：不勾选则处理视频全程；勾选后按时间截取片段。", foreground="gray").grid(row=3, column=0, columnspan=4, sticky="w", pady=(3, 0))
        self._v_update_time()

        # ---- 延迟截图参数 ----
        f3 = ttk.LabelFrame(scrollable, text="⏱ 延迟截图参数", padding=8)
        f3.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f3, text="延迟帧数：").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Spinbox(f3, from_=0, to=50, textvariable=self.v_delay_frames, width=8).grid(row=0, column=1, sticky="w", pady=3)
        ttk.Label(f3, text="稳定性检查：").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Checkbutton(f3, text="启用", variable=self.v_stable_check, command=self._v_update_stable).grid(row=1, column=1, sticky="w", pady=3)
        self.v_stable_label = ttk.Label(f3, text="稳定性阈值：")
        self.v_stable_label.grid(row=2, column=0, sticky="w", pady=3)
        stf = ttk.Frame(f3)
        stf.grid(row=2, column=1, sticky="ew", pady=3)
        self.v_stable_scale = ttk.Scale(stf, from_=0.8, to=1.0, variable=self.v_stable_threshold, orient=tk.HORIZONTAL,
                  command=lambda v: self.v_stable_val_label.config(text=f"{float(v):.2f}"))
        self.v_stable_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.v_stable_val_label = ttk.Label(stf, text="0.95", width=5)
        self.v_stable_val_label.pack(side=tk.RIGHT, padx=5)
        f3.columnconfigure(1, weight=1)
        self._v_update_stable()

        # ---- 黑边裁剪 ----
        f4 = ttk.LabelFrame(scrollable, text="✂️ 黑边裁剪", padding=8)
        f4.pack(fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(f4, text="启用黑边裁剪", variable=self.v_crop_enable, command=self._v_update_crop).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 5))
        ttk.Label(f4, text="裁剪模式：").grid(row=1, column=0, sticky="w", pady=3)
        mf = ttk.Frame(f4)
        mf.grid(row=1, column=1, columnspan=3, sticky="w", pady=3)
        ttk.Radiobutton(mf, text="自动检测", variable=self.v_crop_mode, value="auto", command=self._v_update_crop_mode).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(mf, text="固定像素", variable=self.v_crop_mode, value="fixed", command=self._v_update_crop_mode).pack(side=tk.LEFT, padx=2)
        ttk.Label(f4, text="裁剪方向：").grid(row=2, column=0, sticky="w", pady=3)
        self.v_crop_top_cb = ttk.Checkbutton(f4, text="上边", variable=self.v_crop_top)
        self.v_crop_top_cb.grid(row=2, column=1, sticky="w", pady=2)
        self.v_crop_bottom_cb = ttk.Checkbutton(f4, text="下边", variable=self.v_crop_bottom)
        self.v_crop_bottom_cb.grid(row=2, column=2, sticky="w", pady=2)
        self.v_crop_left_cb = ttk.Checkbutton(f4, text="左边", variable=self.v_crop_left)
        self.v_crop_left_cb.grid(row=3, column=1, sticky="w", pady=2)
        self.v_crop_right_cb = ttk.Checkbutton(f4, text="右边", variable=self.v_crop_right)
        self.v_crop_right_cb.grid(row=3, column=2, sticky="w", pady=2)

        # 自动参数
        self.v_auto_frame = ttk.Frame(f4)
        self.v_auto_frame.grid(row=4, column=0, columnspan=4, sticky="ew", pady=2)
        ttk.Label(self.v_auto_frame, text="黑边阈值：").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Scale(self.v_auto_frame, from_=1, to=80, variable=self.v_black_threshold, orient=tk.HORIZONTAL,
                  command=lambda v: self.v_bt_label.config(text=f"{int(float(v))}")).grid(row=0, column=1, sticky="ew", pady=3, padx=5)
        self.v_bt_label = ttk.Label(self.v_auto_frame, text="30", width=5)
        self.v_bt_label.grid(row=0, column=2)
        self.v_auto_frame.columnconfigure(1, weight=1)

        # 固定参数
        self.v_fixed_frame = ttk.Frame(f4)
        self.v_fixed_frame.grid(row=5, column=0, columnspan=4, sticky="ew", pady=2)
        ttk.Label(self.v_fixed_frame, text="上边距：").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Spinbox(self.v_fixed_frame, from_=0, to=500, textvariable=self.v_crop_margin_top, width=7).grid(row=0, column=1, padx=2)
        ttk.Label(self.v_fixed_frame, text="下边距：").grid(row=0, column=2, sticky="w", padx=(10, 0), pady=3)
        ttk.Spinbox(self.v_fixed_frame, from_=0, to=500, textvariable=self.v_crop_margin_bottom, width=7).grid(row=0, column=3, padx=2)
        ttk.Label(self.v_fixed_frame, text="左边距：").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Spinbox(self.v_fixed_frame, from_=0, to=500, textvariable=self.v_crop_margin_left, width=7).grid(row=1, column=1, padx=2)
        ttk.Label(self.v_fixed_frame, text="右边距：").grid(row=1, column=2, sticky="w", padx=(10, 0), pady=3)
        ttk.Spinbox(self.v_fixed_frame, from_=0, to=500, textvariable=self.v_crop_margin_right, width=7).grid(row=1, column=3, padx=2)
        self._v_update_crop()
        self._v_update_crop_mode()

        # ---- 输出参数 ----
        f5 = ttk.LabelFrame(scrollable, text="🖼 输出参数", padding=8)
        f5.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f5, text="输出格式：").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Combobox(f5, textvariable=self.v_format, values=["png", "jpg", "jpeg"], state="readonly", width=8).grid(row=0, column=1, sticky="w", pady=3)
        self.v_quality_label = ttk.Label(f5, text="图片质量：")
        self.v_quality_label.grid(row=1, column=0, sticky="w", pady=3)
        qf = ttk.Frame(f5)
        qf.grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Scale(qf, from_=1, to=100, variable=self.v_quality, orient=tk.HORIZONTAL,
                  command=lambda v: self.v_qval_label.config(text=f"{int(float(v))}")).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.v_qval_label = ttk.Label(qf, text="100", width=5)
        self.v_qval_label.pack(side=tk.RIGHT, padx=5)
        f5.columnconfigure(1, weight=1)

        # ---- 其他 ----
        f6 = ttk.LabelFrame(scrollable, text="⚙ 其他参数", padding=8)
        f6.pack(fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(f6, text="显示详细信息（调试用）", variable=self.v_verbose).pack(anchor=tk.W, pady=3)

        # ---- 控制按钮 ----
        cf = ttk.Frame(scrollable)
        cf.pack(fill=tk.X, pady=8)
        self.v_start_btn = ttk.Button(cf, text="▶ 开始处理", command=self._v_start)
        self.v_start_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        self.v_stop_btn = ttk.Button(cf, text="■ 停止", command=self._v_stop, state=tk.DISABLED)
        self.v_stop_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        ttk.Button(cf, text="🔄 重置参数", command=self._v_reset).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

    def _v_update_stable(self):
        en = self.v_stable_check.get()
        ns = "normal" if en else "disabled"
        try:
            self.v_stable_scale.config(state=ns)
        except Exception:
            pass
        self.v_stable_label.config(foreground="black" if en else "gray")
        self.v_stable_val_label.config(foreground="black" if en else "gray")

    def _v_update_crop(self):
        en = self.v_crop_enable.get()
        ns = "normal" if en else "disabled"
        for w in [self.v_crop_top_cb, self.v_crop_bottom_cb, self.v_crop_left_cb, self.v_crop_right_cb]:
            w.config(state=ns)
        self._v_update_crop_mode()

    def _v_update_crop_mode(self):
        en = self.v_crop_enable.get()
        is_auto = self.v_crop_mode.get() == "auto"
        ns_auto = "normal" if (en and is_auto) else "disabled"
        ns_fixed = "normal" if (en and not is_auto) else "disabled"
        for child in self.v_auto_frame.winfo_children():
            if isinstance(child, ttk.Scale):
                child.config(state=ns_auto)
            elif isinstance(child, ttk.Label):
                child.config(foreground="black" if ns_auto == "normal" else "gray")
        for child in self.v_fixed_frame.winfo_children():
            if isinstance(child, ttk.Spinbox):
                child.config(state=ns_fixed)
            elif isinstance(child, ttk.Label):
                child.config(foreground="black" if ns_fixed == "normal" else "gray")

    def _v_update_time(self):
        en = self.v_time_enable.get()
        ns = "normal" if en else "disabled"
        for w in [self.v_start_h_sb, self.v_start_m_sb, self.v_start_s_sb,
                  self.v_end_h_sb, self.v_end_m_sb, self.v_end_s_sb]:
            w.config(state=ns)
        self.v_time_start_label.config(foreground="black" if en else "gray")
        self.v_time_end_label.config(foreground="black" if en else "gray")

    def _v_select_files(self):
        files = filedialog.askopenfilenames(title="选择视频文件",
            filetypes=[("视频文件", "*.mp4 *.avi *.mov *.mkv *.wmv *.flv"), ("所有文件", "*.*")])
        for f in files:
            if f not in self.v_input_files:
                self.v_input_files.append(f)
                self.v_listbox.insert(tk.END, os.path.basename(f))
        if self.v_input_files:
            self.status_var.set(f"已选择 {len(self.v_input_files)} 个视频文件")

    def _v_clear_files(self):
        self.v_input_files.clear()
        self.v_listbox.delete(0, tk.END)
        self.status_var.set("已清空文件列表")

    def _v_reset(self):
        d = ConfigManager.DEFAULTS
        self.v_similarity.set(d["v_similarity"])
        self.v_min_interval.set(d["v_min_interval"])
        self.v_delay_frames.set(d["v_delay_frames"])
        self.v_stable_check.set(d["v_stable_check"])
        self.v_stable_threshold.set(d["v_stable_threshold"])
        self.v_format.set(d["v_format"])
        self.v_quality.set(d["v_quality"])
        self.v_skip.set(d["v_skip"])
        self.v_resize.set(d["v_resize"])
        self.v_verbose.set(d["v_verbose"])
        self.v_crop_enable.set(d["v_crop_enable"])
        self.v_crop_mode.set(d["v_crop_mode"])
        self.v_crop_top.set(d["v_crop_top"])
        self.v_crop_bottom.set(d["v_crop_bottom"])
        self.v_crop_left.set(d["v_crop_left"])
        self.v_crop_right.set(d["v_crop_right"])
        self.v_black_threshold.set(d["v_black_threshold"])
        self.v_crop_margin_top.set(d["v_crop_margin_top"])
        self.v_crop_margin_bottom.set(d["v_crop_margin_bottom"])
        self.v_crop_margin_left.set(d["v_crop_margin_left"])
        self.v_crop_margin_right.set(d["v_crop_margin_right"])
        self.v_time_enable.set(d["v_time_enable"])
        self.v_start_h.set(d["v_start_h"])
        self.v_start_m.set(d["v_start_m"])
        self.v_start_s.set(d["v_start_s"])
        self.v_end_h.set(d["v_end_h"])
        self.v_end_m.set(d["v_end_m"])
        self.v_end_s.set(d["v_end_s"])
        self.v_sim_label.config(text=f"{d['v_similarity']:.2f}")
        self.v_stable_val_label.config(text=f"{d['v_stable_threshold']:.2f}")
        self.v_qval_label.config(text=str(d["v_quality"]))
        self.v_bt_label.config(text=str(d["v_black_threshold"]))
        self._v_update_stable()
        self._v_update_crop()
        self._v_update_crop_mode()
        self._v_update_time()
        self.status_var.set("参数已重置")

    def _v_start(self):
        if self.running:
            return
        if not self.v_input_files:
            messagebox.showwarning("警告", "请先选择视频文件！")
            return
        output_name = self.v_output_dir.get().strip()
        if not output_name:
            messagebox.showwarning("警告", "请输入输出目录！")
            return
        self._stop_flag = False
        self._extractor = None
        self.running = True
        self.v_start_btn.config(state=tk.DISABLED)
        self.v_stop_btn.config(state=tk.NORMAL)
        self.v_select_btn.config(state=tk.DISABLED)
        self.v_clear_btn.config(state=tk.DISABLED)
        self.v_listbox.config(state=tk.DISABLED)
        self.global_progress["value"] = 0
        self.global_progress_label.config(text="0%")
        self.status_var.set("正在处理...")
        self.v_log.delete(1.0, tk.END)
        self._log_to_widget(self.v_log, "=" * 60)
        self._log_to_widget(self.v_log, "视频截图 - 开始处理")
        self._log_to_widget(self.v_log, "=" * 60)

        crop_sides = {"top": self.v_crop_top.get(), "bottom": self.v_crop_bottom.get(),
                       "left": self.v_crop_left.get(), "right": self.v_crop_right.get()} if self.v_crop_enable.get() else None
        crop_margins = {"top": self.v_crop_margin_top.get(), "bottom": self.v_crop_margin_bottom.get(),
                        "left": self.v_crop_margin_left.get(), "right": self.v_crop_margin_right.get()}
        if self.v_time_enable.get():
            start_seconds = self.v_start_h.get() * 3600 + self.v_start_m.get() * 60 + self.v_start_s.get()
            end_seconds = self.v_end_h.get() * 3600 + self.v_end_m.get() * 60 + self.v_end_s.get()
            if end_seconds > 0 and end_seconds <= start_seconds:
                messagebox.showwarning("警告", "结束时间必须大于开始时间！")
                self.running = False
                self.v_start_btn.config(state=tk.NORMAL)
                self.v_stop_btn.config(state=tk.DISABLED)
                self.v_select_btn.config(state=tk.NORMAL)
                self.v_clear_btn.config(state=tk.NORMAL)
                self.v_listbox.config(state=tk.NORMAL)
                return
        else:
            start_seconds = 0
            end_seconds = 0
        config = {
            "input": self.v_input_files,
            "output": output_name,
            "similarity": round(self.v_similarity.get(), 2),
            "min_interval": self.v_min_interval.get(),
            "delay_frames": self.v_delay_frames.get(),
            "stable_check": self.v_stable_check.get(),
            "stable_threshold": round(self.v_stable_threshold.get(), 2),
            "format": self.v_format.get(),
            "quality": self.v_quality.get(),
            "skip": self.v_skip.get(),
            "resize": self.v_resize.get(),
            "verbose": self.v_verbose.get(),
            "crop_sides": crop_sides,
            "crop_mode": self.v_crop_mode.get() if self.v_crop_enable.get() else "auto",
            "black_threshold": self.v_black_threshold.get(),
            "crop_margins": crop_margins,
            "start_seconds": start_seconds,
            "end_seconds": end_seconds
        }
        self.worker_thread = threading.Thread(target=self._v_run_extractor, args=(config,), daemon=True)
        self.worker_thread.start()

    def _v_run_extractor(self, config):
        try:
            extractor = VideoFrameExtractor(config,
                                            log_callback=lambda msg: self.log_queue.put(('v', msg)),
                                            progress_callback=self._v_progress_cb)
            self._extractor = extractor
            extractor.run()
        except Exception as e:
            self.log_queue.put(('v', f"错误：{str(e)}"))
        finally:
            self.root.after(0, self._v_finished)

    def _v_progress_cb(self, current, total, video_name, overall_pct=None):
        if total > 0:
            pct = overall_pct if overall_pct is not None else int(current / total * 100)
            self.progress_queue.put((pct, f"处理：{video_name} - {current:,}/{total:,} 帧 ({pct}%)"))

    def _v_stop(self):
        self._stop_flag = True
        extractor = getattr(self, '_extractor', None)
        if extractor is not None:
            extractor.stop()
        self.status_var.set("正在停止...")
        self.v_stop_btn.config(state=tk.DISABLED)
        self._log_to_widget(self.v_log, "用户请求停止处理...")

    def _v_finished(self):
        self.running = False
        self.v_start_btn.config(state=tk.NORMAL)
        self.v_stop_btn.config(state=tk.DISABLED)
        self.v_select_btn.config(state=tk.NORMAL)
        self.v_clear_btn.config(state=tk.NORMAL)
        self.v_listbox.config(state=tk.NORMAL)
        self.global_progress["value"] = 100
        self.global_progress_label.config(text="100%")
        self.status_var.set("视频截图处理完成！")
        self._log_to_widget(self.v_log, "=" * 60)
        self._log_to_widget(self.v_log, "视频截图处理已结束")
        self._log_to_widget(self.v_log, "=" * 60)

    # ==================================================================
    #  Tab 2: 转PPT
    # ==================================================================
    def _build_tab_pptx(self, notebook):
        tab, scrollable, self.p_log = self._create_tab_layout(notebook, "  转PPT  ")

        # ---- 变量 ----
        self.p_image_dir = tk.StringVar(value="ppt_output_GUI")
        self.p_ppt_path = tk.StringVar(value="output_presentation.pptx")
        self.p_page_w = tk.DoubleVar(value=16.0)
        self.p_page_h = tk.DoubleVar(value=9.0)
        self.p_bg_r = tk.IntVar(value=255)
        self.p_bg_g = tk.IntVar(value=255)
        self.p_bg_b = tk.IntVar(value=255)
        self.p_title = tk.StringVar(value="")
        self.p_title_size = tk.IntVar(value=24)
        self.p_title_r = tk.IntVar(value=0)
        self.p_title_g = tk.IntVar(value=0)
        self.p_title_b = tk.IntVar(value=0)
        self.p_sort_mode = tk.StringVar(value="name")
        self.p_sort_order = tk.StringVar(value="asc")
        self.p_fit_mode = tk.StringVar(value="stretch")

        # ---- 输入输出 ----
        f1 = ttk.LabelFrame(scrollable, text="📁 输入输出设置", padding=8)
        f1.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f1, text="图片文件夹：").grid(row=0, column=0, sticky="w", pady=3)
        df1 = ttk.Frame(f1)
        df1.grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Entry(df1, textvariable=self.p_image_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(df1, text="浏览...", command=lambda: self.p_image_dir.set(filedialog.askdirectory(title="选择图片文件夹") or self.p_image_dir.get())).pack(side=tk.RIGHT)
        ttk.Label(f1, text="PPT保存路径：").grid(row=1, column=0, sticky="w", pady=3)
        df2 = ttk.Frame(f1)
        df2.grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Entry(df2, textvariable=self.p_ppt_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(df2, text="浏览...", command=self._p_browse_ppt_save).pack(side=tk.RIGHT)
        f1.columnconfigure(1, weight=1)

        # ---- 页面设置 ----
        f2 = ttk.LabelFrame(scrollable, text="📐 页面设置", padding=8)
        f2.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f2, text="页面比例（宽×高）：").grid(row=0, column=0, sticky="w", pady=3)
        pf = ttk.Frame(f2)
        pf.grid(row=0, column=1, sticky="w", pady=3)
        ttk.Spinbox(pf, from_=1, to=100, textvariable=self.p_page_w, width=5).pack(side=tk.LEFT)
        ttk.Label(pf, text="×").pack(side=tk.LEFT, padx=3)
        ttk.Spinbox(pf, from_=1, to=100, textvariable=self.p_page_h, width=5).pack(side=tk.LEFT)
        ttk.Label(pf, text="英寸", foreground="gray").pack(side=tk.LEFT, padx=3)
        ttk.Button(pf, text="16:9", command=lambda: (self.p_page_w.set(16), self.p_page_h.set(9))).pack(side=tk.LEFT, padx=5)
        ttk.Button(pf, text="4:3", command=lambda: (self.p_page_w.set(10), self.p_page_h.set(7.5))).pack(side=tk.LEFT, padx=2)
        ttk.Label(f2, text="背景颜色 (RGB)：").grid(row=1, column=0, sticky="w", pady=3)
        bf = ttk.Frame(f2)
        bf.grid(row=1, column=1, sticky="w", pady=3)
        ttk.Spinbox(bf, from_=0, to=255, textvariable=self.p_bg_r, width=5).pack(side=tk.LEFT)
        ttk.Label(bf, text=",").pack(side=tk.LEFT, padx=2)
        ttk.Spinbox(bf, from_=0, to=255, textvariable=self.p_bg_g, width=5).pack(side=tk.LEFT)
        ttk.Label(bf, text=",").pack(side=tk.LEFT, padx=2)
        ttk.Spinbox(bf, from_=0, to=255, textvariable=self.p_bg_b, width=5).pack(side=tk.LEFT)
        ttk.Button(bf, text="白色", command=lambda: (self.p_bg_r.set(255), self.p_bg_g.set(255), self.p_bg_b.set(255))).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="黑色", command=lambda: (self.p_bg_r.set(0), self.p_bg_g.set(0), self.p_bg_b.set(0))).pack(side=tk.LEFT, padx=2)
        ttk.Label(f2, text="每页标题（留空为无）：").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(f2, textvariable=self.p_title, width=30).grid(row=2, column=1, sticky="w", pady=3)
        ttk.Label(f2, text="标题字号：").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Spinbox(f2, from_=8, to=72, textvariable=self.p_title_size, width=8).grid(row=3, column=1, sticky="w", pady=3)
        ttk.Label(f2, text="标题颜色 (RGB)：").grid(row=4, column=0, sticky="w", pady=3)
        tf = ttk.Frame(f2)
        tf.grid(row=4, column=1, sticky="w", pady=3)
        ttk.Spinbox(tf, from_=0, to=255, textvariable=self.p_title_r, width=5).pack(side=tk.LEFT)
        ttk.Label(tf, text=",").pack(side=tk.LEFT, padx=2)
        ttk.Spinbox(tf, from_=0, to=255, textvariable=self.p_title_g, width=5).pack(side=tk.LEFT)
        ttk.Label(tf, text=",").pack(side=tk.LEFT, padx=2)
        ttk.Spinbox(tf, from_=0, to=255, textvariable=self.p_title_b, width=5).pack(side=tk.LEFT)
        f2.columnconfigure(1, weight=1)

        # ---- 排序 ----
        f3 = ttk.LabelFrame(scrollable, text="🔢 图片排序", padding=8)
        f3.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f3, text="排序方式：").grid(row=0, column=0, sticky="w", pady=3)
        sf = ttk.Frame(f3)
        sf.grid(row=0, column=1, sticky="w", pady=3)
        ttk.Radiobutton(sf, text="按文件名", variable=self.p_sort_mode, value="name").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(sf, text="按修改时间", variable=self.p_sort_mode, value="date").pack(side=tk.LEFT, padx=2)
        ttk.Label(f3, text="排序方向：").grid(row=1, column=0, sticky="w", pady=3)
        of = ttk.Frame(f3)
        of.grid(row=1, column=1, sticky="w", pady=3)
        ttk.Radiobutton(of, text="正序 ↑", variable=self.p_sort_order, value="asc").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(of, text="倒序 ↓", variable=self.p_sort_order, value="desc").pack(side=tk.LEFT, padx=2)
        f3.columnconfigure(1, weight=1)

        # ---- 图片适应方式 ----
        f4 = ttk.LabelFrame(scrollable, text="🖼 图片适应方式", padding=8)
        f4.pack(fill=tk.X, pady=(0, 6))
        ttk.Radiobutton(f4, text="全屏拉伸（可能变形）", variable=self.p_fit_mode, value="stretch").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(f4, text="保持比例居中留白（不变形）", variable=self.p_fit_mode, value="fit").pack(anchor=tk.W, pady=2)

        # ---- 控制 ----
        cf = ttk.Frame(scrollable)
        cf.pack(fill=tk.X, pady=8)
        self.p_btn = ttk.Button(cf, text="▶ 开始生成PPT", command=self._p_start)
        self.p_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        self.p_stop_btn = ttk.Button(cf, text="■ 停止", command=self._p_stop, state=tk.DISABLED)
        self.p_stop_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        ttk.Button(cf, text="🔄 重置参数", command=self._p_reset).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

    def _p_browse_ppt_save(self):
        path = filedialog.asksaveasfilename(title="保存PPT为", defaultextension=".pptx",
            filetypes=[("PPTX文件", "*.pptx"), ("所有文件", "*.*")])
        if path:
            self.p_ppt_path.set(path)

    def _p_reset(self):
        d = ConfigManager.DEFAULTS
        self.p_page_w.set(d["p_page_w"])
        self.p_page_h.set(d["p_page_h"])
        self.p_bg_r.set(d["p_bg_r"])
        self.p_bg_g.set(d["p_bg_g"])
        self.p_bg_b.set(d["p_bg_b"])
        self.p_title.set(d["p_title"])
        self.p_title_size.set(d["p_title_size"])
        self.p_title_r.set(d["p_title_r"])
        self.p_title_g.set(d["p_title_g"])
        self.p_title_b.set(d["p_title_b"])
        self.p_sort_mode.set(d["p_sort_mode"])
        self.p_sort_order.set(d["p_sort_order"])
        self.p_fit_mode.set(d["p_fit_mode"])
        self.status_var.set("PPT参数已重置")

    def _p_start(self):
        if self.running:
            return
        img_dir = self.p_image_dir.get().strip()
        if not img_dir or not os.path.isdir(img_dir):
            messagebox.showwarning("警告", "图片文件夹路径无效或不存在！")
            return
        ppt_path = self.p_ppt_path.get().strip()
        if not ppt_path:
            messagebox.showwarning("警告", "请输入PPT保存路径！")
            return

        self._stop_flag = False
        self.running = True
        self.p_btn.config(state=tk.DISABLED)
        self.p_stop_btn.config(state=tk.NORMAL)
        self.global_progress["value"] = 0
        self.global_progress_label.config(text="0%")
        self.status_var.set("正在生成PPT...")
        self.p_log.delete(1.0, tk.END)
        self._log_to_widget(self.p_log, "=" * 60)
        self._log_to_widget(self.p_log, "转PPT - 开始处理")
        self._log_to_widget(self.p_log, "=" * 60)

        page_size = (self.p_page_w.get(), self.p_page_h.get())
        bg_color = (self.p_bg_r.get(), self.p_bg_g.get(), self.p_bg_b.get())
        title_color = (self.p_title_r.get(), self.p_title_g.get(), self.p_title_b.get())
        fit_mode = self.p_fit_mode.get()

        def run():
            try:
                count = images_to_pptx(
                    img_dir, ppt_path, page_size=page_size,
                    bg_color=bg_color, slide_title=self.p_title.get().strip(),
                    title_font_size=self.p_title_size.get(), title_color=title_color,
                    sort_mode=self.p_sort_mode.get(), sort_order=self.p_sort_order.get(),
                    fit_mode=fit_mode,
                    stop_check=lambda: self._stop_flag,
                    log_callback=lambda msg: self.log_queue.put(('p', msg)),
                    progress_callback=self._p_progress_cb
                )
                result_msg = f"共处理 {count} 张图片" if count else "未生成PPT"
                self.progress_queue.put((100, f"PPT生成完成！{result_msg}"))
            except Exception as e:
                self.log_queue.put(('p', f"❌ 错误：{str(e)}"))
                self.progress_queue.put((100, "PPT生成失败"))
            finally:
                self.root.after(0, self._p_finished)

        self.worker_thread = threading.Thread(target=run, daemon=True)
        self.worker_thread.start()

    def _p_stop(self):
        self._stop_flag = True
        self.p_stop_btn.config(state=tk.DISABLED)
        self.status_var.set("正在停止...")
        self._log_to_widget(self.p_log, "用户请求停止PPT生成...")

    def _p_progress_cb(self, pct, info):
        self.progress_queue.put((pct, f"PPT生成中... {info}"))

    def _p_finished(self):
        self.running = False
        self.p_btn.config(state=tk.NORMAL)
        self.p_stop_btn.config(state=tk.DISABLED)
        self.global_progress["value"] = 100
        self.global_progress_label.config(text="100%")
        self._log_to_widget(self.p_log, "=" * 60)
        self._log_to_widget(self.p_log, "PPT生成已结束")
        self._log_to_widget(self.p_log, "=" * 60)

    # ==================================================================
    #  Tab 3: 转PDF（图片→PDF）
    # ==================================================================
    def _build_tab_pdf(self, notebook):
        tab, scrollable, self.d_log = self._create_tab_layout(notebook, "  转PDF  ")

        # ---- 变量 ----
        self.d_image_dir = tk.StringVar(value="ppt_output_GUI")
        self.d_pdf_path = tk.StringVar(value="output_presentation.pdf")
        self.d_sort_mode = tk.StringVar(value="name")
        self.d_sort_order = tk.StringVar(value="asc")
        self.d_dpi = tk.IntVar(value=150)
        self.d_enable_compress = tk.BooleanVar(value=True)
        self.d_jpeg_quality = tk.IntVar(value=85)
        self.d_max_width = tk.IntVar(value=1920)

        # ---- 输入输出 ----
        f1 = ttk.LabelFrame(scrollable, text="📁 输入输出设置", padding=8)
        f1.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f1, text="图片文件夹：").grid(row=0, column=0, sticky="w", pady=3)
        df1 = ttk.Frame(f1)
        df1.grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Entry(df1, textvariable=self.d_image_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(df1, text="浏览...", command=lambda: self.d_image_dir.set(filedialog.askdirectory(title="选择图片文件夹") or self.d_image_dir.get())).pack(side=tk.RIGHT)
        ttk.Label(f1, text="PDF保存路径：").grid(row=1, column=0, sticky="w", pady=3)
        df2 = ttk.Frame(f1)
        df2.grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Entry(df2, textvariable=self.d_pdf_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(df2, text="浏览...", command=self._d_browse_pdf_save).pack(side=tk.RIGHT)
        f1.columnconfigure(1, weight=1)

        # ---- 说明 ----
        f2 = ttk.LabelFrame(scrollable, text="ℹ 说明", padding=8)
        f2.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f2, text="本功能将图片直接生成PDF文件，每张图片对应一页。", foreground="gray").pack(anchor=tk.W)
        ttk.Label(f2, text="无需安装 LibreOffice 或 Microsoft Office，基于 Pillow 实现。", foreground="gray").pack(anchor=tk.W)

        # ---- 输出参数 ----
        f2b = ttk.LabelFrame(scrollable, text="📐 输出参数", padding=8)
        f2b.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f2b, text="DPI 分辨率：").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Combobox(f2b, textvariable=self.d_dpi, values=[72, 96, 150, 200, 300], state="readonly", width=8).grid(row=0, column=1, sticky="w", pady=3)
        ttk.Label(f2b, text="（值越大越清晰，文件也越大）", foreground="gray").grid(row=0, column=2, sticky="w", pady=3, padx=5)

        # 图片压缩选项
        self.d_compress_cb = ttk.Checkbutton(f2b, text="启用图片压缩（缩放+JPEG，大幅减小PDF体积）", variable=self.d_enable_compress, command=self._d_update_compress)
        self.d_compress_cb.grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 3))
        self.d_jpeg_quality_label = ttk.Label(f2b, text="JPEG 质量：")
        self.d_jpeg_quality_label.grid(row=2, column=0, sticky="w", pady=3)
        qf = ttk.Frame(f2b)
        qf.grid(row=2, column=1, columnspan=2, sticky="ew", pady=3)
        self.d_quality_scale = ttk.Scale(qf, from_=10, to=100, variable=self.d_jpeg_quality, orient=tk.HORIZONTAL,
                                         command=lambda v: self.d_qval_label.config(text=f"{int(float(v))}"))
        self.d_quality_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.d_qval_label = ttk.Label(qf, text="85", width=5)
        self.d_qval_label.pack(side=tk.RIGHT, padx=5)
        self.d_max_width_label = ttk.Label(f2b, text="最大宽度(px)：")
        self.d_max_width_label.grid(row=3, column=0, sticky="w", pady=3)
        self.d_maxwidth_sb = ttk.Spinbox(f2b, from_=0, to=8000, textvariable=self.d_max_width, width=8)
        self.d_maxwidth_sb.grid(row=3, column=1, sticky="w", pady=3)
        ttk.Label(f2b, text="0=不缩放；超过此宽度的图片按比例缩小", foreground="gray").grid(row=3, column=2, sticky="w", pady=3, padx=5)
        pf = ttk.Frame(f2b)
        pf.grid(row=4, column=0, columnspan=3, sticky="w", pady=(5, 0))
        ttk.Label(pf, text="快捷预设：").pack(side=tk.LEFT)
        ttk.Button(pf, text="高压缩(60/1280)", width=16,
                   command=lambda: (self.d_jpeg_quality.set(60), self.d_max_width.set(1280), self.d_qval_label.config(text="60"))).pack(side=tk.LEFT, padx=3)
        ttk.Button(pf, text="均衡(85/1920)", width=14,
                   command=lambda: (self.d_jpeg_quality.set(85), self.d_max_width.set(1920), self.d_qval_label.config(text="85"))).pack(side=tk.LEFT, padx=3)
        ttk.Button(pf, text="高清(95/2560)", width=14,
                   command=lambda: (self.d_jpeg_quality.set(95), self.d_max_width.set(2560), self.d_qval_label.config(text="95"))).pack(side=tk.LEFT, padx=3)
        f2b.columnconfigure(1, weight=1)
        self._d_update_compress()

        # ---- 排序 ----
        f3 = ttk.LabelFrame(scrollable, text="🔢 图片排序", padding=8)
        f3.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f3, text="排序方式：").grid(row=0, column=0, sticky="w", pady=3)
        sf = ttk.Frame(f3)
        sf.grid(row=0, column=1, sticky="w", pady=3)
        ttk.Radiobutton(sf, text="按文件名", variable=self.d_sort_mode, value="name").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(sf, text="按修改时间", variable=self.d_sort_mode, value="date").pack(side=tk.LEFT, padx=2)
        ttk.Label(f3, text="排序方向：").grid(row=1, column=0, sticky="w", pady=3)
        of = ttk.Frame(f3)
        of.grid(row=1, column=1, sticky="w", pady=3)
        ttk.Radiobutton(of, text="正序 ↑", variable=self.d_sort_order, value="asc").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(of, text="倒序 ↓", variable=self.d_sort_order, value="desc").pack(side=tk.LEFT, padx=2)
        f3.columnconfigure(1, weight=1)

        # ---- 控制 ----
        cf = ttk.Frame(scrollable)
        cf.pack(fill=tk.X, pady=8)
        self.d_btn = ttk.Button(cf, text="▶ 开始生成PDF", command=self._d_start)
        self.d_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        ttk.Button(cf, text="🔄 重置参数", command=self._d_reset).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

    def _d_browse_pdf_save(self):
        path = filedialog.asksaveasfilename(title="保存PDF为", defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")])
        if path:
            self.d_pdf_path.set(path)

    def _d_update_compress(self):
        """根据压缩开关启用/禁用压缩相关控件"""
        en = self.d_enable_compress.get()
        ns = "normal" if en else "disabled"
        for w in [getattr(self, 'd_quality_scale', None),
                  getattr(self, 'd_maxwidth_sb', None)]:
            if w is not None:
                try:
                    w.config(state=ns)
                except Exception:
                    pass
        fg = "black" if en else "gray"
        for lbl in ['d_jpeg_quality_label', 'd_max_width_label', 'd_qval_label']:
            w = getattr(self, lbl, None)
            if w is not None:
                try:
                    w.config(foreground=fg)
                except Exception:
                    pass

    def _d_reset(self):
        d = ConfigManager.DEFAULTS
        self.d_image_dir.set(d["d_image_dir"])
        self.d_pdf_path.set(d["d_pdf_path"])
        self.d_sort_mode.set(d["d_sort_mode"])
        self.d_sort_order.set(d["d_sort_order"])
        self.d_dpi.set(d["d_dpi"])
        self.d_enable_compress.set(d["d_enable_compress"])
        self.d_jpeg_quality.set(d["d_jpeg_quality"])
        self.d_max_width.set(d["d_max_width"])
        self.d_qval_label.config(text=str(d["d_jpeg_quality"]))
        self._d_update_compress()
        self.status_var.set("PDF参数已重置")

    def _d_start(self):
        if self.running:
            return
        img_dir = self.d_image_dir.get().strip()
        if not img_dir or not os.path.isdir(img_dir):
            messagebox.showwarning("警告", "图片文件夹路径无效或不存在！")
            return
        pdf_path = self.d_pdf_path.get().strip()
        if not pdf_path:
            messagebox.showwarning("警告", "请输入PDF保存路径！")
            return

        self._stop_flag = False
        self.running = True
        self.d_btn.config(state=tk.DISABLED)
        self.global_progress["value"] = 0
        self.global_progress_label.config(text="0%")
        self.status_var.set("正在生成PDF...")
        self.d_log.delete(1.0, tk.END)
        self._log_to_widget(self.d_log, "=" * 60)
        self._log_to_widget(self.d_log, "转PDF - 开始处理（图片→PDF）")
        self._log_to_widget(self.d_log, "=" * 60)

        sort_mode = self.d_sort_mode.get()
        sort_order = self.d_sort_order.get()

        def run():
            try:
                count = images_to_pdf(
                    img_dir, pdf_path, sort_mode=sort_mode, sort_order=sort_order,
                    dpi=self.d_dpi.get(),
                    enable_compress=self.d_enable_compress.get(),
                    jpeg_quality=self.d_jpeg_quality.get(),
                    max_width=self.d_max_width.get(),
                    log_callback=lambda msg: self.log_queue.put(('d', msg)),
                    progress_callback=self._d_progress_cb
                )
                result_msg = f"共处理 {count} 张图片" if count else "未生成PDF"
                self.progress_queue.put((100, f"PDF生成完成！{result_msg}"))
            except Exception as e:
                self.log_queue.put(('d', f"❌ 错误：{str(e)}"))
                self.progress_queue.put((100, "PDF生成失败"))
            finally:
                self.root.after(0, self._d_finished)

        self.worker_thread = threading.Thread(target=run, daemon=True)
        self.worker_thread.start()

    def _d_progress_cb(self, pct, info):
        self.progress_queue.put((pct, f"PDF生成中... {info}"))

    def _d_finished(self):
        self.running = False
        self.d_btn.config(state=tk.NORMAL)
        self.global_progress["value"] = 100
        self.global_progress_label.config(text="100%")
        self._log_to_widget(self.d_log, "=" * 60)
        self._log_to_widget(self.d_log, "PDF生成已结束")
        self._log_to_widget(self.d_log, "=" * 60)

    # ==================================================================
    #  全局队列处理
    # ==================================================================
    def _process_queues(self):
        # 处理日志队列：按消息自带的目标Tab分发到对应日志控件（在主线程执行，线程安全）
        try:
            while True:
                target, msg = self.log_queue.get_nowait()
                if target == 'v' and hasattr(self, 'v_log'):
                    self._log_to_widget(self.v_log, msg)
                elif target == 'p' and hasattr(self, 'p_log'):
                    self._log_to_widget(self.p_log, msg)
                elif target == 'd' and hasattr(self, 'd_log'):
                    self._log_to_widget(self.d_log, msg)
        except queue.Empty:
            pass
        try:
            while True:
                pct, text = self.progress_queue.get_nowait()
                self.global_progress["value"] = pct
                self.global_progress_label.config(text=f"{pct}%")
                self.status_var.set(text)
        except queue.Empty:
            pass
        if self.running:
            self._drain_count = 0
            self.after_id = self.root.after(100, self._process_queues)
        elif self._drain_count < self._max_drain:
            self._drain_count += 1
            self.after_id = self.root.after(100, self._process_queues)
        else:
            self.after_id = self.root.after(200, self._process_queues)

    def _on_closing(self):
        if self.running:
            if not messagebox.askokcancel("确认退出", "正在处理中，确定要退出吗？"):
                return
            self._stop_flag = True
            extractor = getattr(self, '_extractor', None)
            if extractor is not None:
                extractor.stop()
        # 等待工作线程退出（最多5秒）
        if self.worker_thread and self.worker_thread.is_alive():
            self.worker_thread.join(timeout=5)
        # 退出前自动保存设置
        self._save_all_config()
        if self.after_id:
            self.root.after_cancel(self.after_id)
        self.root.destroy()


# =====================================================================
#  启动入口
# =====================================================================
def main():
    root = tk.Tk()
    app = IntegratedApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()