#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""video2ppt 技能脚本的共享工具函数。"""

import os
import re
import sys

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tiff"}


def setup_utf8_stdio():
    """Windows 终端下保证中文日志正常输出。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def natural_key(path):
    """自然排序键：page_2 < page_10，而非字典序 page_10 < page_2。"""
    return [int(t) if t.isdigit() else t.lower()
            for t in re.split(r"(\d+)", os.path.basename(path))]


def human_size(num):
    """将字节数转为人类可读大小字符串。"""
    num = float(num)
    for unit in ["B", "KB", "MB", "GB"]:
        if num < 1024.0:
            return f"{num:.1f}{unit}"
        num /= 1024.0
    return f"{num:.1f}TB"


def collect_images(image_dir, sort_mode="name", sort_order="asc"):
    """收集目录中的图片并按名称/修改时间排序，返回绝对路径列表。"""
    image_paths = []
    for f in os.listdir(image_dir):
        if os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS:
            image_paths.append(os.path.join(image_dir, f))
    if sort_mode == "name":
        image_paths.sort(key=natural_key)
    elif sort_mode == "date":
        image_paths.sort(key=lambda x: os.path.getmtime(x))
    if str(sort_order).lower() == "desc":
        image_paths = image_paths[::-1]
    return image_paths


def parse_time(value):
    """把 '90'、'MM:SS'、'HH:MM:SS' 解析为秒数。"""
    s = str(value).strip()
    if not s:
        return 0.0
    if ":" not in s:
        seconds = float(s)
        if seconds < 0:
            raise ValueError(f"时间不能为负数：{value}")
        return seconds
    parts = s.split(":")
    if len(parts) not in (2, 3):
        raise ValueError(f"无法解析时间：{value}（支持 秒 / MM:SS / HH:MM:SS）")
    try:
        nums = [float(p) for p in parts]
    except ValueError:
        raise ValueError(f"无法解析时间：{value}（支持 秒 / MM:SS / HH:MM:SS）")
    if any(n < 0 for n in nums):
        raise ValueError(f"时间不能为负数：{value}")
    seconds = 0.0
    for n in nums:
        seconds = seconds * 60 + n
    return seconds


def parse_rgb(value):
    """把 'R,G,B' 解析为三元组，取值范围 0-255。"""
    try:
        parts = [int(p.strip()) for p in str(value).split(",")]
    except ValueError:
        raise ValueError(f"颜色格式应为 R,G,B：{value}")
    if len(parts) != 3 or any(p < 0 or p > 255 for p in parts):
        raise ValueError(f"颜色需为 3 个 0-255 的整数（R,G,B）：{value}")
    return tuple(parts)


def parse_crop(value):
    """解析裁剪方向：none / all / 逗号分隔的 top,bottom,left,right。"""
    v = str(value or "").strip().lower()
    sides = {"top": False, "bottom": False, "left": False, "right": False}
    if v in ("", "none", "no", "false", "0", "off"):
        return None
    if v in ("all", "yes", "true", "1", "on"):
        return {k: True for k in sides}
    for token in re.split(r"[,\s]+", v):
        if token not in sides:
            raise ValueError(f"未知裁剪方向：{token}（可选 top,bottom,left,right,all,none）")
        sides[token] = True
    return sides


def make_logger(quiet=False):
    """返回写 stderr 的日志函数；quiet 时仅保留错误/警告。"""
    def _log(msg):
        if quiet and not any(k in str(msg) for k in ("错误", "失败", "警告", "无法", "❌", "⚠")):
            return
        print(msg, file=sys.stderr, flush=True)
    return _log


class ThrottledProgress:
    """按百分比增量节流输出进度，避免逐帧刷屏。"""

    def __init__(self, enabled=True, min_delta=5.0, prefix="进度"):
        self.enabled = enabled
        self.min_delta = min_delta
        self.prefix = prefix
        self.last_pct = None

    def __call__(self, pct, text=""):
        if not self.enabled:
            return
        pct = int(pct)
        if self.last_pct is None or pct - self.last_pct >= self.min_delta or pct >= 100:
            self.last_pct = pct
            suffix = f" {text}" if text else ""
            print(f"[{self.prefix}] {pct}%{suffix}", file=sys.stderr, flush=True)
