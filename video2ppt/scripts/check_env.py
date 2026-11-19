#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检查 video2ppt 运行依赖，输出 JSON 结果。

每个依赖在独立子进程中探测，避免某个损坏的库（例如 OpenBLAS 初始化失败）
把检查进程本身一起拖垮。

退出码：0 = 必需依赖齐全；3 = 缺少必需依赖。
"""

import json
import subprocess
import sys
from importlib import metadata

from _common import setup_utf8_stdio

PACKAGES = [
    ("cv2", "opencv-python", True, "视频解码、抽帧、图像编码"),
    ("numpy", "numpy", True, "数组运算"),
    ("skimage", "scikit-image", False, "SSIM 相似度（缺失时回退差值算法）"),
    ("pptx", "python-pptx", True, "生成 PPTX"),
    ("PIL", "Pillow", True, "生成 PDF、图片压缩"),
    ("img2pdf", "img2pdf", False, "PDF 流式合并（缺失时回退 Pillow）"),
]


def probe(module):
    """在子进程中尝试 import 模块，返回 (是否成功, 错误信息, 版本)。"""
    try:
        version = metadata.version(module if module != "PIL" else "Pillow")
    except Exception:
        version = None
    try:
        proc = subprocess.run(
            [sys.executable, "-c", f"import {module}"],
            capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=120,
        )
    except subprocess.TimeoutExpired:
        return False, "import 超时（120s）", version
    if proc.returncode == 0:
        return True, None, version
    err = (proc.stderr or proc.stdout or "").strip()
    return False, err[-300:] or f"退出码 {proc.returncode}", version


def main():
    setup_utf8_stdio()
    packages = {}
    missing_required = []
    for module, pkg, required, usage in PACKAGES:
        ok, err, version = probe(module)
        entry = {"required": required, "usage": usage, "installed": ok}
        if version:
            entry["version"] = version
        if not ok:
            entry["error"] = err
            if required:
                missing_required.append(pkg)
        packages[pkg] = entry

    result = {
        "ready": not missing_required,
        "python": sys.version.split()[0],
        "packages": packages,
        "missing_required": missing_required,
        "install_hint": ("python -m pip install -r requirements.txt"
                         if missing_required else None),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not missing_required else 3


if __name__ == "__main__":
    sys.exit(main())
