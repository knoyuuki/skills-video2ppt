---
name: video2ppt
description: 从视频/录屏中智能提取 PPT 画面帧，并生成 PPTX 演示文稿或 PDF（每图一页）。当用户给出视频、录屏、课程录像、会议回放并希望得到幻灯片/PPT/PDF 时使用；也适用于把一个图片文件夹批量合成为 PPT 或 PDF。支持时间范围截取、SSIM 画面去重、延迟截图+清晰度选帧、黑边裁剪、PDF 压缩。
---

# video2ppt

把视频中的幻灯片画面提取成图片，再合成为 PPTX（每页一张图）或 PDF。
三个能力可单独使用，也可一键串联：

| 脚本 | 作用 | 典型场景 |
| --- | --- | --- |
| `scripts/video2ppt.py` | 一键流水线：视频 → 帧 → PPTX/PDF | 用户只要最终 PPT/PDF（**优先使用**） |
| `scripts/video_to_frames.py` | 从视频智能抽帧，输出图片序列 | 用户需要截图、或想人工筛选后再合成 |
| `scripts/images_to_pptx.py` | 图片文件夹 → PPTX | 已有截图/图片，要生成演示文稿 |
| `scripts/images_to_pdf.py` | 图片文件夹 → PDF | 已有截图/图片，要生成 PDF |

所有命令均在**本 Skill 根目录**（包含本文件的目录）下执行；脚本也会写日志到 stderr，可配合 `--json` 读取 stdout 的机器可读摘要。

## 0. 环境准备（首次使用）

```bash
python scripts/check_env.py
```

- 退出码 0 表示依赖齐全；退出码 3 表示缺少必装依赖。
- 缺依赖时安装：`python -m pip install -r scripts/requirements.txt`
- 必需：`opencv-python`、`numpy`、`python-pptx`、`Pillow`；可选增强：`scikit-image`（SSIM 去重）、`img2pdf`（PDF 流式合并，缺失时自动回退 Pillow）。
- 脚本不读取任何 API Key 或环境变量，无硬编码凭据。

## 1. 一键流水线（推荐入口）

```bash
python scripts/video2ppt.py -i <视频文件/目录/通配符> -o <输出目录> --formats pptx,pdf --json
```

示例：

```bash
# 整个视频 → PPT + PDF
python scripts/video2ppt.py -i lecture.mp4 -o ./out --json

# 只处理 00:05:00 ~ 00:40:00，输出保持比例不变形
python scripts/video2ppt.py -i lecture.mp4 -o ./out --start 00:05:00 --end 00:40:00 --fit fit

# 批量处理目录内所有视频，只生成 PPTX
python scripts/video2ppt.py -i ./videos -o ./out --formats pptx
```

产物：`<outdir>/frames/`（提取的画面）、`<outdir>/<name>.pptx`、`<outdir>/<name>.pdf`。
`--name` 可指定输出文件基名，默认取第一个视频的文件名。

## 2. 分步执行

### 2.1 视频抽帧

```bash
python scripts/video_to_frames.py -i <视频/目录> -o <帧目录> [参数] --json
```

- 关键参数：`--start` / `--end`（秒或 `HH:MM:SS`）、`--similarity`（默认 0.8）、`--skip`（默认 15）、`--crop all`（去除黑边）。
- 输出文件名形如 `<视频名>_ppt_001_frame_000123.png`，按页码自然排序。

### 2.2 图片 → PPTX

```bash
python scripts/images_to_pptx.py -i <图片目录> -o <输出.pptx> --fit fit --json
```

- `--fit stretch`（默认，全屏拉伸）或 `--fit fit`（保持比例，居中留白）。
- 可选 `--title 课程名`、`--slide-width/--slide-height`（英寸，默认 16×9）、`--bg 255,255,255`。

### 2.3 图片 → PDF

```bash
python scripts/images_to_pdf.py -i <图片目录> -o <输出.pdf> --json
```

- 默认开启压缩（`--max-width 1920`、`--jpeg-quality 85`）；用 `--no-compress` 保留原画质。
- `--max-width 0` 表示不缩放。

## 3. 参数调节指南

| 问题 | 调整 |
| --- | --- |
| 提取的页太少 | 提高 `--similarity`（0.8 → 0.85~0.9），降低 `--skip` 和 `--min-interval` |
| 重复页太多 | 降低 `--similarity`（0.8 → 0.7~0.75），提高 `--skip` / `--min-interval` |
| 截到转场动画/模糊帧 | 提高 `--delay-frames`（如 8~12），保持稳定性检查开启 |
| 画面一直判为不稳定 | 降低 `--stable-threshold`（0.95 → 0.9）或加 `--no-stable-check` |
| 输出带黑边 | `--crop all`；噪点导致检测异常时用 `--crop-mode fixed` 配合边距参数 |
| 视频有 4:3 内容但输出变形 | 合成时改用 `--fit fit` |
| PDF 太大 | 保持压缩开启，降低 `--jpeg-quality`（85→60）或 `--max-width`（1920→1280） |
| 处理太慢 | 提高 `--skip`（如 25）、降低 `--resize`（480→320）、关闭 `--verbose` |

完整参数表见 `references/parameters.md`（仅在需要细调时阅读）。

## 4. 结果校验与异常处理

1. 先看退出码：`0` 成功；`1` 未产出任何结果（帧数为 0 / 输入为空）；`2` 参数错误；`3` 缺少依赖；`130` 被中断。
2. 使用 `--json` 时解析 stdout（日志在 stderr），关键字段：
   - 抽帧：`frames_saved`、`output_dir`、`errors`
   - 流水线：`frames`、`outputs.pptx.slides`、`outputs.pdf.pages`、`errors`
3. 成功后确认输出文件存在且大小 > 0，并向用户报告页数与绝对路径。
4. 常见失败：
   - **无法打开视频**：路径含特殊字符或编码不支持。先用 ffmpeg 转码为 H.264 MP4：`ffmpeg -i in.mov -c:v libx264 -c:a aac out.mp4`。
   - **frames_saved = 0**：按第 3 节提高相似度敏感度后重试。
   - **导入 cv2/numpy 时出现 OpenBLAS / 内存分配错误**：本机 numpy 安装损坏。重建虚拟环境（`python -m venv .venv` 后重装 `requirements.txt`），或 `python -m pip install --force-reinstall numpy opencv-python`。`check_env.py` 与抽帧脚本已用子进程隔离探测，会以退出码 3 报出该问题而不是直接崩溃。
   - **中文路径/日志乱码**：脚本已强制 UTF-8 输出；若终端仍乱码，先执行 `export PYTHONIOENCODING=utf-8`（Windows: `set PYTHONIOENCODING=utf-8`）。
5. 不要用 GUI 版参数文件交互；本 Skill 的脚本全部无界面、可重复执行。

## 5. 环境变量

无必需的环境变量。脚本不包含硬编码的密钥、密码或令牌，也不访问网络（仅首次安装依赖时 pip 需要联网）。
