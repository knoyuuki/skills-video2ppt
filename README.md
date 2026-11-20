# skills-video2ppt

一个 Agent Skill：**视频/录屏 → 幻灯片画面 → PPTX / PDF**。
将视频或图片文件夹一键转换为 PPTX 演示文稿（每图一页）或 PDF，无需 Office / LibreOffice。

来源：由 GUI 工具 `integrated_tool_v4.py` 拆分重构而来，保留原抽帧算法（SSIM 去重、延迟截图 + 清晰度选帧、黑边裁剪、PDF 压缩），全部改为无界面 CLI，适合 Claude Code 等 Agent 直接调用。

## 安装

### 使用 skills CLI（推荐）

```bash
npx skills add knoyuuki/skills-video2ppt
```

### 手动安装（Claude Code）

```bash
git clone git@github.com:knoyuuki/skills-video2ppt.git /tmp/skills-video2ppt
mkdir -p ~/.claude/skills
cp -r /tmp/skills-video2ppt/video2ppt ~/.claude/skills/video2ppt
```

安装后目录应为 `~/.claude/skills/video2ppt/SKILL.md`。

### 安装 Python 依赖

```bash
cd ~/.claude/skills/video2ppt
python -m pip install -r scripts/requirements.txt
python scripts/check_env.py
```

## 快速使用

```bash
python scripts/video2ppt.py -i lecture.mp4 -o ./out --formats pptx,pdf --json
```

产物：`out/frames/`、`out/lecture.pptx`、`out/lecture.pdf`。
省略 `-o` 时输出到视频所在目录（如 `video/lecture.mp4` → `video/frames/`、`video/lecture.pptx`、`video/lecture.pdf`）。

单独能力：

```bash
python scripts/video_to_frames.py -i lecture.mp4 -o frames/ --start 00:05:00 --crop all
python scripts/images_to_pptx.py -i frames/ -o slides.pptx --fit fit
python scripts/images_to_pdf.py  -i frames/ -o slides.pdf --jpeg-quality 70
```

详细说明见 [`video2ppt/SKILL.md`](video2ppt/SKILL.md)，完整参数见 [`video2ppt/references/parameters.md`](video2ppt/references/parameters.md)。

## 目录结构

```
skills-video2ppt/
├── README.md
├── video2ppt/               # Skill 本体
│   ├── SKILL.md
│   ├── scripts/
│   │   ├── video2ppt.py        # 一键流水线
│   │   ├── video_to_frames.py  # 视频抽帧
│   │   ├── images_to_pptx.py   # 图片 → PPTX
│   │   ├── images_to_pdf.py    # 图片 → PDF
│   │   ├── check_env.py        # 依赖自检
│   │   ├── _common.py
│   │   └── requirements.txt
│   └── references/
│       └── parameters.md
└── legacy/
    └── integrated_tool_v4.py   # 原始 GUI 版本（人工操作用，非 Skill 入口）
```

## 安全说明

仓库中不包含任何 API Key、密码或令牌，脚本也不读取环境变量或访问网络（仅安装依赖时 pip 需要联网）。
