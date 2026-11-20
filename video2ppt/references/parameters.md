# video2ppt 参数参考

所有脚本的通用约定：

- 日志写 stderr，加 `--json` 时在 stdout 输出单行 JSON 摘要。
- 路径支持绝对/相对路径；`--input` 支持文件、目录、通配符（跑 `*?[`）。
- 退出码：`0` 成功，`1` 无产出，`2` 参数错误，`3` 缺少依赖，`130` 中断。

---

## scripts/video_to_frames.py

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `-i, --input` | 必填 | 视频文件、目录或通配符，可多个 |
| `-o, --output` | 视频所在目录/`frames` | 帧输出目录；省略时放在第一个视频旁边 |
| `--similarity` | `0.8` | 相似度阈值 0–1。相似度**低于**该值判定为画面切换。值越低越不敏感（页更少），越高越敏感（页更多） |
| `--min-interval` | `5` | 两次保存之间的最小帧号差 |
| `--delay-frames` | `5` | 检测到切换后延迟的帧数，用于跳过转场动画 |
| `--no-stable-check` | 关闭 | 关闭稳定性检查；缓冲满后直接保存最后一帧 |
| `--stable-threshold` | `0.95` | 稳定性阈值 0–1，画面连续相似度均高于它才认为稳定 |
| `--skip` | `15` | 跳帧采样：每 `skip+1` 帧处理 1 帧；`0` = 逐帧（最慢最准） |
| `--resize` | `480` | 相似度计算前把灰度图缩放到该宽度；`0` = 不缩放 |
| `--start` | `0` | 开始时间：秒、`MM:SS` 或 `HH:MM:SS` |
| `--end` | `0` | 结束时间：同上；`0` = 到结尾 |
| `--crop` | `none` | 裁剪方向：`none`、`all` 或 `top,bottom,left,right` 组合 |
| `--crop-mode` | `auto` | `auto` 按亮度自动检测黑边；`fixed` 用固定边距 |
| `--black-threshold` | `30` | 自动裁剪时判定“黑”的行/列灰度均值上限 |
| `--margin-top/bottom/left/right` | 10/10/8/8 | `fixed` 模式裁剪像素 |
| `--frame-format` | `png` | `png`、`jpg`、`jpeg` |
| `--quality` | `100` | JPEG 质量 1–100（仅 jpg） |
| `--verbose` | 关闭 | 输出逐帧相似度、缓冲、清晰度评分 |
| `--quiet` | 关闭 | 仅输出错误/警告 |
| `--json` | 关闭 | stdout 输出 JSON 摘要 |

JSON 字段：`status`（`ok`/`empty`）、`output_dir`、`videos`、`frames_saved`、`frame_format`、`time_range`、`errors`。

### 算法流程

1. 打开视频，按 `--start/--end` 定位帧范围。
2. 每 `--skip+1` 帧读一帧，缩放到 `--resize` 宽度后转灰度。
3. 与上一张参考帧计算 SSIM 相似度（无 scikit-image 时回退为变化像素占比）。
4. 相似度 < `--similarity` → 画面切换，当前帧进入延迟缓冲。
5. 后续画面稳定（最近 3 帧两两相似度 ≥ `--stable-threshold`）或缓冲达到 `delay+5` 帧时，用拉普拉斯方差挑最清晰的一帧保存。
6. 保存时按需裁剪黑边，输出文件名 `<视频名>_ppt_<页码>_frame_<帧号>.<ext>`。

---

## scripts/images_to_pptx.py

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `-i, --input` | 必填 | 图片文件夹（支持 jpg/jpeg/png/gif/webp/bmp/tiff） |
| `-o, --output` | `output_presentation.pptx` | 输出路径 |
| `--slide-width` / `--slide-height` | `16` / `9` | 页面尺寸（英寸） |
| `--bg` | `255,255,255` | 背景色 `R,G,B` |
| `--title` | 空 | 每页标题前缀，生成“标题 - 第N页”；空则无标题栏 |
| `--title-size` | `24` | 标题字号 |
| `--title-color` | `0,0,0` | 标题颜色 `R,G,B` |
| `--fit` | `stretch` | `stretch` 拉伸铺满页面（可能变形）；`fit` 等比缩放居中留白 |
| `--sort` | `name` | `name` 自然排序（page_2 < page_10）；`date` 按修改时间 |
| `--order` | `asc` | `asc` / `desc` |
| `--quiet` / `--json` | 关闭 | 同上 |

JSON 字段：`status`、`output`、`images_found`、`slides_added`、`page_size_in`、`fit_mode`。

---

## scripts/images_to_pdf.py

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `-i, --input` | 必填 | 图片文件夹 |
| `-o, --output` | `output_presentation.pdf` | 输出路径 |
| `--sort` / `--order` | `name` / `asc` | 同 PPTX |
| `--dpi` | `150` | PDF 页面物理分辨率声明，不影响压缩后像素 |
| `--no-compress` | 关闭 | 保留原画质嵌入（等同旧版行为） |
| `--jpeg-quality` | `85` | 压缩时 JPEG 质量 10–100 |
| `--max-width` | `1920` | 压缩时最大像素宽度，`0` = 不缩放 |
| `--quiet` / `--json` | 关闭 | 同上 |

JSON 字段：`status`、`output`、`images_found`、`pages`、`size_bytes`、`compress`、`engine`。

压缩引擎优先级：`img2pdf`（JPEG 直接嵌入）→ Pillow 逐页写入。
预设参考：高压缩 `--jpeg-quality 60 --max-width 1280`；均衡 `85/1920`；高清 `95/2560`。

---

## scripts/video2ppt.py（流水线）

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `-i, --input` | 必填 | 视频文件、目录或通配符 |
| `-o, --outdir` | 第一个视频所在目录 | 输出目录；省略时与视频放在一起 |
| `--name` | 首个视频名 | 输出文件基名 |
| `--formats` | `pptx,pdf` | 逗号分隔；留空则只抽帧 |
| 抽帧参数 | — | `--similarity`、`--min-interval`、`--delay-frames`、`--no-stable-check`、`--stable-threshold`、`--skip`、`--resize`、`--start`、`--end`、`--crop`、`--crop-mode`、`--black-threshold`、`--frame-format`、`--quality`（含义同 `video_to_frames.py`） |
| PPTX 参数 | — | `--slide-width`、`--slide-height`、`--bg`、`--title`、`--fit` |
| PDF 参数 | — | `--dpi`、`--no-compress`、`--jpeg-quality`、`--max-width` |
| 排序 | — | `--sort`、`--order` |
| 其他 | — | `--verbose`、`--quiet`、`--json` |

JSON 字段：`status`（`ok`/`empty`/`partial`）、`outdir`、`frames_dir`、`frames`、`outputs.pptx.{path,slides}`、`outputs.pdf.{path,pages}`、`errors`。

---

## 依赖矩阵

| 包 | 必需 | 用途 | 缺失时 |
| --- | --- | --- | --- |
| `opencv-python` | 是 | 视频解码、抽帧、图像编码 | 无法抽帧（退出码 3） |
| `numpy` | 是 | 数组运算 | 无法抽帧（退出码 3） |
| `python-pptx` | 合成 PPTX 时 | 生成演示文稿 | 只能抽帧/转 PDF |
| `Pillow` | 合成 PDF 时 | PDF 写入、图片压缩、fit 计算 | 只能抽帧/转 PPT |
| `scikit-image` | 否 | SSIM 去重 | 回退为变化像素占比算法 |
| `img2pdf` | 否 | PDF 流式合并 | 回退 Pillow 逐页写入 |
