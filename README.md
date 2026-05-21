# **ScenePilot-4K: A Large-Scale First-Person Dataset and Benchmark for Vision-Language Models in Autonomous Driving**

# 📄 Supplementary Material
The supplementary material is now publicly available and includes additional details on the dataset construction pipeline, benchmark design, evaluation metrics, etc.

👉 **Access the supplementary material here:**  
[Supplementary Material (PDF)](https://drive.google.com/file/d/14CisDMqTLfFrd8PADajuL0TNR5V90BSr/view?usp=drive_link)

<div align="center">
  <img src="assets/fig1.png" width="800px">
  <p>Figure 1: Overview of the ScenePilot-Bench benchmark and evaluation metrics.</p>
</div>


[![Project Page](https://img.shields.io/badge/Project-Website-blue?style=flat-square)](https://github.com/yjwangtj/ScenePilot-Bench)
[![Dataset](https://img.shields.io/badge/Dataset-Download-green?style=flat-square)](https://huggingface.co/datasets/larswangtj/ScenePilot-4K/tree/main) 
[![Paper](https://img.shields.io/badge/Paper-Arxiv-red?style=flat-square)](https://arxiv.org/abs/2601.19582)
[![Supplementary Material](https://img.shields.io/badge/Supplementary-Material-orange?style=flat-square)](https://drive.google.com/file/d/14CisDMqTLfFrd8PADajuL0TNR5V90BSr/view?usp=drive_link)

# 📖 Introduction
We introduce ScenePilot-4K, a large-scale first-person driving dataset for safety-aware vision-language learning and evaluation in autonomous driving. Built from public online driving videos, ScenePilot-4K contains 3,847 hours of video and 27.7M front-view frames spanning 63 countries/regions and 1,210 cities. It jointly provides scene-level natural-language descriptions, risk assessment labels, key-participant annotations, ego trajectories, and camera parameters through a unified multi-stage annotation pipeline. Building on this dataset, we establish ScenePilot-Bench, a standardized benchmark that evaluates vision-language models along four complementary axes: scene understanding, spatial perception, motion planning, and GPT-based semantic alignment. The benchmark includes fine-grained metrics and geographic generalization settings that expose model robustness under cross-region and cross-traffic domain shifts. Baseline results on representative open-source and proprietary vision-language models show that current models remain competitive in high-level scene semantics but still exhibit substantial limitations in geometry-aware perception and planning-oriented reasoning.

# 🛠️ Installation

```bash
# 1. Clone the repository
git clone https://github.com/yjwangtj/ScenePilot-Bench.git
cd ScenePilot-Bench

# 2. Create and activate a Conda environment
conda create -n scenepilot python=3.10 -y
conda activate scenepilot

# 3. Install required dependencies

pip install -r requirements.txt
```

## Batch video downloader from Excel

The repository now includes `video_batch_downloader.py`, a standalone script for reading an Excel/CSV file and downloading video links from platforms such as YouTube and Bilibili with `yt-dlp`.

### Supported manifest columns

The script auto-detects these logical columns and also supports the Chinese aliases below:

- first-level folder: `folder_name` / `文件夹名称`
- slice folder: `slice_name` / `切片文件夹`
- video url: `video_url` / `视频链接`
- sequence or name (optional): `序号`

Example table:

| 文件夹名称 | 切片文件夹 | 序号 | 视频链接 |
| --- | --- | --- | --- |
| 城市场景 | 路口 | 01 | https://www.youtube.com/watch?v=... |
| 城市场景 | 路口 | 02 | https://www.bilibili.com/video/BV... |
| 高速场景 | 夜间 | 01 | https://www.youtube.com/watch?v=... |

The output structure is:

```text
downloads/
  城市场景/
    路口/
      01.mp4
      02.mp4
  高速场景/
    夜间/
      01.mp4
```

### Dry run first

Use `--plan-only` first to verify the parsed rows and output paths without downloading:

```bash
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --plan-only
```

### Start the download

```bash
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads
```

If your sheet uses custom headers, pass them explicitly:

```bash
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --folder-column "一级目录" \
  --slice-column "切片目录" \
  --name-column "编号" \
  --url-column "链接"
```

### Stability and anti-bot precautions

The script enables a conservative download profile by default:

- single-fragment concurrency to reduce burst requests
- randomized sleep between video tasks and individual requests
- extractor retries plus task-level retries with exponential backoff
- resume support via `--continue`
- Bilibili-specific `Referer` / `Origin` headers
- automatic cleanup of Bilibili share tracking query parameters
- optional rate limiting via `--limit-rate`
- optional proxy via `--proxy`
- optional YouTube-only proxy routing via `--youtube-proxy` or `--youtube-proxy-port`
- optional browser impersonation via `--impersonate`
- CSV report generation for retrying failed rows later

For YouTube/Bilibili, authenticated cookies usually improve stability for rate-limited or age-gated content:

```bash
# Use exported cookies.txt
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --cookies-file /path/to/cookies.txt
```

```bash
# Or read cookies from a local browser profile
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --cookies-browser chrome
```

On Windows, `--cookies-browser chrome` can fail with `Could not copy Chrome cookie database` if Chrome/Edge is still running and the cookies DB is locked. The quickest fixes are:

1. fully close the browser first, including background processes in Task Manager
2. rerun the downloader
3. if you do not want to close the browser, export `cookies.txt` manually and use `--cookies-file`

For Bilibili specifically, `HTTP Error 412` usually means anti-bot blocking. In practice, the most effective order is:

1. open the exact Bilibili video in a normal browser first
2. rerun with `--cookies-browser chrome` (or `edge` / `firefox`)
3. if your current IP is a VPN / server / data-center exit, switch to a residential or home network
4. optionally install `curl-cffi` and try `--impersonate chrome`

The downloader treats Bilibili `HTTP 412` as a known anti-bot block and now stops retrying that row immediately, so one blocked link does not waste multiple retry cycles before moving on to the next row.

Example:

```bash
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --cookies-browser chrome \
  --impersonate chrome
```

Why `yt-dlp` here instead of separate libraries?

- `yt-dlp` is already a dedicated downloader for YouTube, Bilibili, and many other sites
- YouTube-only libraries such as `pytube` break more often when YouTube changes
- Bilibili-specific Python libraries are useful for metadata or account operations, but for bulk downloading across both platforms, `yt-dlp` is the most practical default
- this script simply wraps `yt-dlp` with Excel parsing, folder naming, retries, cookies, and anti-bot precautions

To make the crawl less bursty, the script now also sleeps a random amount before **every video row** by default:

- `--task-sleep-min 5`
- `--task-sleep-max 15`

You can tune or disable it:

```bash
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --task-sleep-min 8 \
  --task-sleep-max 25
```

```bash
# Disable per-video random waiting
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --task-sleep-min 0 \
  --task-sleep-max 0
```

Additional useful flags:

```bash
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --limit-rate 2M \
  --sleep-interval 3 \
  --max-sleep-interval 8 \
  --sleep-requests 1.5 \
  --proxy socks5://127.0.0.1:7890
```

If your sheet mixes Bilibili and YouTube links, you can keep Bilibili on the direct connection while routing only YouTube through a local proxy:

```bash
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --youtube-proxy http://127.0.0.1:7897
```

Or just provide the local port:

```bash
python3 video_batch_downloader.py \
  --input manifest.xlsx \
  --output-root downloads \
  --youtube-proxy-port 7897
```

This is useful when:

- Bilibili works better without VPN / proxy
- YouTube requires a proxy in your network
- you want one mixed Excel file to run in a single pass

When `--youtube-proxy` or `--youtube-proxy-port` is set, the script logs that YouTube rows use the proxy while Bilibili rows stay direct.

Notes:

- `ffmpeg` is recommended so separate audio/video streams can be merged into `.mp4`. If it is missing, the script logs the warning once and keeps the original container format.
- `downloads/download_report.csv` records `downloaded`, `skipped_existing`, and `failed` rows.
- Please make sure your downloads comply with the target platform's terms and the content owner's rights.

# 🚀 Inference

```bash
# 1. Load Model & Processor
import torch
import requests
from PIL import Image
from io import BytesIO
from transformers import AutoProcessor, AutoModelForImageTextToText
from qwen_vl_utils import process_vision_info

# 1. Load Model & Processor
# Replace with your local model weight directory
model_path = "path/to/ScenePilot_model" 
model = AutoModelForImageTextToText.from_pretrained(
    model_path, 
    torch_dtype=torch.bfloat16, 
    device_map="auto", 
    trust_remote_code=True
)
processor = AutoProcessor.from_pretrained(model_path, trust_remote_code=True)

# 2. Prepare Input
# You can replace this URL with a local image path: Image.open("your_image.jpg")
url = "https://raw.githubusercontent.com/yjwangtj/ScenePilot-Bench/main/assets/sample_drive.jpg"
image = Image.open(BytesIO(requests.get(url).content)).convert("RGB")

# Define Autonomous Driving VQA Prompt
prompt = "Report the current weather, time, road type, how many lanes, if it’s an intersection, and the risk level."

messages = [
    {
        "role": "user",
        "content": [
            {"type": "image", "image": image},
            {"type": "text", "text": prompt}
        ]
    }
]

text_prompt = processor.apply_chat_template(messages, tokenize=False, add_generation_prompt=True)
image_inputs, _ = process_vision_info(messages)
inputs = processor(
    text=[text_prompt], 
    images=image_inputs, 
    return_tensors="pt"
).to(model.device)

# Generate  response
output_ids = model.generate(**inputs, max_new_tokens=128)
answer = processor.batch_decode(
    output_ids[:, inputs.input_ids.shape[1]:], 
    skip_special_tokens=True
)[0]

print(f"--- ScenePilot Result ---\n{answer}")
```


# 📊ScenePilot Benchmark

This repository provides a **two-step evaluation pipeline** for benchmarking Vision-Language Models (VLMs) on the ScenePilot-Bench dataset.


## Step 1: Scene Graph Parsing

Use `scene_graph_parser_final-all.py` to parse **model-generated answers** into a standardized **scene semantic graph representation**.
The parsed results will be saved as a JSON file, which serves as the input for the benchmark scoring stage.

```bash
python scene_graph_parser_final-all.py \
    --input_path path/to/model_outputs.json \
    --output_path path/to/parsed_scene_graph.json
```

**Output**

* A JSON file containing structured scene graph representations extracted from model predictions.


## Step 2: Benchmark Scoring

Use `benchmark_score_final-all.py` to compute **evaluation metrics and final benchmark scores** based on the parsed scene graphs.

```bash
python benchmark_score_final-all.py \
    --input_path path/to/parsed_scene_graph.json \
    --output_dir path/to/save_results
```

Before running the script, please ensure the following paths are properly configured inside the code or via arguments:

* **GPT output log path** (optional, can be commented out if not required)
* **Normalization parameters JSON file**, used for metric scaling and score normalization

**Outputs**

* A JSON file containing detailed evaluation results for each sample
* A CSV file summarizing all benchmark metrics in tabular form, suitable for comparison across models

---

# 🤗 Data and Annotation Pipeline
The detailed annotation pipeline resources of ScenePilot-4K are publicly available on Hugging Face. This release includes the dataset annotations, metadata, and supporting files for the annotation pipeline, which together enable a clearer understanding of how ScenePilot-4K is constructed and organized.

👉 **Access the full dataset and annotation resources here:**  
[ScenePilot-4K on Hugging Face](https://huggingface.co/datasets/larswangtj/ScenePilot-4K/tree/main)

## Citation

```bibtex
@misc{wang2026scenepilot4klargescalefirstpersondataset,
      title={ScenePilot-4K: A Large-Scale First-Person Dataset and Benchmark for Vision-Language Models in Autonomous Driving}, 
      author={Yujin Wang and Yutong Zheng and Wenxian Fan and Tianyi Wang and Hongqing Chu and Li Zhang and Bingzhao Gao and Daxin Tian and Jianqiang Wang and Hong Chen},
      year={2026},
      eprint={2601.19582},
      archivePrefix={arXiv},
      primaryClass={cs.CV},
      url={https://arxiv.org/abs/2601.19582}, 
}
```

## License

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)


This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.
