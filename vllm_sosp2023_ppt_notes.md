# vLLM SOSP 2023 PPT 生成说明

## 输出文件

- `generated/vllm_sosp2023_pagedattention_analysis.pptx`

## 中间代码

- `build_vllm_sosp2023_ppt.py`：使用 `python-pptx` 绘制 8 页 PPT，图表主要由 PPT 原生 shape/table 组成，便于后续编辑。
- `requirements-vllm-ppt.txt`：复现生成脚本所需依赖。

## 复现方式

```bash
python3 -m pip install --user -r requirements-vllm-ppt.txt
python3 build_vllm_sosp2023_ppt.py
```

## 数据与口径

- 主要参考：Kwon et al., *Efficient Memory Management for Large Language Model Serving with PagedAttention*, SOSP 2023。
- 对标对象：HuggingFace Transformers (HF) 与 NVIDIA FasterTransformer (FT)。
- 核心结论：
  - vLLM 吞吐量相对 HF 约 22-28x，相对 FT 约 2-3x。
  - 传统 KV Cache 管理内存浪费率约 60%-80%，vLLM 降至 4% 以下。
  - 实验模型包括 Llama-7B / 13B、OPT-13B / 175B，硬件环境为 NVIDIA A100 80GB。
- PPT 中吞吐量柱状图采用官方倍率校准的 `req/min` 示意值，用于在 6 分钟学术演示中可视化倍率关系。
