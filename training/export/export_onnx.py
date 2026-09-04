#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Export Tiny Vietnamese Contextual Candidate Ranker to ONNX format.
Inputs:
  - input_ids: [batch_candidates, seq_length] (int64)
  - attention_mask: [batch_candidates, seq_length] (int64)
Outputs:
  - scores: [batch_candidates] (float32)
"""

import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")
import torch

# Add root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))
from training.models.context_ranker import TinyVietnameseContextRanker


def export_to_onnx(output_path: str = "artifacts/models/vietnamese-context-ranker-v1.onnx"):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    model = TinyVietnameseContextRanker()
    model.eval()

    sample_candidates = 8
    seq_length = 64
    dummy_input_ids = torch.randint(0, 16000, (sample_candidates, seq_length), dtype=torch.long)
    dummy_attention_mask = torch.ones((sample_candidates, seq_length), dtype=torch.long)

    dynamic_axes = {
        "input_ids": {0: "batch_candidates", 1: "seq_length"},
        "attention_mask": {0: "batch_candidates", 1: "seq_length"},
        "scores": {0: "batch_candidates"}
    }

    print(f"Exporting PyTorch model to ONNX: {output_path}...")
    torch.onnx.export(
        model,
        (dummy_input_ids, dummy_attention_mask),
        output_path,
        export_params=True,
        opset_version=18,
        do_constant_folding=True,
        input_names=["input_ids", "attention_mask"],
        output_names=["scores"],
        dynamic_axes=dynamic_axes
    )

    size_mb = os.path.getsize(output_path) / (1024 * 1024)
    print(f"Successfully exported ONNX FP32 model. Size: {size_mb:.2f} MB")
    return output_path


if __name__ == "__main__":
    export_to_onnx()
