#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Tiny Vietnamese Contextual Candidate Ranker (Tiny-B Architecture).
Encoder-only candidate scorer designed for fast INT8 CPU inference.
Conforms strictly to TASK requirements: narrow task, listwise scoring, ~14M params.
"""

import math
import torch
import torch.nn as nn
import torch.nn.functional as F


class TinyVietnameseContextRanker(nn.Module):
    def __init__(
        self,
        vocab_size: int = 16000,
        hidden_size: int = 320,
        num_layers: int = 6,
        num_heads: int = 8,
        intermediate_size: int = 1280,
        max_position_embeddings: int = 128,
        dropout_prob: float = 0.1
    ):
        super().__init__()
        self.hidden_size = hidden_size
        self.word_embeddings = nn.Embedding(vocab_size, hidden_size, padding_idx=0)
        self.position_embeddings = nn.Embedding(max_position_embeddings, hidden_size)
        self.embedding_layer_norm = nn.LayerNorm(hidden_size)
        self.dropout = nn.Dropout(dropout_prob)

        encoder_layer = nn.TransformerEncoderLayer(
            d_model=hidden_size,
            nhead=num_heads,
            dim_feedforward=intermediate_size,
            dropout=dropout_prob,
            activation="gelu",
            batch_first=True,
            norm_first=False
        )
        self.encoder = nn.TransformerEncoder(encoder_layer, num_layers=num_layers)
        self.classifier = nn.Sequential(
            nn.Linear(hidden_size, hidden_size // 2),
            nn.GELU(),
            nn.Linear(hidden_size // 2, 1)
        )

        self._init_weights()

    def _init_weights(self):
        for m in self.modules():
            if isinstance(m, nn.Linear):
                nn.init.normal_(m.weight, std=0.02)
                if m.bias is not None:
                    nn.init.zeros_(m.bias)
            elif isinstance(m, nn.Embedding):
                nn.init.normal_(m.weight, std=0.02)
            elif isinstance(m, nn.LayerNorm):
                nn.init.ones_(m.weight)
                nn.init.zeros_(m.bias)

    def forward(self, input_ids: torch.Tensor, attention_mask: torch.Tensor = None) -> torch.Tensor:
        """
        Args:
            input_ids: [batch_candidates, seq_length]
            attention_mask: [batch_candidates, seq_length] (1 for valid tokens, 0 for pad)
        Returns:
            logits: [batch_candidates] unnormalized scores
        """
        batch_size, seq_length = input_ids.shape
        positions = torch.arange(seq_length, device=input_ids.device).unsqueeze(0).expand(batch_size, -1)

        embeddings = self.word_embeddings(input_ids) + self.position_embeddings(positions)
        embeddings = self.embedding_layer_norm(embeddings)
        embeddings = self.dropout(embeddings)

        src_key_padding_mask = None
        if attention_mask is not None:
            # TransformerEncoder uses True for positions to be ignored
            src_key_padding_mask = (attention_mask == 0)

        hidden_states = self.encoder(embeddings, src_key_padding_mask=src_key_padding_mask)

        # Use CLS token at index 0
        cls_rep = hidden_states[:, 0, :]
        logits = self.classifier(cls_rep).squeeze(-1)
        return logits

    def count_parameters(self) -> int:
        return sum(p.numel() for p in self.parameters() if p.requires_grad)


class ListwiseRankingLoss(nn.Module):
    """
    Computes listwise cross-entropy loss over candidate alternatives:
    P(c_i) = softmax(scores)
    Loss = CrossEntropy(scores, correct_index)
    """
    def __init__(self, temperature: float = 1.0):
        super().__init__()
        self.temperature = temperature
        self.ce = nn.CrossEntropyLoss()

    def forward(self, candidate_scores: torch.Tensor, target_index: torch.Tensor) -> torch.Tensor:
        """
        candidate_scores: [batch_size, num_candidates]
        target_index: [batch_size]
        """
        scaled_scores = candidate_scores / self.temperature
        return self.ce(scaled_scores, target_index)


if __name__ == "__main__":
    model = TinyVietnameseContextRanker()
    total_params = model.count_parameters()
    print(f"Tiny-B Architecture instantiated.")
    print(f"Total Trainable Parameters: {total_params:,} (~{total_params / 1e6:.1f}M params)")

    # Smoke inference test
    sample_batch_candidates = 8  # 8 candidates in batch
    sample_seq_len = 64
    x = torch.randint(0, 16000, (sample_batch_candidates, sample_seq_len))
    mask = torch.ones((sample_batch_candidates, sample_seq_len), dtype=torch.long)

    scores = model(x, mask)
    print(f"Input shape: {x.shape}")
    print(f"Candidate output scores shape: {scores.shape}")
    probs = F.softmax(scores, dim=-1)
    print(f"Probabilities: {probs.tolist()}")
