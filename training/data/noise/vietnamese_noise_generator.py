#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Vietnamese Noise Generator for Spelling Correction Dataset Generation.
Supports Telex artifacts, tone confusion, phonetic swaps, and real-word errors.
Ensures a realistic error distribution with a high proportion of NO-ERROR sentences.
"""

from __future__ import annotations

import random
import sys
import unicodedata
from typing import List, Tuple, Dict

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


CONFUSION_INITIAL = {
    "s": ["x"], "x": ["s"],
    "ch": ["tr"], "tr": ["ch"],
    "d": ["gi", "r"], "gi": ["d", "r"], "r": ["d", "gi"],
    "l": ["n"], "n": ["l"]
}

CONFUSION_FINAL = {
    "c": ["t"], "t": ["c"],
    "n": ["ng"], "ng": ["n"]
}

HOI_NGA_MAP = {
    "ả": "ã", "ã": "ả", "ẳ": "ẵ", "ẵ": "ẳ", "ẩ": "ẫ", "ẫ": "ẩ",
    "ẻ": "ẽ", "ẽ": "ẻ", "ể": "ễ", "ễ": "ể", "ỉ": "ĩ", "ĩ": "ỉ",
    "ỏ": "õ", "õ": "ỏ", "ổ": "ỗ", "ỗ": "ổ", "ở": "ỡ", "ỡ": "ở",
    "ủ": "ũ", "ũ": "ủ", "ử": "ữ", "ữ": "ử", "ỷ": "ỹ", "ỹ": "ỷ"
}

REAL_WORD_ERRORS = {
    "bàn giao": "bàn dao",
    "điều khoản": "điều khoảng",
    "xử lý": "sử lý",
    "sáp nhập": "sát nhập",
    "chia sẻ": "chia sẽ",
    "sơ suất": "sơ xuất",
    "hướng dẫn": "hướng dẩn",
    "trân trọng": "chân trọng",
    "rút gọn": "dút gọn",
    "nghỉ việc": "ngỉ việc"
}


class VietnameseNoiseGenerator:
    def __init__(self, seed: int = 42, no_error_ratio: float = 0.60):
        self.rng = random.Random(seed)
        self.no_error_ratio = no_error_ratio

    def corrupt_sentence(self, sentence: str) -> Tuple[str, List[Dict]]:
        """
        Corrupts a sentence according to real-world Vietnamese typo distribution.
        Returns: (corrupted_sentence, list_of_error_spans)
        """
        sentence = unicodedata.normalize("NFC", sentence)

        # 60% probability of returning untouched clean sentence (crucial for low False Positive Rate)
        if self.rng.random() < self.no_error_ratio:
            return sentence, []

        words = sentence.split(" ")
        if not words:
            return sentence, []

        # Check for multi-word real-word errors first
        for correct_phrase, wrong_phrase in REAL_WORD_ERRORS.items():
            if correct_phrase in sentence and self.rng.random() < 0.40:
                idx = sentence.index(correct_phrase)
                corrupted = sentence[:idx] + wrong_phrase + sentence[idx + len(correct_phrase):]
                return corrupted, [{
                    "start": idx,
                    "length": len(wrong_phrase),
                    "original": wrong_phrase,
                    "correct": correct_phrase,
                    "type": "real_word"
                }]

        # Pick 1 or 2 words to corrupt
        error_count = 1 if self.rng.random() < 0.80 else 2
        candidate_indices = [
            i for i, w in enumerate(words)
            if len(w) >= 2 and w.isalpha() and not w.isupper()
        ]

        if not candidate_indices:
            return sentence, []

        selected_indices = self.rng.sample(
            candidate_indices, min(error_count, len(candidate_indices))
        )

        error_spans = []
        new_words = list(words)

        for idx in selected_indices:
            orig_word = words[idx]
            corrupted_word, error_type = self._corrupt_word(orig_word)
            if corrupted_word != orig_word:
                new_words[idx] = corrupted_word

        corrupted_sentence = " ".join(new_words)
        return corrupted_sentence, error_spans

    def _corrupt_word(self, word: str) -> Tuple[str, str]:
        choice = self.rng.random()

        # 1. Hoi / Nga swap (30%)
        if choice < 0.30:
            chars = list(word)
            for i, c in enumerate(chars):
                if c.lower() in HOI_NGA_MAP:
                    swapped = HOI_NGA_MAP[c.lower()]
                    chars[i] = swapped.upper() if c.isupper() else swapped
                    return "".join(chars), "tone_hoi_nga"

        # 2. Initial consonant swap (25%)
        elif choice < 0.55:
            lower = word.lower()
            for k, v in CONFUSION_INITIAL.items():
                if lower.startswith(k):
                    replacement = self.rng.choice(v)
                    is_title = word[0].isupper()
                    new_w = replacement + lower[len(k):]
                    if is_title:
                        new_w = new_w.capitalize()
                    return new_w, "initial_consonant"

        # 3. Final consonant swap (15%)
        elif choice < 0.70:
            lower = word.lower()
            for k, v in CONFUSION_FINAL.items():
                if lower.endswith(k):
                    replacement = self.rng.choice(v)
                    prefix = word[:len(word) - len(k)]
                    return prefix + replacement, "final_consonant"

        # 4. Telex unexpanded artifact (e.g. ngỉ, dd, aw) (15%)
        elif choice < 0.85:
            if word.lower().startswith("ngh"):
                return word[0:2] + word[3:], "telex_missing_h"
            if "đ" in word.lower():
                return word.replace("đ", "dd").replace("Đ", "Dd"), "telex_dd"

        # 5. Character typo / transposition (15%)
        else:
            if len(word) >= 3:
                pos = self.rng.randint(0, len(word) - 2)
                chars = list(word)
                chars[pos], chars[pos + 1] = chars[pos + 1], chars[pos]
                return "".join(chars), "transposition"

        return word, "none"


if __name__ == "__main__":
    generator = VietnameseNoiseGenerator(seed=123)
    sample_texts = [
        "Đồng chí gửi bàn giao hồ sơ trước ngày mai.",
        "Tôi muốn nghỉ việc vào cuối tháng.",
        "Quy trình xử lý văn bản hành chính theo quy định.",
        "Văn phòng thực hiện sáp nhập phòng ban.",
        "Điều khoản hợp đồng đã được thống nhất rõ ràng."
    ]

    print("=== Noise Generation Test ===")
    for text in sample_texts:
        corrupted, errors = generator.corrupt_sentence(text)
        print(f"Original : {text}")
        print(f"Corrupted: {corrupted}")
        print("-" * 50)
