#!/usr/bin/env python3
"""
정지 이미지에서 특정 단어 위에 형광펜(하이라이트) 효과를 합성하고,
캡컷(CapCut) 오버레이 트랙용 투명 PNG를 생성하는 스크립트.

설치:
    pip install easyocr opencv-python pillow numpy

실행:
    python highlighter_overlay.py

출력:
    result_full.png    - 원본 위에 형광펜이 합성된 완성 미리보기
    result_overlay.png - 형광펜만 있는 투명 배경 PNG (원본과 동일 해상도)
"""

import os
import random

import numpy as np
from PIL import Image, ImageDraw

# ── 설정 ─────────────────────────────────────────────────────────────
IMAGE_PATH = "./input.png"
TARGET_WORD = "노가다"

OUTPUT_FULL = "result_full.png"
OUTPUT_OVERLAY = "result_overlay.png"

HIGHLIGHT_COLOR = (255, 235, 59, 110)  # 노란 형광, 반투명 RGBA

BAND_TOP_FRAC = 0.45      # 글자 높이 기준 형광펜이 시작되는 위치 (세로 중앙 부근)
BAND_HEIGHT_FRAC = 0.65   # 형광펜 밴드의 높이 (글자 높이의 약 65%)
SIDE_OVERFLOW_RANGE = (6, 8)  # 좌우로 넘치는 픽셀 범위
TILT_RANGE_DEG = (-2.5, 2.5)  # 살짝 기울이는 각도 범위

RANDOM_SEED = 42
# ─────────────────────────────────────────────────────────────────────


def normalize(text):
    return text.replace(" ", "").strip()


def get_axis_aligned_bbox(box_points):
    xs = [p[0] for p in box_points]
    ys = [p[1] for p in box_points]
    x_min, x_max = min(xs), max(xs)
    y_min, y_max = min(ys), max(ys)
    return x_min, y_min, x_max - x_min, y_max - y_min


def build_highlight_patch(width, height, color, angle_deg):
    """형광펜 스트로크 하나를 그려서 (회전까지 적용한) RGBA 패치로 반환."""
    width = max(1, int(round(width)))
    height = max(1, int(round(height)))

    corner_radius = max(2, int(min(width, height) / 2))

    # 회전 시 잘리지 않도록 여백을 두고 그린다.
    pad = int(max(width, height) * 0.4) + 8
    canvas_w = width + pad * 2
    canvas_h = height + pad * 2

    patch = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(patch)
    rect_box = [pad, pad, pad + width, pad + height]
    draw.rounded_rectangle(rect_box, radius=corner_radius, fill=color)

    rotated = patch.rotate(angle_deg, resample=Image.BICUBIC, expand=True)
    return rotated


def paste_patch(overlay, patch, center_x, center_y):
    """patch를 overlay 캔버스의 (center_x, center_y) 위치에 알파 블렌딩으로 합성."""
    px, py = patch.size
    offset_x = int(round(center_x - px / 2))
    offset_y = int(round(center_y - py / 2))

    layer = Image.new("RGBA", overlay.size, (0, 0, 0, 0))
    layer.paste(patch, (offset_x, offset_y), patch)
    overlay.alpha_composite(layer)


def main():
    if not os.path.exists(IMAGE_PATH):
        raise FileNotFoundError(f"입력 이미지를 찾을 수 없습니다: {IMAGE_PATH}")

    import easyocr

    original = Image.open(IMAGE_PATH).convert("RGBA")
    width, height = original.size
    print(f"[INFO] 이미지 로드 완료: {IMAGE_PATH} ({width}x{height})")

    print("[INFO] EasyOCR 초기화 중 (lang=['ko', 'en'])...")
    reader = easyocr.Reader(["ko", "en"], gpu=False)

    image_np = np.array(original.convert("RGB"))
    results = reader.readtext(image_np)

    print(f"[INFO] 총 {len(results)}개의 텍스트가 검출되었습니다.")
    print("-" * 60)

    norm_target = normalize(TARGET_WORD)
    matched_boxes = []

    for box_points, text, conf in results:
        x, y, w, h = get_axis_aligned_bbox(box_points)
        print(
            f"텍스트: '{text}' | 좌표: (x={x:.0f}, y={y:.0f}, w={w:.0f}, h={h:.0f}) "
            f"| 신뢰도: {conf:.3f}"
        )
        if norm_target in normalize(text):
            matched_boxes.append((x, y, w, h, text, conf))

    print("-" * 60)

    if not matched_boxes:
        print(f"[경고] '{TARGET_WORD}' 와 일치하는 텍스트를 찾지 못했습니다.")
        print("[INFO] 검출된 전체 텍스트 목록:")
        for _, text, conf in results:
            print(f"  - '{text}' (신뢰도 {conf:.3f})")
        return

    print(f"[INFO] '{TARGET_WORD}' 매칭 박스 {len(matched_boxes)}개 발견:")
    for x, y, w, h, text, conf in matched_boxes:
        print(f"  -> '{text}' at (x={x:.0f}, y={y:.0f}, w={w:.0f}, h={h:.0f})")

    random.seed(RANDOM_SEED)
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))

    for x, y, w, h, text, conf in matched_boxes:
        overflow_left = random.uniform(*SIDE_OVERFLOW_RANGE)
        overflow_right = random.uniform(*SIDE_OVERFLOW_RANGE)
        angle = random.uniform(*TILT_RANGE_DEG)

        band_height = h * BAND_HEIGHT_FRAC
        patch_width = w + overflow_left + overflow_right
        patch_height = band_height

        patch = build_highlight_patch(
            width=patch_width,
            height=patch_height,
            color=HIGHLIGHT_COLOR,
            angle_deg=angle,
        )

        center_x = x + w / 2 + (overflow_right - overflow_left) / 2
        center_y = y + h * BAND_TOP_FRAC + band_height / 2

        paste_patch(overlay, patch, center_x, center_y)

    result_full = Image.alpha_composite(original, overlay)

    assert result_full.size == original.size
    assert overlay.size == original.size

    result_full.save(OUTPUT_FULL)
    overlay.save(OUTPUT_OVERLAY)

    print(f"[완료] '{OUTPUT_FULL}' 저장 (원본 + 형광펜 미리보기)")
    print(f"[완료] '{OUTPUT_OVERLAY}' 저장 (투명 배경, 캡컷 오버레이용, {width}x{height})")


if __name__ == "__main__":
    main()
