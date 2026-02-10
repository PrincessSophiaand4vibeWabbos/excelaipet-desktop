import argparse
import random
import sys
import time
from pathlib import Path
from typing import Dict, Tuple

from PIL import ImageGrab

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from lib.ExcelUILocator import EXCEL_UI_ELEMENTS, ExcelUILocator


def _box_size_for_key(key: str) -> Tuple[int, int]:
    if key.endswith("_tab"):
        return 86, 28
    if key in {"auto_sum", "merge_center", "freeze_panes"}:
        return 96, 32
    if key in {"name_box", "formula_bar"}:
        return 130, 34
    if key == "cell_area":
        return 360, 220
    if key == "sheet_tab":
        return 100, 28
    return 42, 30


def _clamp(v: float, low: float, high: float) -> float:
    return max(low, min(high, v))


def _to_yolo_bbox(x: float, y: float, w: float, h: float, img_w: int, img_h: int):
    cx = (x + w / 2.0) / img_w
    cy = (y + h / 2.0) / img_h
    nw = w / img_w
    nh = h / img_h
    return (
        _clamp(cx, 0.0, 1.0),
        _clamp(cy, 0.0, 1.0),
        _clamp(nw, 0.0, 1.0),
        _clamp(nh, 0.0, 1.0),
    )


def _write_dataset_yaml(dataset_root: Path, class_names):
    content = [
        f"path: {dataset_root.as_posix()}",
        "train: images/train",
        "val: images/val",
        "",
        f"nc: {len(class_names)}",
        "names:",
    ]
    for idx, name in enumerate(class_names):
        content.append(f"  {idx}: {name}")
    (dataset_root / "dataset.yaml").write_text("\n".join(content), encoding="utf-8")


def collect(args):
    locator = ExcelUILocator()
    if not locator.find_excel_window():
        raise RuntimeError("未找到Excel窗口，请先打开Excel并保持可见")

    left, top, right, bottom = locator.excel_rect
    img_w = right - left
    img_h = bottom - top
    base_w, base_h = img_w, img_h

    dataset_root = Path(args.output).resolve()
    train_img_dir = dataset_root / "images" / "train"
    val_img_dir = dataset_root / "images" / "val"
    train_lbl_dir = dataset_root / "labels" / "train"
    val_lbl_dir = dataset_root / "labels" / "val"
    for d in (train_img_dir, val_img_dir, train_lbl_dir, val_lbl_dir):
        d.mkdir(parents=True, exist_ok=True)

    class_keys = list(EXCEL_UI_ELEMENTS.keys())
    key2idx: Dict[str, int] = {k: i for i, k in enumerate(class_keys)}
    _write_dataset_yaml(dataset_root, class_keys)

    print(f"[info] 采集开始，共 {args.samples} 张；分辨率 {img_w}x{img_h}")
    for i in range(args.samples):
        # 每次循环重新取窗口位置，兼容用户拖动Excel窗口
        if not locator.find_excel_window():
            print("[warn] Excel窗口丢失，跳过此帧")
            time.sleep(args.interval)
            continue
        left, top, right, bottom = locator.excel_rect
        img_w = right - left
        img_h = bottom - top
        if base_w > 0 and base_h > 0:
            drift_w = abs(img_w - base_w) / base_w
            drift_h = abs(img_h - base_h) / base_h
            if drift_w > args.max_size_drift or drift_h > args.max_size_drift:
                print(
                    f"[warn] 跳过第{i + 1}帧: Excel窗口尺寸漂移过大 "
                    f"({img_w}x{img_h} vs {base_w}x{base_h})"
                )
                time.sleep(args.interval)
                continue

        image = ImageGrab.grab(bbox=(left, top, right, bottom))
        split = "val" if random.random() < args.val_ratio else "train"
        stem = f"excel_{i:05d}"
        img_path = (val_img_dir if split == "val" else train_img_dir) / f"{stem}.png"
        lbl_path = (val_lbl_dir if split == "val" else train_lbl_dir) / f"{stem}.txt"
        image.save(img_path)

        lines = []
        for key, element in EXCEL_UI_ELEMENTS.items():
            w, h = _box_size_for_key(key)
            x = float(element.offset_x - w / 2)
            y_offset = element.offset_y if element.offset_y >= 0 else (img_h + element.offset_y)
            y = float(y_offset - h / 2)

            if args.jitter > 0:
                x += random.uniform(-args.jitter, args.jitter)
                y += random.uniform(-args.jitter, args.jitter)

            x = _clamp(x, 0, max(0, img_w - w))
            y = _clamp(y, 0, max(0, img_h - h))

            cx, cy, nw, nh = _to_yolo_bbox(x, y, w, h, img_w, img_h)
            lines.append(f"{key2idx[key]} {cx:.6f} {cy:.6f} {nw:.6f} {nh:.6f}")

        lbl_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"[{i + 1}/{args.samples}] {split} -> {img_path.name}")
        time.sleep(args.interval)

    print(f"[done] 数据集已生成: {dataset_root}")
    print(f"[next] 用以下命令训练模型:\n  python tools/vision/train_excel_vision.py --data {dataset_root / 'dataset.yaml'}")


def main():
    parser = argparse.ArgumentParser(description="采集Excel窗口并生成YOLO弱监督标注")
    parser.add_argument("--output", default="datasets/excel_ui_yolo_auto", help="数据集输出目录")
    parser.add_argument("--samples", type=int, default=240, help="采样数量")
    parser.add_argument("--interval", type=float, default=0.8, help="采样间隔秒")
    parser.add_argument("--val-ratio", type=float, default=0.2, help="验证集比例")
    parser.add_argument("--jitter", type=float, default=4.0, help="标注抖动像素")
    parser.add_argument("--max-size-drift", type=float, default=0.06, help="窗口尺寸允许漂移比例，超出则跳过该帧")
    args = parser.parse_args()
    collect(args)


if __name__ == "__main__":
    main()
