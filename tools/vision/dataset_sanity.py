from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple


@dataclass
class DatasetStats:
    root: Path
    train_images: int
    val_images: int
    train_labels: int
    val_labels: int
    nc: int
    invalid_label_lines: int
    out_of_range_class_ids: int
    total_boxes: int
    avg_boxes_per_image: float
    warnings: List[str]

    @property
    def severe(self) -> bool:
        return any(w.startswith("SEVERE:") for w in self.warnings)


def _load_yaml(path: Path) -> Dict:
    try:
        import yaml
    except Exception as exc:
        raise RuntimeError(f"缺少PyYAML依赖，请安装后重试: {exc}")
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"dataset yaml 格式错误: {path}")
    return data


def _resolve_dataset_root(yaml_path: Path, data: Dict) -> Path:
    raw = str(data.get("path", "")).strip()
    if raw:
        root = Path(raw)
        if not root.is_absolute():
            root = (yaml_path.parent / root).resolve()
        return root
    return yaml_path.parent.resolve()


def _iter_label_files(labels_dir: Path):
    if not labels_dir.exists():
        return []
    return sorted(labels_dir.glob("*.txt"))


def _count_files(path: Path, patterns: Tuple[str, ...]) -> int:
    if not path.exists():
        return 0
    total = 0
    for p in patterns:
        total += len(list(path.glob(p)))
    return total


def inspect_dataset(yaml_path: Path) -> DatasetStats:
    yaml_path = yaml_path.resolve()
    data = _load_yaml(yaml_path)
    root = _resolve_dataset_root(yaml_path, data)

    train_rel = str(data.get("train", "images/train"))
    val_rel = str(data.get("val", "images/val"))
    train_images_dir = root / train_rel
    val_images_dir = root / val_rel
    train_labels_dir = root / train_rel.replace("images", "labels")
    val_labels_dir = root / val_rel.replace("images", "labels")

    train_images = _count_files(train_images_dir, ("*.png", "*.jpg", "*.jpeg"))
    val_images = _count_files(val_images_dir, ("*.png", "*.jpg", "*.jpeg"))
    train_labels = len(_iter_label_files(train_labels_dir))
    val_labels = len(_iter_label_files(val_labels_dir))
    nc = int(data.get("nc", 0) or 0)

    invalid_label_lines = 0
    out_of_range_class_ids = 0
    total_boxes = 0

    for label_file in _iter_label_files(train_labels_dir) + _iter_label_files(val_labels_dir):
        lines = [ln.strip() for ln in label_file.read_text(encoding="utf-8").splitlines() if ln.strip()]
        for line in lines:
            parts = line.split()
            if len(parts) != 5:
                invalid_label_lines += 1
                continue
            total_boxes += 1
            try:
                cls_id = int(parts[0])
                cx, cy, w, h = [float(v) for v in parts[1:]]
            except Exception:
                invalid_label_lines += 1
                continue

            if nc > 0 and (cls_id < 0 or cls_id >= nc):
                out_of_range_class_ids += 1
            if not (0 <= cx <= 1 and 0 <= cy <= 1 and 0 < w <= 1 and 0 < h <= 1):
                invalid_label_lines += 1

    total_images = max(train_images + val_images, 1)
    avg_boxes_per_image = total_boxes / total_images

    warnings: List[str] = []
    if train_images == 0 or val_images == 0:
        warnings.append("SEVERE: 训练或验证图片数量为0")
    if train_labels == 0 or val_labels == 0:
        warnings.append("SEVERE: 训练或验证标签数量为0")
    if train_images != train_labels:
        warnings.append(f"WARN: train 图片({train_images})与标签({train_labels})数量不一致")
    if val_images != val_labels:
        warnings.append(f"WARN: val 图片({val_images})与标签({val_labels})数量不一致")
    if invalid_label_lines > 0:
        warnings.append(f"SEVERE: 存在无效标注行 {invalid_label_lines}")
    if out_of_range_class_ids > 0:
        warnings.append(f"SEVERE: 存在越界class id {out_of_range_class_ids}")
    if avg_boxes_per_image < 1.0:
        warnings.append(f"WARN: 平均每图标注框偏少 ({avg_boxes_per_image:.2f})")

    return DatasetStats(
        root=root,
        train_images=train_images,
        val_images=val_images,
        train_labels=train_labels,
        val_labels=val_labels,
        nc=nc,
        invalid_label_lines=invalid_label_lines,
        out_of_range_class_ids=out_of_range_class_ids,
        total_boxes=total_boxes,
        avg_boxes_per_image=avg_boxes_per_image,
        warnings=warnings,
    )


def print_stats(stats: DatasetStats):
    print(f"[data] root={stats.root}")
    print(
        "[data] train_images={0}, val_images={1}, train_labels={2}, val_labels={3}, nc={4}, boxes={5}, avg_boxes_per_image={6:.2f}".format(
            stats.train_images,
            stats.val_images,
            stats.train_labels,
            stats.val_labels,
            stats.nc,
            stats.total_boxes,
            stats.avg_boxes_per_image,
        )
    )
    if stats.warnings:
        for item in stats.warnings:
            print(f"[data] {item}")
    else:
        print("[data] 数据集检查通过")
