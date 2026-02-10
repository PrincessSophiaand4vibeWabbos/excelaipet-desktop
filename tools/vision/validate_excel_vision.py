import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from tools.vision.dataset_sanity import inspect_dataset, print_stats


def validate(args):
    try:
        from ultralytics import YOLO
    except Exception as e:
        raise RuntimeError(f"未安装ultralytics，请先执行 pip install ultralytics: {e}")

    model_path = Path(args.model).resolve()
    data_yaml = Path(args.data).resolve()
    if not model_path.exists():
        raise FileNotFoundError(f"模型不存在: {model_path}")
    if not data_yaml.exists():
        raise FileNotFoundError(f"数据集配置不存在: {data_yaml}")

    stats = inspect_dataset(data_yaml)
    print_stats(stats)
    if stats.severe and not args.allow_bad_data:
        raise RuntimeError("数据集检查存在严重问题，请修复后重试；或使用 --allow-bad-data 强制继续")

    project_dir = Path(args.project).resolve()
    project_dir.mkdir(parents=True, exist_ok=True)

    model = YOLO(str(model_path))
    model_train_data = str(model.overrides.get("data", "")).strip()
    if model_train_data:
        print(f"[model] checkpoint trained_on={model_train_data}")
        if Path(model_train_data).resolve() != data_yaml:
            print("[model] WARN: 当前验证数据集与模型训练数据集路径不一致")
    metrics = model.val(
        data=str(data_yaml),
        split=args.split,
        imgsz=args.imgsz,
        batch=args.batch,
        workers=args.workers,
        device=args.device,
        conf=args.conf,
        iou=args.iou,
        project=str(project_dir),
        name=args.name,
        verbose=False,
        plots=True,
    )

    print("[done] 验证完成")
    print(f"[metric] mAP50-95: {metrics.box.map:.4f}")
    print(f"[metric] mAP50: {metrics.box.map50:.4f}")
    print(f"[metric] precision: {metrics.box.mp:.4f}")
    print(f"[metric] recall: {metrics.box.mr:.4f}")
    print(f"[output] {metrics.save_dir}")
    if metrics.box.map50 <= 1e-6 and model_train_data:
        print("[hint] mAP50为0，优先检查: 1) 是否用错数据集路径 2) 标注是否与截图分辨率匹配")
        print(f"[hint] 推荐先用模型原训练集复测: python tools/vision/validate_excel_vision.py --model {model_path} --data \"{model_train_data}\"")


def main():
    parser = argparse.ArgumentParser(description="验证本地Excel视觉模型(YOLO)")
    parser.add_argument("--model", default="models/excel_ui_yolo.pt", help="模型路径")
    parser.add_argument("--data", default="datasets/excel_ui_yolo_auto/dataset.yaml", help="YOLO数据集yaml路径")
    parser.add_argument("--split", default="val", help="验证集划分，默认val")
    parser.add_argument("--imgsz", type=int, default=640, help="输入图像尺寸")
    parser.add_argument("--batch", type=int, default=8, help="batch size")
    parser.add_argument("--workers", type=int, default=0, help="数据加载worker")
    parser.add_argument("--device", default="cpu", help="验证设备，CPU用cpu")
    parser.add_argument("--conf", type=float, default=0.25, help="置信度阈值")
    parser.add_argument("--iou", type=float, default=0.7, help="NMS IoU阈值")
    parser.add_argument("--project", default="runs/excel_vision", help="验证输出目录")
    parser.add_argument("--name", default="val_run", help="验证run名称")
    parser.add_argument("--allow-bad-data", action="store_true", help="即使数据集检查异常也强制验证")
    args = parser.parse_args()
    validate(args)


if __name__ == "__main__":
    main()
