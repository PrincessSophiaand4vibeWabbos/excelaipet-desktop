import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from tools.vision.dataset_sanity import inspect_dataset, print_stats


def train(args):
    try:
        from ultralytics import YOLO
    except Exception as e:
        raise RuntimeError(f"未安装ultralytics，请先 pip install ultralytics: {e}")

    data_yaml = Path(args.data).resolve()
    if not data_yaml.exists():
        raise FileNotFoundError(f"数据集配置不存在: {data_yaml}")

    stats = inspect_dataset(data_yaml)
    print_stats(stats)
    if stats.severe and not args.allow_bad_data:
        raise RuntimeError("数据集检查存在严重问题，请修复后重试；或使用 --allow-bad-data 强制继续")

    project_dir = Path(args.project).resolve()
    project_dir.mkdir(parents=True, exist_ok=True)

    print(f"[info] data={data_yaml}")
    print(f"[info] model={args.model}, epochs={args.epochs}, imgsz={args.imgsz}, batch={args.batch}")

    model = YOLO(args.model)
    train_result = model.train(
        data=str(data_yaml),
        epochs=args.epochs,
        imgsz=args.imgsz,
        batch=args.batch,
        workers=args.workers,
        project=str(project_dir),
        name=args.name,
        device=args.device,
        patience=args.patience,
        close_mosaic=args.close_mosaic,
        lr0=args.lr0,
        weight_decay=args.weight_decay
    )

    run_dir = project_dir / args.name
    best_pt = run_dir / "weights" / "best.pt"
    if not best_pt.exists():
        raise RuntimeError(f"训练完成但未找到 best.pt: {best_pt}")

    target_model = Path(args.out_model).resolve()
    target_model.parent.mkdir(parents=True, exist_ok=True)
    target_model.write_bytes(best_pt.read_bytes())

    print(f"[done] 训练完成, best model: {best_pt}")
    print(f"[done] 已复制到: {target_model}")
    print("[next] 将 settings.json 中 local_vision.enabled 设为 true 并确保 model_path 指向该文件")
    _ = train_result


def main():
    parser = argparse.ArgumentParser(description="训练本地Excel视觉识别模型(YOLO)")
    parser.add_argument("--data", default="datasets/excel_ui_yolo_auto/dataset.yaml", help="YOLO数据集yaml路径")
    parser.add_argument("--model", default="yolov8n.pt", help="基础模型权重")
    parser.add_argument("--epochs", type=int, default=60, help="训练轮数")
    parser.add_argument("--imgsz", type=int, default=960, help="输入图像尺寸")
    parser.add_argument("--batch", type=int, default=8, help="batch size")
    parser.add_argument("--workers", type=int, default=0, help="数据加载worker")
    parser.add_argument("--device", default="0", help="训练设备，CPU用 cpu")
    parser.add_argument("--project", default="runs/excel_vision", help="训练输出目录")
    parser.add_argument("--name", default="excel_ui", help="本次训练run名称")
    parser.add_argument("--out-model", default="models/excel_ui_yolo.pt", help="导出模型路径")
    parser.add_argument("--patience", type=int, default=30, help="早停耐心")
    parser.add_argument("--close-mosaic", type=int, default=10, help="关闭mosaic的epoch")
    parser.add_argument("--lr0", type=float, default=0.005, help="初始学习率")
    parser.add_argument("--weight-decay", type=float, default=0.0005, help="权重衰减")
    parser.add_argument("--allow-bad-data", action="store_true", help="即使数据集检查异常也强制训练")
    args = parser.parse_args()
    train(args)


if __name__ == "__main__":
    main()
