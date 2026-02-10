import argparse
import sys
from pathlib import Path

from PIL import ImageGrab

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main():
    parser = argparse.ArgumentParser(description="本地Excel视觉模型快速验证")
    parser.add_argument("--model", default="models/excel_ui_yolo.pt", help="模型路径")
    parser.add_argument("--image", default="", help="测试图片路径；为空则抓取当前屏幕")
    parser.add_argument("--conf", type=float, default=0.25, help="置信度阈值")
    parser.add_argument("--save", default="runs/excel_vision/predict.jpg", help="可视化结果输出路径")
    args = parser.parse_args()

    try:
        from ultralytics import YOLO
    except Exception as e:
        raise RuntimeError(f"未安装ultralytics，请先 pip install ultralytics: {e}")

    model_path = Path(args.model).resolve()
    if not model_path.exists():
        raise FileNotFoundError(f"模型不存在: {model_path}")

    model = YOLO(str(model_path))
    if args.image:
        source = str(Path(args.image).resolve())
    else:
        shot_path = Path("runs/excel_vision/screen_test.png")
        shot_path.parent.mkdir(parents=True, exist_ok=True)
        ImageGrab.grab().save(shot_path)
        source = str(shot_path)

    results = model.predict(source=source, conf=args.conf, verbose=False)
    if not results:
        print("[warn] 无预测结果")
        return

    img = results[0].plot()
    save_path = Path(args.save).resolve()
    save_path.parent.mkdir(parents=True, exist_ok=True)
    from PIL import Image

    Image.fromarray(img).save(save_path)
    print(f"[done] 预测完成，结果图: {save_path}")


if __name__ == "__main__":
    main()
