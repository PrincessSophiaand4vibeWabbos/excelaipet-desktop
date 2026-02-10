import argparse
import json
import sys
from pathlib import Path

from PIL import ImageGrab

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from lib.DashScopeAPIManager import DashScopeAPIManager


def load_settings():
    cfg_path = PROJECT_ROOT / "config" / "settings.json"
    if not cfg_path.exists():
        raise FileNotFoundError(f"配置文件不存在: {cfg_path}")
    return json.loads(cfg_path.read_text(encoding="utf-8"))


def test_remote_vision(image_path: str = ""):
    cfg = load_settings()
    api_key = cfg.get("api_key", "").strip()
    base_url = cfg.get("base_url", "").strip()
    model = cfg.get("vision_model", "").strip() or cfg.get("model", "").strip()
    if not api_key:
        raise RuntimeError("settings.json 未配置 api_key")

    if image_path:
        img_bytes = Path(image_path).resolve().read_bytes()
        src = str(Path(image_path).resolve())
    else:
        tmp = PROJECT_ROOT / "runs" / "excel_vision" / "remote_api_test.png"
        tmp.parent.mkdir(parents=True, exist_ok=True)
        ImageGrab.grab().save(tmp)
        img_bytes = tmp.read_bytes()
        src = str(tmp)

    mgr = DashScopeAPIManager(api_key=api_key, model=model, base_url=base_url)
    ok, result, error = mgr.process_image(
        image_bytes=img_bytes,
        system_prompt="你是Excel屏幕识别助手。",
        user_prompt="请判断画面是否包含Excel表格数据，并给出一句简短说明。",
        max_tokens=120,
        model=model,
    )
    print(f"[info] image={src}")
    if ok:
        print(f"[done] vision api ok: {result}")
    else:
        print(f"[fail] vision api error: {error}")


def main():
    parser = argparse.ArgumentParser(description="测试远端图像识别API可用性")
    parser.add_argument("--image", default="", help="测试图片路径，空则截屏")
    args = parser.parse_args()
    test_remote_vision(image_path=args.image)


if __name__ == "__main__":
    main()
