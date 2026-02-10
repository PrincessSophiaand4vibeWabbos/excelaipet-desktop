# 本地Excel视觉模型训练说明

该项目已接入本地视觉后端（`lib/LocalExcelVision.py`），支持使用你本地训练的 YOLO 模型识别 Excel 界面元素。

## 1. 环境准备

先进入第1项目目录（不是工作区根目录）：

```bash
cd pet/pyCatAI-pet-main/pyCatAI-pet-main
```

然后在项目根目录执行：

```bash
pip install -r requirements.txt
```

至少需要：`ultralytics`、`numpy`、`opencv-python`。

## 2. 自动采集弱监督数据

先打开 Excel 窗口并保持可见，然后运行：

```bash
python tools/vision/collect_excel_weak_labels.py --samples 300 --interval 0.8
```

输出目录默认为 `datasets/excel_ui_yolo_auto/`，会自动生成：

- `images/train`, `images/val`
- `labels/train`, `labels/val`
- `dataset.yaml`

说明：

- 这是弱监督标注（根据预设控件偏移自动生成），用于快速起步。
- 你可在此基础上抽样人工修正标签，以提升准确率。

## 3. 训练模型

```bash
python tools/vision/train_excel_vision.py ^
  --data datasets/excel_ui_yolo_auto/dataset.yaml ^
  --model yolov8n.pt ^
  --epochs 60 ^
  --imgsz 960 ^
  --batch 8 ^
  --device 0
```

训练完成后会把最佳模型复制到：

- `models/excel_ui_yolo.pt`

## 4. 快速预测验证

```bash
python tools/vision/predict_excel_vision.py --model models/excel_ui_yolo.pt
```

默认会抓当前屏幕并保存可视化结果到：

- `runs/excel_vision/predict.jpg`

## 4.1 标准验证（mAP/Precision/Recall）

```bash
python tools/vision/validate_excel_vision.py ^
  --model models/excel_ui_yolo.pt ^
  --data datasets/excel_ui_yolo_auto/dataset.yaml ^
  --imgsz 640 ^
  --batch 8 ^
  --device cpu
```

验证输出目录在：

- `runs/excel_vision/val_run/`

终端会打印：

- `mAP50-95`
- `mAP50`
- `precision`
- `recall`

若你看到 `mAP50-95=0` 且 `recall=0`：

- 先检查是否把模型拿去验证了另一套数据集（路径不一致会导致近似全错）。
- 用脚本提示的 `checkpoint trained_on=...` 路径做一次复测。
- 建议统一使用 `datasets/excel_ui_yolo_auto/dataset.yaml` 做训练和验证。

## 5. 在桌宠中启用本地视觉

编辑 `config/settings.json`：

```json
"local_vision": {
  "enabled": true,
  "model_path": "models/excel_ui_yolo.pt",
  "conf_threshold": 0.25,
  "max_det": 20
}
```

然后运行：

```bash
python main.py
```

启用后，截图评论会优先走本地模型识别；本地识别不可用时再回退远端视觉接口。

## 6. 本地模型效果不佳时，切换远端视觉API

1) 在 `config/settings.json` 中暂时关闭本地视觉：

```json
"local_vision": {
  "enabled": false
}
```

2) 保持已有的 `api_key`、`base_url`、`vision_model` 配置不变。

3) 先测试远端视觉接口：

```bash
python tools/vision/test_remote_vision_api.py
```

4) 若返回 `vision api ok`，再运行主程序：

```bash
python main.py
```
