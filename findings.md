# Findings

## 1. 查阅与排查范围
- 代码文件：
- `lib/DashScopeAPIManager.py`
- `lib/CommentGenerator.py`
- `lib/LocalExcelVision.py`
- `tools/vision/collect_excel_weak_labels.py`
- `tools/vision/train_excel_vision.py`
- `tools/vision/validate_excel_vision.py`
- `tools/vision/test_remote_vision_api.py`
- 文档文件：
- `docs/local_vision_training.md`
- `README.md`

## 2. 关键发现
- `Connection error` 主要发生在文本评论链路：视觉分析成功后，评论生成仍调用远端文本API，导致失败。
- 当 API 连续失败时，原流程会继续尝试远端调用，用户侧感知为“评论失败”。
- `mAP=0` 并不总是“模型没训练”，在当前项目中已复现为“验证集路径与训练集路径不一致”导致的全错评估。
- 当前存在两套数据集：
- `datasets/excel_ui_yolo_auto`（模型 checkpoint 显示训练来源）
- `datasets/excel_ui_yolo`（用于验证时得到 0 指标）

## 3. 试过的方案与结果
- 方案A：使用 `models/excel_ui_yolo.pt` 验证 `datasets/excel_ui_yolo`
- 结果：`mAP50-95=0.0000, recall=0.0000`
- 方案B：使用同一模型验证 `datasets/excel_ui_yolo_auto`
- 结果：可复现非零指标（示例：`mAP50-95=0.5781, recall=0.9692`）
- 结论：模型可用，问题在数据集一致性与流程默认值不统一。

## 4. 已调整的API接口点
- `DashScopeAPIManager`：
- 新增连接错误判定 `is_connection_error()`。
- 重试退避对连接错误使用更慢基线，降低抖动时失败风暴。
- `CommentGenerator`：
- 新增API暂不可用冷却窗口，避免短时间重复失败。
- 连接失败后切换本地降级评论，避免报错中断用户交互。

## 5. 如果本地图像模型效果欠佳的API方案
- 推荐方案：继续使用当前 OpenAI 兼容接口（`base_url + vision_model`）作为远端视觉兜底，保持现有 `api_key` 生态不变。
- 调用入口：`DashScopeAPIManager.process_image()`（`image_url` data URL 方式）。
- 快速自检命令：`python tools/vision/test_remote_vision_api.py`
- 本地实测：`python tools/vision/test_remote_vision_api.py --image runs/excel_vision/screen_test.png` 返回 `vision api ok`。
- 建议策略：
- 本地优先，远端兜底（已实现）。
- 对“低置信度/疑似”结果允许远端复核（可后续加阈值开关）。
- 当连接错误时自动降级本地，不阻塞主流程（已实现）。

## 6. 当前建议的稳定命令
- 采集：`python tools/vision/collect_excel_weak_labels.py --samples 300 --interval 0.8`
- 训练：`python tools/vision/train_excel_vision.py --data datasets/excel_ui_yolo_auto/dataset.yaml`
- 验证：`python tools/vision/validate_excel_vision.py --model models/excel_ui_yolo.pt --data datasets/excel_ui_yolo_auto/dataset.yaml`
