# Progress Log

## 2026-02-09

### 1) 复现与定位
- 复现API问题：
- 日志出现 `API调用失败，第5/5次: Connection error.` 与 `Comment agent error: Connection error.`。
- 复现训练验证问题：
- 使用 `models/excel_ui_yolo.pt` + `datasets/excel_ui_yolo/dataset.yaml` 验证，得到：
- `mAP50-95=0.0000`
- `recall=0.0000`
- 对照验证：
- 同模型 + `datasets/excel_ui_yolo_auto/dataset.yaml`，得到非零指标（`mAP50-95=0.5781`，`recall=0.9692`）。

### 2) 代码修改
- `lib/DashScopeAPIManager.py`
- 增加 `is_connection_error()`。
- 重试退避在连接错误场景下放慢，减少失败风暴。

- `lib/CommentGenerator.py`
- 增加 API 暂不可用冷却机制（120秒）。
- 连接错误时切换本地降级回复，避免评论流程报错中断。
- 保留本地视觉数据问题的直接回复能力。

- `tools/vision/dataset_sanity.py`（新增）
- 增加数据集检查：图片/标签数量、标注格式、class id、边界范围、平均框数。
- 输出严重问题并可在训练/验证前阻断。

- `tools/vision/train_excel_vision.py`
- 接入 `dataset_sanity` 检查。
- 默认数据集改为 `datasets/excel_ui_yolo_auto/dataset.yaml`。
- 新增 `--allow-bad-data` 强制训练开关。

- `tools/vision/validate_excel_vision.py`
- 接入 `dataset_sanity` 检查。
- 输出 checkpoint 的 `trained_on` 数据集路径。
- 当指标为0时输出路径不一致排查提示与推荐复测命令。
- 新增 `--allow-bad-data` 强制验证开关。

- `tools/vision/test_remote_vision_api.py`（新增）
- 增加远端视觉API可用性测试脚本，支持截屏或指定图片测试。

- `tools/vision/collect_excel_weak_labels.py`
- 默认输出目录改为 `datasets/excel_ui_yolo_auto`。
- 增加窗口尺寸漂移过滤（`--max-size-drift`），减少弱标注错位样本。

- `lib/LocalExcelVision.py`
- 新增 `predict_imgsz` 配置（默认640）减少CPU推理耗时。

- `config/settings.json`
- 增加 `local_vision.predict_imgsz=640`。

- 文档更新：
- `README.md`
- `docs/local_vision_training.md`

### 3) 验证结果
- 语法检查通过：
- `py_compile` 覆盖上述修改文件，均通过。
- 远端视觉API测试：
- `python tools/vision/test_remote_vision_api.py --image runs/excel_vision/screen_test.png`
- 返回 `vision api ok`，说明接口可连通（文本内容存在编码显示异常，不影响连通性判定）。
- 快速验证（320分辨率）：
- `val_quick_auto`: 非零指标（`mAP50-95=0.1653`，快速低分辨率测试）。
- `val_quick_old`: 全零指标并触发“路径不一致”提示。
- 标准验证（640分辨率，历史记录）：
- `datasets/excel_ui_yolo_auto` 上可复现非零指标（`mAP50-95=0.5781`）。

### 4) 已知问题与后续计划
- CPU推理在高分辨率截图下仍可能较慢，已通过 `predict_imgsz` 缓解。
- 单列“少量数据”边界场景仍建议继续用OCR增强做二次确认（后续任务）。
