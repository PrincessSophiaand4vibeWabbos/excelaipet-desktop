# Task Plan

## 阶段1：接口稳定性与容错
- [x] 复现 `Connection error`，确认报错链路发生在评论文本生成阶段（`CommentGenerator -> DashScopeAPIManager.generate_text`）。
- [x] 增加连接错误识别与冷却机制，避免连续5次失败后仍频繁请求。
- [x] 连接失败时启用本地降级评论，保证截图评论功能可继续使用。

## 阶段2：本地视觉训练复现与优化
- [x] 复现 `mAP50-95=0 / recall=0`。
- [x] 定位根因：模型训练数据集与验证数据集路径不一致（`excel_ui_yolo_auto` vs `excel_ui_yolo`）。
- [x] 在训练/验证脚本中加入数据集一致性与质量检查。
- [x] 统一采集/训练/验证默认数据集为 `datasets/excel_ui_yolo_auto`。
- [x] 优化本地推理速度参数（`predict_imgsz`），减轻CPU推理卡顿。

## 阶段3：交付文档与可追溯记录
- [x] 创建全景任务清单 `task_plan.md`。
- [x] 创建项目发现文档 `findings.md`（含查阅资料、方案对比、接口调整点）。
- [x] 创建操作日志 `progress.md`（含训练、验证结果与问题记录）。
- [x] 给出图像识别API兜底方案并提供测试脚本 `tools/vision/test_remote_vision_api.py`。

## 阶段4：后续可选增强（未执行）
- [ ] 增加基于OCR的单元格文本密度检测，提高“少量数据”场景准确率。
- [ ] 增加“当前选中列”识别，支持“这个列是否有数据”无列号问法。
