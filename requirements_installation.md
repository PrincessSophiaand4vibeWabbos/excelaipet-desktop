# requirements installation

## 1. 环境要求
- 操作系统：Windows 10/11
- Python：建议 3.11 或 3.12
- 建议使用 PowerShell 执行安装命令

## 2. 安装步骤
1. 进入项目目录
```powershell
cd pet/pyCatAI-pet-main/pyCatAI-pet-main
```

2. 创建虚拟环境
```powershell
python -m venv .venv
```

3. 激活虚拟环境
```powershell
.\.venv\Scripts\Activate.ps1
```

4. 升级 pip
```powershell
python -m pip install --upgrade pip
```

5. 安装依赖
```powershell
pip install -r requirements.txt
```

## 3. 当前依赖清单（requirements.txt）
- google-generativeai
- pywin32
- pyttsx3
- Pillow
- openai>=1.0.0
- pandas>=2.0.0
- openpyxl>=3.0.0
- numpy>=1.24.0
- ultralytics>=8.2.0
- opencv-python>=4.8.0

## 4. 安装后快速检查
```powershell
python -c "import tkinter, win32gui, pandas, openpyxl, openai, ultralytics, cv2, PIL; print('ok')"
```

## 5. 常用运行命令
启动主程序：
```powershell
python main.py
```

本地视觉训练/验证：
```powershell
python tools/vision/collect_excel_weak_labels.py --samples 300 --interval 0.8
python tools/vision/train_excel_vision.py --data datasets/excel_ui_yolo_auto/dataset.yaml
python tools/vision/validate_excel_vision.py --model models/excel_ui_yolo.pt --data datasets/excel_ui_yolo_auto/dataset.yaml
```

远端视觉接口测试：
```powershell
python tools/vision/test_remote_vision_api.py
```
