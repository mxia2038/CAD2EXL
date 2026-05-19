# CAD2EXL Project Environment

## Python Environment
- Python path: `/mnt/d/CAD2EXL/.venv/Scripts/python.exe`
- Pip path: `/mnt/d/CAD2EXL/.venv/Scripts/pip.exe`
- Virtual environment: `.venv`

## Project Structure
```
/mnt/d/CAD2EXL/
├── .venv/                  # Python virtual environment
├── extractor_core.py       # 核心逻辑（项目类型、正则、解析、Excel输出）
├── pid_extractor.py        # CLI 入口（仅命令行解析 + 调用 core）
├── pid_extractor_gui.py    # GUI 入口（仅界面逻辑 + 调用 core）
├── pid_extractor.spec      # PyInstaller 打包配置
├── requirements.txt        # 依赖声明
├── fig/
│   └── logo.jpg            # 界面 Logo
└── test/                   # 用户工作目录（内容不纳入 git）
```

## Dependencies
- pyautocad>=0.2.0 (AutoCAD COM 接口，读取 DWG)
- pandas>=1.5.0 (数据处理)
- openpyxl>=3.0.0 (Excel 导出)
- Pillow>=9.0.0 (Logo 图像处理)
- tkinterdnd2>=0.4.0 (GUI 拖拽功能)

## Usage Commands
```bash
# Run GUI version (recommended)
/mnt/d/CAD2EXL/.venv/Scripts/python.exe pid_extractor_gui.py

# Run CLI version (default: 巨化项目)
/mnt/d/CAD2EXL/.venv/Scripts/python.exe pid_extractor.py

# Run CLI version with specific project type
/mnt/d/CAD2EXL/.venv/Scripts/python.exe pid_extractor.py --project-type "天华项目"

# CLI with custom files
/mnt/d/CAD2EXL/.venv/Scripts/python.exe pid_extractor.py \
  --project-type "巨化项目" --dwg-file "custom.dwg" --output-file "result.xlsx"

# Build exe
/mnt/d/CAD2EXL/.venv/Scripts/pyinstaller.exe pid_extractor.spec

# Install dependencies
/mnt/d/CAD2EXL/.venv/Scripts/pip.exe install -r requirements.txt
```

## Architecture
核心逻辑集中在 `extractor_core.py`，CLI 和 GUI 仅保留各自的入口代码。
**新增项目类型只需修改 `extractor_core.py` 一个文件**，无需同步多处。

`extractor_core.py` 暴露的主要接口：
- `SUPPORTED_PROJECT_TYPES` — 支持的项目列表
- `PROJECT_FORMAT_EXAMPLES` — 各项目格式示例字典
- `extract_text_from_dwg(dwg_path, log_fn=None)` — 从 DWG 提取文本
- `find_pipeline_numbers(text_entities, project_type, log_fn=None)` — 正则匹配管道号
- `load_medium_codes(code_file_path)` — 加载介质代码 Excel
- `parse_pipeline_number(pipeline_number, medium_codes, project_type)` — 解析各字段
- `create_excel_output(pipeline_data, output_path)` — 写出 Excel 报告

`log_fn` 参数接受任意回调，CLI 传 `logger.info`，GUI 传 `self.log_message`。

## Project Types Supported

### 巨化项目 (Juhua Project)
- Format: `4101BRR-02457-200-03CBMB1-H`
- Output: `4101BRR-02457`
- Pattern: `(\d{4}[A-Z0-9]{1,4})-([A-Z0-9]{4,6})-(\d{2,4})-(\d{2}[A-Z0-9]{3,6})-([A-Z]{1,2})`

### 乌兹项目 (Uzbekistan Project)
- Format: `PA-2001002A-100-C1C-N`
- Output: `PA-2001002A`
- Pattern: `([A-Z0-9]{1,4})-([A-Z0-9]{4,8})-(\d{2,4})-([A-Z0-9]{1,4})-([A-Z0-9]{1,2})`

### 天华项目 (Tianhua Project)
- Format: `01PL-216061-125-C22S-H`
- Output: `01PL-216061`
- Pattern: `(\d{2}[A-Z]{2,4})-(\d{5,7})-(\d{2,4})-([A-Z0-9]{2,6})-([A-Z]{1,2})`

### 金昱元项目 (Jinyuyuan Project)
- Format: `01PC03012-100-C12S-H`
- Output: `01PC03012`
- Pattern: `(\d{2}[A-Z]{2}\d{5,7})-(\d{2,4})-([A-Z0-9]{2,6})-([A-Z]{1,2})`
- 注意：第一段将装置号、介质代码、管道号合并，无内部连字符

## Known Issues
- DWG 读取依赖本地安装的 AutoCAD（通过 COM 接口）
- DWG 2018+ 格式建议先在 AutoCAD 中另存为低版本再处理
