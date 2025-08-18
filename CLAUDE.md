# CAD2EXL Project Environment

## Python Environment
- Python path: `/mnt/d/CAD2EXL/.venv/Scripts/python.exe`
- Pip path: `/mnt/d/CAD2EXL/.venv/Scripts/pip.exe`
- Virtual environment: `.venv`

## Project Structure
```
/mnt/d/CAD2EXL/
├── .venv/              # Python virtual environment
├── pid_extractor.py    # CLI extraction script  
├── pid_extractor_gui.py # GUI extraction script
├── requirements.txt    # Dependencies
├── test/
│   └── test.dwg       # Test DWG file
└── *.xlsx             # Generated Excel files
```

## Dependencies
- ezdxf>=1.1.0 (DXF/DWG file reading)
- pandas>=1.5.0 (Data processing)
- openpyxl>=3.0.0 (Excel export)
- pathlib2>=2.3.0 (Path utilities)

## Usage Commands
```bash
# Run GUI version (recommended)
/mnt/d/CAD2EXL/.venv/Scripts/python.exe pid_extractor_gui.py

# Run CLI version (default: 巨化项目)
/mnt/d/CAD2EXL/.venv/Scripts/python.exe pid_extractor.py

# Run CLI version with specific project type
/mnt/d/CAD2EXL/.venv/Scripts/python.exe pid_extractor.py --project-type "乌兹项目"

# CLI with custom files
/mnt/d/CAD2EXL/.venv/Scripts/python.exe pid_extractor.py --project-type "巨化项目" --dwg-file "custom.dwg" --output-file "result.xlsx"

# Install dependencies
/mnt/d/CAD2EXL/.venv/Scripts/pip.exe install -r requirements.txt
```

## Recent Improvements (v1.4.0 - 2025-08-18)
- **Dual Project Support**: Added support for "乌兹项目" pipeline numbering standard alongside existing "巨化项目"
- **Improved User Interface**: 
  - Replaced radio buttons with dropdown selector for project types
  - Dynamic format examples that update based on selected project
  - Optimized layout with 90% settings area and 10% results area
  - Enlarged drag-and-drop areas (80px height)
- **Enhanced Drag-and-Drop**: Fixed handling of filenames with spaces
- **CLI Enhancement**: Added `--project-type` parameter for command-line project selection
- **Extensible Architecture**: Easy addition of new project types through centralized configuration

## Project Types Supported
### 巨化项目 (Juhua Project)
- Format: `4101BRR-02457-200-03CBMB1-H`
- Output: `4101BRR-02457`
- Pattern: Unit+Medium-PipeNumber-Diameter-Grade-Insulation

### 乌兹项目 (Uzbekistan Project)  
- Format: `PA-2001002A-100-C1C-N`
- Output: `PA-2001002A`
- Pattern: Medium-PipeNumber-Diameter-Grade-Insulation

## Development Notes
⚠️ **CRITICAL**: Always update BOTH files when making code changes:
- `pid_extractor.py` (CLI version)
- `pid_extractor_gui.py` (GUI version)

## Known Issues
- DWG 2018/2019/2020 format has limited support with ezdxf
- Consider converting DWG to DXF format for better compatibility