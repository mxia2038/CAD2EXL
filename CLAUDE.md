# CAD2EXL Project Environment

## Python Environment
- Python path: `/mnt/d/CAD2EXL/.venv/Scripts/python.exe`
- Pip path: `/mnt/d/CAD2EXL/.venv/Scripts/pip.exe`
- Virtual environment: `.venv`

## Project Structure
```
/mnt/d/CAD2EXL/
├── .venv/              # Python virtual environment
├── pid_pipeline_extractor.py  # Main extraction script
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
# Run with virtual environment
/mnt/d/CAD2EXL/.venv/Scripts/python.exe pid_pipeline_extractor.py

# Install dependencies
/mnt/d/CAD2EXL/.venv/Scripts/pip.exe install -r requirements.txt
```

## Known Issues
- DWG 2018/2019/2020 format has limited support with ezdxf
- Consider converting DWG to DXF format for better compatibility