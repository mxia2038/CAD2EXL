# P&ID管道数据提取工具

一个用于从P&ID图纸中自动提取管道号信息并生成Excel报告的工具。

## ✨ 功能特点

- 🔍 **自动识别管道号** - 从DWG文件中智能识别管道号格式
- 📊 **Excel报告生成** - 生成包含管道详细信息的Excel报告
- 🧪 **智能相态判断** - 基于介质名称自动判断液相/气相
- 🎯 **用户友好界面** - 简洁的图形用户界面
- 📝 **自定义介质代码** - 支持从Excel文件加载介质代码映射
- 🔄 **实时进度显示** - 处理过程可视化

## 📋 系统要求

- Windows 操作系统
- 已安装 AutoCAD（用于读取DWG文件）
- Python 3.9+ （开发环境）

## 🚀 快速开始

### 方式一：使用预编译版本（推荐）

1. 下载最新的 [Release](../../releases) 版本
2. 解压后运行 `PID_Extractor.exe`
3. 按照界面提示选择文件并开始提取

### 方式二：从源码运行

1. 克隆仓库
```bash
git clone https://github.com/your-username/CAD2EXL.git
cd CAD2EXL
```

2. 安装依赖
```bash
pip install -r requirements.txt
```

3. 运行程序
```bash
python pid_extractor_gui.py
```

## 📖 使用说明

### 1. 准备文件

- **DWG文件**: P&ID图纸文件
- **介质代码文件**: Excel文件，包含介质代码和名称的映射关系
  - 第一列：介质代码（如：BR、BRC、BP等）
  - 第二列：介质名称（如：锅炉水、锅炉给水等）

### 2. 操作步骤

1. 启动程序
2. 选择DWG文件
3. 选择介质代码Excel文件
4. 指定输出文件位置
5. 点击"开始提取"
6. 等待处理完成

### 3. 输出结果

生成的Excel文件包含以下信息：
- 管道号
- 管径
- 管道等级
- 保温型式
- 介质名称
- 相态

## 🛠️ 开发

### 项目结构

```
CAD2EXL/
├── pid_extractor_gui.py      # GUI版本主程序
├── pid_extractor.spec        # PyInstaller打包配置
├── requirements.txt          # Python依赖
├── CLAUDE.md                 # 项目开发文档
├── test/
│   ├── code.xlsx            # 测试用介质代码文件
│   └── test.dwg             # 测试DWG文件
└── dist/                    # 发布文件（打包后生成）
    ├── PID_Extractor.exe
    ├── 介质代码示例.xlsx
    └── 使用说明.txt
```

### 技术栈

- **界面**: tkinter (Python标准库)
- **数据处理**: pandas, openpyxl
- **AutoCAD接口**: pyautocad
- **打包**: PyInstaller

### 构建可执行文件

```bash
# 安装打包工具
pip install pyinstaller

# 打包
pyinstaller pid_extractor.spec
```

### 主要依赖

- pandas>=1.5.0
- openpyxl>=3.0.0
- pyautocad>=0.2.0
- comtypes (Windows COM组件)

## 📄 管道号格式

工具支持以下管道号格式：
```
装置号-管径-介质代码-管道号-管道等级-保温型式
示例: 1001-100-BR-123456A-C1C-A
```

## 🤝 贡献

欢迎提交Issue和Pull Request！

## 📝 许可证

MIT License

## 📞 支持

如果您遇到问题，请：
1. 查看使用说明文件
2. 检查系统要求是否满足
3. 提交Issue描述问题

## 🔄 版本历史

查看 [CHANGELOG.md](CHANGELOG.md) 了解详细的版本更新记录。

---

⚡ 由 [Claude Code](https://claude.ai/code) 协助开发