# P&ID管道数据提取工具

从P&ID图纸中自动提取管道号信息并生成Excel报告。

## ✨ 功能特点

- 🏗️ **多项目标准支持** — 内置四种管道编号标准，可轻松扩展
- 🔍 **智能管道号识别** — 正则表达式匹配，支持完整和简化格式
- 📊 **Excel报告生成** — 自动填充管道号、管径、等级、介质、相态
- 🧪 **智能相态判断** — 根据介质名称自动判断液相/气相
- 🖱️ **拖拽支持** — 文件拖拽到界面即可加载，支持含空格文件名
- 📋 **最近文件记忆** — 每种文件类型记住最近5个路径
- 🔧 **CLI工具** — 支持命令行批量操作

## 📋 系统要求

- Windows 操作系统
- 已安装 AutoCAD（用于读取DWG文件）
- Python 3.9+（源码运行时需要）

## 🚀 快速开始

### 方式一：使用预编译版本（推荐）

1. 下载最新的 [Release](../../releases) 版本
2. 运行 `PID_Extractor.exe`
3. 按界面提示选择文件并开始提取

### 方式二：从源码运行

```bash
git clone https://github.com/mxia2038/CAD2EXL.git
cd CAD2EXL
pip install -r requirements.txt

# GUI版本
python pid_extractor_gui.py

# CLI版本
python pid_extractor.py --project-type "巨化项目"
```

## 📖 使用说明

### 操作步骤

1. 启动程序
2. **选择项目类型** — 从下拉菜单选择对应项目
3. **加载文件**（支持拖拽、浏览按钮两种方式）：
   - DWG源文件（P&ID图纸）
   - 介质代码Excel文件（第一列：代码，第二列：名称）
4. 指定输出文件位置
5. 点击「🚀 开始提取数据」

💡 拖拽时有视觉反馈：蓝色 = 可拖放，绿色 = 成功，红色 = 格式错误

### 输出列

| 列名 | 说明 |
|------|------|
| 管道号 | 提取并简化后的唯一标识 |
| 管径 | 公称直径 |
| 管道等级 | 压力等级代码 |
| 保温等级 | 保温类型代码 |
| 介质名称 | 从介质代码文件查找 |
| 相态 | 气相 / 液相 / 未知相态 |

## 📄 支持的管道号格式

### 🏗️ 巨化项目
```
格式: 装置号+介质代码-管道号-管径-管道等级-保温等级
示例: 4101BRR-02457-200-03CBMB1-H
输出: 4101BRR-02457
```

### 🌍 乌兹项目
```
格式: 介质代码-管道号-管径-管道等级-保温类型
示例: PA-2001002A-100-C1C-N
输出: PA-2001002A
```

### 🏢 天华项目
```
格式: 装置号+介质代码-管道号-管径-管道等级-保温类型
示例: 01PL-216061-125-C22S-H
输出: 01PL-216061
```

### 🔶 金昱元项目
```
格式: 装置号+介质代码+管道号-管径-管道等级-保温类型
示例: 01PC03012-100-C12S-H
输出: 01PC03012
注意: 第一段将装置号、介质代码、管道号合并，无内部连字符
```

## 🛠️ 项目结构

```
CAD2EXL/
├── extractor_core.py       # 核心逻辑（项目类型、正则、解析、Excel输出）
├── pid_extractor.py        # CLI 入口
├── pid_extractor_gui.py    # GUI 入口
├── pid_extractor.spec      # PyInstaller 打包配置
├── requirements.txt
└── fig/
    └── logo.jpg
```

核心逻辑集中在 `extractor_core.py`，**新增项目类型只需修改这一个文件**。

## 🔨 构建可执行文件

```bash
pip install pyinstaller
pyinstaller pid_extractor.spec
# 输出: dist/PID_Extractor.exe
```

## 📦 依赖

| 库 | 用途 |
|----|------|
| pyautocad | AutoCAD COM 接口，读取 DWG |
| pandas | 数据处理 |
| openpyxl | Excel 读写 |
| Pillow | Logo 图像处理 |
| tkinterdnd2 | GUI 拖拽功能 |

## 🔄 版本历史

查看 [CHANGELOG.md](CHANGELOG.md) 了解详细更新记录。

---

⚡ 由 [Claude Code](https://claude.ai/code) 协助开发
