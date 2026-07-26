# 费用报销 Excel 自动生成工具

拖拉发票文件夹，自动识别费用明细，一键生成标准格式的 Excel 报销表。

## 功能

- 📂 **文件夹导入**：拖拽或选择发票文件夹，自动识别子文件夹（高速通行费）和根目录 PDF（加油费）
- ✏️ **手动编辑**：表格直接编辑、增删行，灵活调整
- 📋 **文本粘贴**：粘贴文件夹结构文本，一键解析
- 📥 **下载 Excel**：生成与参考模板样式一致的 Excel 文件（微软雅黑、细线边框、合计行）
- 🤖 **智能分类**：自动识别通行费/加油费 → 汽车费用，默认部门「产品质量」

## 使用方式

### 网页版（推荐）

直接在浏览器打开 `index.html`，支持 Chrome / Edge。

### 本地 Python 脚本

```bash
pip install openpyxl
python generate_expense_report.py "发票文件夹路径"
```

不传参数会自动扫描桌面默认文件夹。

## 项目结构

```
费用报销工具/
├── index.html                    # 网页版应用
├── generate_expense_report.py    # 本地 Python 脚本
├── .gitignore
└── README.md
```
