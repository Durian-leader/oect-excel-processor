# OECT Excel Processor

用于处理OECT（有机电化学晶体管）性能测试后的Excel数据并转换为CSV格式的Python包。

## 功能特点

- ✨ **图形界面** - 提供易用的GUI应用，支持单文件和批量处理
- 📊 支持两种工作表类型：`transfer` 和 `transient`
- 🔄 **类型序列循环** - 工作表类型序列会自动循环应用到所有工作表
- 📁 支持批量处理多个Excel文件
- 🧹 自动去除空行和不完整数据行
- 📦 可打包为独立exe运行

## 安装

### 从PyPI安装

```bash
pip install oect-excel-processor
```

### 从源码安装

```bash
git clone https://github.com/Durian-leader/oect-excel-processor.git
cd oect-excel-processor
pip install -e .
```

## 使用方法

### 图形界面 (推荐)

启动GUI应用：

```bash
oect-gui
```

或直接运行 `OECT-Excel-Processor.exe`

![单文件模式](assets/single_file_mode.png)

详细使用说明请参阅 [用户手册](USER_MANUAL.md)

### 命令行工具

```bash
# 单文件处理
oect-processor single data.xls -t transfer,transient

# 批量处理
oect-processor batch ./data_folder -t transfer,transient
```

### Python API

```python
from oect_excel_processor import ExcelProcessor, BatchExcelProcessor

# 单文件处理
processor = ExcelProcessor("data.xls", ["transfer", "transient"], "output")
saved_files = processor.process_and_save()

# 批量处理
batch = BatchExcelProcessor("./data_folder", sheet_types=["transfer", "transient"])
results = batch.process_all_files()
```

## 类型序列说明

类型序列会**循环应用**到所有工作表：

| 类型序列 | 4个工作表的处理结果 |
|---------|-------------------|
| `transfer,transient` | Sheet1=transfer, Sheet2=transient, Sheet3=transfer, Sheet4=transient |
| `transient` | 全部按transient处理 |
| `transfer,transfer,transient` | 2:1比例循环 |

## 工作表类型

- **transfer**: 从第三行开始，共四列数据
- **transient**: 数据按每两列一组排列，自动合并

## 输出文件

```
{前缀}-{序号}-{类型}.csv
```

例如：`processed_-1-transfer.csv`, `processed_-2-transient.csv`

## 许可证

MIT