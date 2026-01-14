# OECT Excel Processor

用于将 OECT（有机电化学晶体管）性能测试后的 Excel 数据（如 LabExpress 导出格式）转换为 CSV 格式的 Python 包。

## 功能特点

- **图形界面** - 提供现代深色主题的 GUI 应用，支持单文件和批量处理
- **命令行工具** - 灵活的 CLI 接口，支持脚本化操作
- **Python API** - 可编程接口，方便集成到数据处理流程
- **类型序列循环** - 工作表类型序列会自动循环应用到所有工作表
- **多进程支持** - 批量处理支持多核并行，提升处理速度
- **自然排序** - 文件按自然顺序排列（file1, file2, file10 而非 file1, file10, file2）
- **数据清洗** - 自动去除空行和不完整数据行
- **独立打包** - 可打包为独立 exe 文件运行

## 安装

### 从 PyPI 安装

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

### 图形界面

启动 GUI 应用：

```bash
oect-gui
```

或运行独立可执行文件 `OECT-Excel-Processor.exe`

#### 单文件模式

选择单个 Excel 文件进行处理：

![单文件模式](assets/single_file_mode.png)

#### 批量处理模式

选择包含多个 Excel 文件的文件夹进行批量处理：

![批量处理模式](assets/batch_mode.png)

#### GUI 操作步骤

1. 选择处理模式（单文件/批量处理）
2. 选择文件或文件夹
3. 配置类型序列（如 `transfer,transient`）
4. 设置输出前缀（可选）
5. 点击「开始处理」按钮

### 命令行工具

#### 单文件处理

```bash
oect-processor single <文件路径> [选项]
```

选项：
| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-t, --sheet-types` | 工作表类型序列，逗号分隔 | `transfer,transient` |
| `-o, --output-prefix` | 输出 CSV 文件前缀 | `output` |

示例：

```bash
oect-processor single data.xls -t transfer,transient -o result
```

#### 批量处理

```bash
oect-processor batch <目录路径> [选项]
```

选项：
| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-t, --sheet-types` | 工作表类型序列，逗号分隔 | `transfer,transient` |
| `-o, --output-prefix` | 输出 CSV 文件前缀 | `batch_output` |
| `-p, --pattern` | 文件匹配模式 | `*.xls` |
| `-d, --output-dir` | 输出目录 | 当前目录 |
| `-m, --multiprocessing` | 启用多进程处理 | 否 |
| `-w, --workers` | 最大工作进程数 | CPU 核心数 |

示例：

```bash
# 基本批量处理
oect-processor batch ./data_folder -t transfer,transient

# 使用多进程加速
oect-processor batch ./data_folder -m -w 4

# 指定输出目录和文件模式
oect-processor batch ./data_folder -p "*.xlsx" -d ./output -m
```

### Python API

#### 单文件处理

```python
from oect_excel_processor import ExcelProcessor

# 创建处理器
processor = ExcelProcessor(
    file_path="data.xls",
    sheet_types=["transfer", "transient"],
    output_prefix="output"
)

# 处理并保存
saved_files = processor.process_and_save()
print(f"生成的文件: {saved_files}")

# 获取工作表信息
sheet_info = processor.get_sheet_info()
print(f"工作表类型映射: {sheet_info}")
```

#### 批量处理

```python
from oect_excel_processor import BatchExcelProcessor

# 创建批处理器
batch = BatchExcelProcessor(
    directory="./data_folder",
    file_pattern="*.xls",
    sheet_types=["transfer", "transient"],
    output_prefix="batch_output"
)

# 获取待处理文件列表
excel_files = batch.get_excel_files()
print(f"找到 {len(excel_files)} 个文件")

# 处理所有文件（单进程）
results = batch.process_all_files(output_dir="./output")

# 或使用多进程加速
results = batch.process_all_files(
    output_dir="./output",
    use_multiprocessing=True,
    max_workers=4
)

# 获取处理摘要
summary = batch.get_processing_summary(results)
print(f"成功: {summary['successful_files']}, 失败: {summary['failed_files']}")
print(f"生成 CSV 文件数: {summary['total_csv_files']}")
```

## 工作表类型

### transfer 类型

- **数据格式**：第 3 行为字段名，第 4 行开始为数据
- **列数**：4 列
- **适用场景**：转移特性测试数据

### transient 类型

- **数据格式**：第 3 行前两列为字段名，第 4 行开始为数据
- **列数**：每两列一组，自动合并所有组
- **适用场景**：瞬态测试数据

## 类型序列循环

类型序列会**循环应用**到所有工作表：

| 类型序列 | 4 个工作表的处理结果 |
|---------|-------------------|
| `transfer,transient` | Sheet1=transfer, Sheet2=transient, Sheet3=transfer, Sheet4=transient |
| `transient` | 所有工作表都按 transient 处理 |
| `transfer,transfer,transient` | 按 2:1 比例循环 |

## 输出文件命名

### 单文件处理

```
{前缀}-{序号}-{类型}.csv
```

示例：`output-1-transfer.csv`, `output-2-transient.csv`

### 批量处理

```
{前缀}-{文件序号}-{工作表序号}-{类型}.csv
```

示例：`batch_output-1-1-transfer.csv`, `batch_output-1-2-transient.csv`

## 常见问题

**Q: 支持哪些 Excel 格式？**
A: 支持 `.xls` 和 `.xlsx` 格式。

**Q: 类型序列如何工作？**
A: 类型序列会循环应用。例如设置 `transfer,transient` 处理 4 个工作表，则 Sheet1/3 为 transfer，Sheet2/4 为 transient。

**Q: 如何提高批量处理速度？**
A: 使用 `-m` 参数启用多进程处理，并通过 `-w` 指定工作进程数。

**Q: 输出文件保存在哪里？**
A: 单文件模式默认保存在当前目录；批量模式可通过 `-d` 参数指定输出目录。

## 依赖

- pandas >= 1.0.0
- numpy >= 1.18.0
- natsort >= 7.0.0
- xlrd >= 2.0.1

## 许可证

MIT
