# OECT Excel Processor 用户手册

将 LabExpress 导出的 OECT 测试数据（Excel）转换为 CSV 格式。

---

## 快速开始

### 1. 启动程序

双击运行 `OECT-Excel-Processor.exe`

### 2. 选择处理模式

| 模式 | 用途 |
|------|------|
| **单文件** | 处理单个 Excel 文件 |
| **批量处理** | 处理整个文件夹中的所有 Excel 文件 |

### 3. 选择文件/文件夹

点击 **「选择文件」** 或 **「选择文件夹」** 按钮

### 4. 配置类型序列

在**「类型序列」**输入框中设置工作表类型，以逗号分隔：

| 输入 | 效果 |
|------|------|
| `transfer,transient` | Sheet1=transfer, Sheet2=transient, Sheet3=transfer... (循环) |
| `transient` | 所有工作表都按 transient 处理 |
| `transfer,transfer,transient` | 2:1 比例循环 |

### 5. 点击开始处理

点击 **「⚡ 开始处理」** 按钮，等待处理完成。

---

## 界面说明

### 主界面

![主界面](assets/gui_main.png)

### 处理完成

![处理完成](assets/gui_complete.png)

---

## 输出文件

CSV 文件保存在与源文件相同的目录，命名格式：

```
{前缀}-{序号}-{类型}.csv
```

例如：`processed_-1-transfer.csv`

---

## 命令行使用

```bash
# 单文件处理
oect-processor single data.xls -t transfer,transient

# 批量处理
oect-processor batch ./data_folder -t transfer,transient
```

---

## 常见问题

**Q: 支持哪些 Excel 格式？**  
A: `.xls` 和 `.xlsx`

**Q: 类型序列如何工作？**  
A: 类型序列会循环应用。例如设置 `transfer,transient` 处理 4 个工作表，则 Sheet1/3 为 transfer，Sheet2/4 为 transient。
