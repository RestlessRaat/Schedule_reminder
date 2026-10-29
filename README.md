# Schedule Reminder · Project 任务状态分析

读取 Microsoft Project 导出的 Excel（`.xlsx`）计划表，提取叶子任务，根据计划日期分析任务状态，并生成 JSON 结果。支持文件选择窗口和任务结果表格，也可通过命令行运行。

## 环境要求

- Python 3.10 或更高版本。
- 核心分析仅使用 Python 标准库，无需安装第三方依赖。
- 图形界面需要 Tkinter；Windows 官方 Python 安装程序通常提供该组件。

## 快速开始

下载或克隆仓库后，在项目目录运行：

```powershell
python extract_project_task_status.py
```

在弹出的文件选择窗口中选择从 Project 导出的 Excel 文件。分析完成后显示任务表格，支持状态筛选和复制任务路径。

也可以指定输入文件并关闭结果窗口：

```powershell
python extract_project_task_status.py "Data/plan.xlsx" --no-popup
```

指定分析日期及输出位置：

```powershell
python extract_project_task_status.py "Data/plan.xlsx" --date 2026-10-30 --output "outputs/result.json" --no-popup
```

默认输出到脚本所在目录的 `outputs/<输入文件名>_task_status.json`。`Data` 和 `outputs` 文件夹中的内容不会提交到 GitHub；请自行准备输入数据。

## Excel 格式

输入表需要包含以下列名：

```text
标识号、活动、任务模式、名称、工期、开始时间、完成时间、前置任务、大纲级别、备注
```

程序优先使用名为 `任务_表` 的工作表，也可通过 `--sheet-name` 指定工作表。使用大纲级别识别任务层级，并提取叶子任务。

## 状态含义

状态按分析日期与计划开始、完成日期计算。它反映计划日期关系，不代表实际完成情况；使用时应结合现场进展判断。

- **滞后**：计划完成日期早于分析日期。
- **进行中**：分析日期位于计划开始与完成日期之间（含边界）。
- **未开始**：计划开始日期晚于分析日期。
- **未知**：缺少可用于判断的日期。

## 命令行参数

| 参数 | 说明 |
| --- | --- |
| `input_xlsx` | 可选的输入文件路径；省略时弹出选择窗口 |
| `--output` | 指定 JSON 输出路径 |
| `--sheet-name` | 指定工作表名称 |
| `--date` | 分析日期，格式 `YYYY-MM-DD`，默认当天 |
| `--include-summary-tasks` | 在 JSON 中同时包含汇总任务记录 |
| `--no-popup` | 关闭分析结果图形窗口，适合命令行批处理 |

## 打包 Windows 程序

可选安装 PyInstaller，然后使用仓库中的配置打包：

```powershell
python -m pip install pyinstaller
python -m PyInstaller ProjectTaskStatus.spec
```

生成的程序位于 `dist/ProjectTaskStatus.exe`。打包产物不纳入源码仓库。
