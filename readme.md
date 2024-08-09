# Docx Text Scanner

[English](#english) | [中文](#中文)

## English

### Project Overview

**Docx Text Scanner** is a Python script designed to process and analyze `.docx` files. The script standardizes file names, reads document content, counts keyword occurrences, and generates a formatted Excel summary. This tool is particularly useful for users who need to handle and analyze multiple `.docx` files systematically.

### Features

- Standardizes file names by padding numbers to three digits and removing leading symbols and spaces.
- Reads the content of `.docx` files.
- Counts occurrences of user-defined keywords.
- Generates a formatted Excel summary (`output.xlsx`) with keyword statistics.
- Customizable keywords via `config.json`.

### Prerequisites

- Python 3.x
- Pipenv for managing dependencies

### Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/yourusername/docx-text-scanner.git
   cd docx-text-scanner
   ```

2. Install dependencies using Pipenv:
   ```bash
   pipenv install
   ```

### Usage

1. Place your `.docx` files in the `files` directory located in the root of the project.

2. Customize your keywords by editing `config.json`:

   ```json
   {
     "keywords": ["keyword1", "keyword2", "keyword3"]
   }
   ```

3. Run the script:

   ```bash
   pipenv run python run.py
   ```

4. The results will be saved in `output.xlsx` in the root directory.

### Example

1. **Original File Names**:

   ```
   files/
   ├── 1_document.docx
   ├── 2_report.docx
   └── 3_summary.docx
   ```

2. **Standardized File Names**:

   ```
   files/
   ├── 001-document.docx
   ├── 002-report.docx
   └── 003-summary.docx
   ```

3. **Excel Output**:
   The script generates an Excel file (`output.xlsx`) with keyword statistics and formatted data.

### License

This project is licensed under the MIT License.

---

## 中文

### 项目概述

**Docx Text Scanner** 是一个用于处理和分析 `.docx` 文件的 Python 脚本。该脚本可以标准化文件名、读取文档内容、统计关键词出现次数，并生成格式化的 Excel 汇总表。对于需要系统化处理和分析多个 `.docx` 文件的用户来说，这个工具非常有用。

### 功能

- 通过将数字填充为三位数并去掉开头的符号和空格来标准化文件名。
- 读取 `.docx` 文件的内容。
- 统计用户定义的关键词出现次数。
- 生成包含关键词统计信息的格式化 Excel 汇总表 (`output.xlsx`)。
- 可以通过 `config.json` 自定义关键词。

### 前提条件

- Python 3.x
- 使用 Pipenv 管理依赖

### 安装

1. 克隆仓库：

   ```bash
   git clone https://github.com/yourusername/docx-text-scanner.git
   cd docx-text-scanner
   ```

2. 使用 Pipenv 安装依赖：
   ```bash
   pipenv install
   ```

### 使用方法

1. 将你的 `.docx` 文件放入项目根目录下的 `files` 文件夹中。

2. 通过编辑 `config.json` 自定义关键词：

   ```json
   {
     "keywords": ["关键词1", "关键词2", "关键词3"]
   }
   ```

3. 运行脚本：

   ```bash
   pipenv run python run.py
   ```

4. 结果将保存在根目录下的 `output.xlsx` 文件中。

### 示例

1. **原始文件名**：

   ```
   files/
   ├── 1_document.docx
   ├── 2_report.docx
   └── 3_summary.docx
   ```

2. **标准化文件名**：

   ```
   files/
   ├── 001-document.docx
   ├── 002-report.docx
   └── 003-summary.docx
   ```

3. **Excel 输出**：
   脚本生成一个包含关键词统计信息和格式化数据的 Excel 文件 (`output.xlsx`)。

### 许可证

本项目使用 MIT 许可证。
