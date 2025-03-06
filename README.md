# Cadence BOM转换工具

## 项目简介
使用Trae AI编写的一个Cadence BOM转换工具，可以直接输入allegro生成的Component Report直接保存的html格式文件或另存为的excel文件，转换的文件为常规使用的Excel格式，包含VALUE值、封装、正反面位号和数量等信息。

以下所有内容包含代码均为AI生成的，仅供学习参考。

Cadence BOM转换工具是一个用于处理和转换Cadence生成的BOM（物料清单）文件的桌面应用程序。该工具支持将HTML和Excel格式的BOM文件转换为标准化的Excel格式，方便后续处理和分析。

## 主要功能

- 支持HTML和Excel格式的BOM文件导入
- 自动识别和处理文件中的关键数据列
- 区分PCB顶层和底层元器件
- 自动统计元器件数量
- 生成标准化的Excel格式输出文件
- 实时显示转换进度
- 友好的图形用户界面

## 安装要求

### 依赖项

- Python 3.6+
- pandas >= 1.0.0
- beautifulsoup4 >= 4.9.0
- lxml >= 4.5.0
- openpyxl >= 3.0.0（用于Excel文件处理）
- tkinter（Python标准库，无需单独安装）

### 安装步骤

#### Windows系统

1. 安装Python：
   - 从[Python官网](https://www.python.org/downloads/)下载并安装Python 3.6或更高版本
   - 安装时勾选"Add Python to PATH"选项
   - 验证安装：打开命令提示符，输入`python --version`确认安装成功

2. 下载项目代码：
   - 使用Git：`git clone [项目仓库URL]`
   - 或直接从项目页面下载ZIP文件并解压

3. 安装依赖项：
   - 打开命令提示符，进入项目目录
   - 执行以下命令安装所需依赖：
     ```bash
     pip install pandas>=1.0.0 beautifulsoup4>=4.9.0 lxml>=4.5.0 openpyxl>=3.0.0
     ```

#### macOS系统

1. 安装Python：
   - 使用Homebrew：`brew install python`
   - 或从[Python官网](https://www.python.org/downloads/)下载安装包
   - 验证安装：打开终端，输入`python3 --version`确认安装成功

2. 下载项目代码：
   - 使用Git：`git clone [项目仓库URL]`
   - 或直接从项目页面下载ZIP文件并解压

3. 安装依赖项：
   - 打开终端，进入项目目录
   - 执行以下命令安装所需依赖：
     ```bash
     pip3 install pandas>=1.0.0 beautifulsoup4>=4.9.0 lxml>=4.5.0 openpyxl>=3.0.0
     ```

#### Linux系统

1. 安装Python：
   - Debian/Ubuntu：`sudo apt-get update && sudo apt-get install python3 python3-pip`
   - CentOS/RHEL：`sudo yum install python3 python3-pip`
   - 验证安装：终端输入`python3 --version`确认安装成功

2. 下载项目代码：
   - 使用Git：`git clone [项目仓库URL]`
   - 或直接从项目页面下载ZIP文件并解压

3. 安装依赖项：
   - 打开终端，进入项目目录
   - 执行以下命令安装所需依赖：
     ```bash
     pip3 install pandas>=1.0.0 beautifulsoup4>=4.9.0 lxml>=4.5.0 openpyxl>=3.0.0
     ```

### 可能遇到的问题及解决方案

1. **pip命令未找到**
   - Windows：确保Python已添加到PATH中，或使用`py -m pip`代替`pip`
   - macOS/Linux：尝试使用`pip3`代替`pip`

2. **lxml安装失败**
   - Windows：尝试安装预编译的wheel文件：`pip install wheel`然后重新安装lxml
   - Linux：安装编译依赖：`sudo apt-get install libxml2-dev libxslt-dev python3-dev`(Ubuntu)或`sudo yum install libxml2-devel libxslt-devel python3-devel`(CentOS)

3. **tkinter缺失**
   - Windows：重新安装Python，确保勾选了tcl/tk选项
   - Ubuntu：`sudo apt-get install python3-tk`
   - CentOS：`sudo yum install python3-tkinter`

### 验证安装

安装完成后，可以运行以下命令验证依赖项是否正确安装：

```bash
python -c "import pandas; import bs4; import lxml; import tkinter; print('所有依赖项已成功安装！')"
```

## 使用说明

1. 运行程序：
   ```bash
   python main.py
   ```

2. 在程序界面中点击"打开文件"按钮，选择需要转换的BOM文件（支持.html、.htm、.xlsx、.xls格式）

3. 程序会自动开始转换，并显示转换进度

4. 转换完成后，选择保存路径，程序会将转换后的文件保存为Excel格式

## 输入文件要求

输入的BOM文件必须包含以下列：
- REFDES（器件位号）
- COMP_VALUE（器件值）
- COMP_PACKAGE（封装类型）
- SYM_MIRROR（器件位置，顶层/底层）

## 输出文件格式

转换后的Excel文件包含以下列：
- Item（序号）
- COMP_VALUE（器件值）
- COMP_PACKAGE（封装类型）
- Top REFDES（顶层器件位号）
- Bottom REFDES（底层器件位号）
- Quantity（器件总数量）

## 版本信息

当前版本：1.0.3
发布日期：2025-02-28

## 许可证

本项目采用MIT许可证。