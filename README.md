# 打卡数据处理脚本

本项目包含一个 Python 脚本 `attend.py`，一个配套的 `.yml` 环境文件，该脚本将 `.xlsx` 文件中的打卡次数提取并更新到相应文件的指定位置。

## 项目描述

此脚本的主要功能包括：
1. **读取数据**：从第一个总统计 `.xlsx` 文件中提取姓名。
2. **匹配打卡天数**：在第二个当日导出的打卡 `.xlsx` 文件中搜索姓名，并获取该姓名的总打卡天数。
3. **更新文件**：将找到的打卡天数写入第一个文件的指定列。

## 使用说明

### 1. 下载项目
使用以下命令克隆此项目：

```bash
git clone https://github.com/karinadegou/attendance_statistics.git
cd attendance_statistics
```

### 2. 安装环境
本项目依赖于多个 Python 库，建议使用 conda 创建与 punch.yml 文件一致的环境：

```bash
conda env create -f punch.yml
```

安装完成后，激活新创建的环境：

```bash
conda activate punch
```

### 3. 准备数据文件
请将待处理的 .xlsx 文件放入项目目录，并确保文件的命名和数据列结构符合以下要求：

第一个文件：包含姓名的列在 A 列。
第二个文件：包含姓名的列在 A 列和 C 列，总打卡天数在 F 列。

### 4. 运行脚本
在激活的环境中运行以下命令执行 attend.py 脚本：

```bash
python attend.py
```

### 5. 输出文件
脚本运行完成后，将生成一个新的 `.xlsx` 文件 `file1_updated.xlsx`，其中更新了总统计文件的打卡数据。

## 注意事项
确保 .xlsx 文件的列结构与脚本中的读取方式一致。
第二个文件的姓名列可能包含其他信息，脚本使用正则表达式模糊匹配。
