# 南乐HQMS上报程序需求文档

## 功能概述

从医院SQL Server数据库执行存储过程提取病案首页数据，生成HQMS上报文件，供人工上传到HQMS系统。

支持两种运行方式：
1. **命令行版本**：`南乐hqms上报.py`
2. **Web界面版本**：`南乐hqms上报_web.py`

## 目录结构

```
南乐HQMS/
├── 南乐hqms上报.py           # 命令行版本
├── 南乐hqms上报_web.py       # Web界面版本
├── 南乐hqms上报.md           # 需求文档
├── templates/                # Web页面模板
│   ├── index.html
│   └── config.html
├── 模板/                    # CSV模板文件
│   └── hqmsts01.CSV
├── 输出/                    # 导出文件目录
├── logs/                    # 日志目录
├── config.json              # Web配置存储
├── 打包.bat                 # 打包脚本
└── 启动Web.bat              # 启动Web界面
```

## 数据源

- **数据库类型**：SQL Server
- **数据库名称**：yiyuandb
- **数据来源**：存储过程 `yiyuandb..FirstPage_exoprt`
- **数据格式**：HQMS标准编码字段（A01、A02、B12等，共811列）

## 需求清单

### 1. 数据库连接
- 支持SQL Server数据库连接
- 支持SQL Server认证方式（用户名/密码）
- 支持Windows认证
- Web界面支持在线修改配置
- 连接参数可配置

### 2. 数据提取
- 执行存储过程 `yiyuandb..FirstPage_exoprt` 获取数据
- 存储过程参数：开始日期、结束日期（格式：'YYYY-MM-DD'）
- 支持多结果集获取

### 3. 日期范围
- **默认日期**：上个月1号至月末
- 命令行参数指定日期范围
- Web界面日期选择器设置默认值为上个月

### 4. 文件输出
- 先复制模板文件到输出目录
- 生成带时间戳的文件名：`hqmsts01_YYYYMMDD_HHMMSS.CSV`
- 输出目录：`南乐HQMS/输出/`
- 数据以文本型输出（字符串格式）
- `nan` 和 `None` 替换为空字符串

### 5. 日志记录
- 日志目录：`南乐HQMS/logs/`
- 记录内容：
  - 程序启动/结束信息
  - 数据库连接状态
  - 存储过程执行状态
  - 数据提取数量
  - 输出文件信息

## 运行方式

### 命令行版本

```bash
cd 南乐HQMS

# 基本运行（提取上个月数据）
python 南乐hqms上报.py

# 指定日期范围
python 南乐hqms上报.py --start-date 2026-03-01 --end-date 2026-03-31

# 指定输出目录
python 南乐hqms上报.py --output "D:\自定义目录"
```

### Web界面版本

```bash
cd 南乐HQMS

# 启动Web服务
python 南乐hqms上报_web.py
```

访问地址：http://127.0.0.1:5001

**Web功能：**
- 设置日期范围
- 选择/修改存储过程名称
- 修改数据库连接配置
- 测试数据库连接
- 执行导出
- 打开输出目录

## 配置参数

### 命令行版本

```python
# 数据库配置
DB_CONFIG = {
    'server': '127.0.0.1',
    'database': 'yiyuandb',
    'username': 'founder',
    'password': 'fd',
    'driver': 'ODBC Driver 17 for SQL Server'
}

# 存储过程名称
STORED_PROCEDURE = 'yiyuandb..FirstPage_exoprt'

# 路径（使用相对路径）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, '模板', 'hqmsts01.CSV')
OUTPUT_DIR = os.path.join(BASE_DIR, '输出')
```

### Web界面版本

配置通过Web页面在线修改，存储在 `config.json` 文件中。

## 数据字段说明

存储过程返回HQMS标准编码字段，部分字段示例：

| 字段编码 | 说明 |
|----------|------|
| A01 | 病案号 |
| A02 | 患者姓名 |
| A48 | 身份证号 |
| A49 | 联系电话 |
| B12 | 入院日期 |
| B15 | 出院日期 |
| A47 | 住院天数 |
| ... | 共811列 |

## 程序输出示例

```
2026-04-14 18:48:21 - INFO - ==================================================
2026-04-14 18:48:21 - INFO - 南乐HQMS上报程序启动
2026-04-14 18:48:21 - INFO - 存储过程: yiyuandb..FirstPage_exoprt
2026-04-14 18:48:21 - INFO - 日期范围: 2026-03-01 至 2026-03-31
2026-04-14 18:48:21 - INFO - 模板文件: D:\codebuddy\南乐HQMS\模板\hqmsts01.CSV
2026-04-14 18:48:21 - INFO - 输出目录: D:\codebuddy\南乐HQMS\输出
2026-04-14 18:48:21 - INFO - ==================================================
2026-04-14 18:48:21 - INFO - 已复制模板到: D:\codebuddy\南乐HQMS\输出\hqmsts01_20260414_184821.CSV
2026-04-14 18:48:21 - INFO - 数据库连接成功
2026-04-14 18:48:21 - INFO - 执行存储过程: yiyuandb..FirstPage_exoprt '2026-03-01', '2026-03-31'
2026-04-14 18:48:23 - INFO - 成功提取 12 条记录
2026-04-14 18:48:23 - INFO - 数据已追加到: D:\codebuddy\南乐HQMS\输出\hqmsts01_20260414_184821.CSV
2026-04-14 18:48:23 - INFO - ==================================================
2026-04-14 18:48:23 - INFO - 执行完成!
2026-04-14 18:48:23 - INFO - 追加记录数: 12
2026-04-14 18:48:23 - INFO - 输出文件: D:\codebuddy\南乐HQMS\输出\hqmsts01_20260414_184821.CSV
2026-04-14 18:48:23 - INFO - ==================================================
```

## 依赖库

- pyodbc：SQL Server数据库连接
- pandas：数据处理
- flask：Web界面框架
- openpyxl：Excel文件处理（如需要）

安装命令：
```bash
pip install pyodbc pandas flask
```

## 注意事项

1. 确保SQL Server服务正常运行
2. 确保ODBC Driver for SQL Server已安装
3. 模板文件必须预先存在且包含正确的标题行
4. 每次运行生成新的时间戳文件，不会覆盖历史数据
5. Web界面配置保存在 `config.json` 中
6. 上报前请人工检查数据完整性
7. 其他电脑运行可能需要安装ODBC Driver
