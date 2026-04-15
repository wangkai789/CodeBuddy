# 南乐HQMS上报程序
# 执行存储过程提取病案首页数据，追加到HQMS上报模板CSV文件

import pyodbc
import pandas as pd
import os
import shutil
import argparse
from datetime import datetime, timedelta
import logging

# ============== 配置区域 ==============
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

# 模板文件路径（使用相对路径，相对于程序所在目录）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, '模板', 'hqmsts01.CSV')

# 输出目录
OUTPUT_DIR = os.path.join(BASE_DIR, '输出')


# 日志配置
LOG_DIR = r'D:\codebuddy\logs'
LOG_FILE = os.path.join(LOG_DIR, 'hqms_log.txt')

# =====================================


def get_last_month_range():
    """获取上个月的日期范围：开始日期为上月1号，结束日期为上月最后一天"""
    today = datetime.now()
    # 上个月
    if today.month == 1:
        last_month = datetime(today.year - 1, 12, 1)
    else:
        last_month = datetime(today.year, today.month - 1, 1)
    
    # 上个月最后一天
    # 本月1号减1天
    first_day_this_month = datetime(today.year, today.month, 1)
    last_day_last_month = first_day_this_month - timedelta(days=1)
    
    start_date = last_month.strftime('%Y-%m-%d')
    end_date = last_day_last_month.strftime('%Y-%m-%d')
    
    return start_date, end_date


def setup_logging():
    """配置日志"""
    os.makedirs(LOG_DIR, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILE, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)


def get_db_connection(config):
    """建立数据库连接"""
    try:
        conn_str = (
            f"DRIVER={{{config['driver']}}};"
            f"SERVER={config['server']};"
            f"DATABASE={config['database']};"
            f"UID={config['username']};"
            f"PWD={config['password']}"
        )
        conn = pyodbc.connect(conn_str)
        logging.info("数据库连接成功")
        return conn
    except pyodbc.Error as e:
        logging.error(f"数据库连接失败: {e}")
        raise


def execute_stored_procedure(conn, start_date, end_date):
    """执行存储过程并获取结果"""
    params = f"'{start_date}', '{end_date}'"

    try:
        logging.info(f"执行存储过程: {STORED_PROCEDURE} {params}")

        cursor = conn.cursor()

        # 执行存储过程
        cursor.execute(f"EXEC {STORED_PROCEDURE} {params}")

        # 尝试获取结果集
        all_rows = []
        all_columns = []

        # 处理多个结果集
        while True:
            try:
                if cursor.description:
                    columns = [column[0] for column in cursor.description]
                    rows = cursor.fetchall()
                    if rows:
                        all_rows.extend(rows)
                        if not all_columns:
                            all_columns = columns
                        logging.info(f"获取到 {len(rows)} 条记录")

                # 尝试获取下一个结果集
                if not cursor.nextset():
                    break

            except pyodbc.ProgrammingError:
                break

        if all_rows and all_columns:
            df = pd.DataFrame.from_records(all_rows, columns=all_columns)
            logging.info(f"成功提取 {len(df)} 条记录")
        else:
            df = pd.DataFrame()
            logging.warning("存储过程无返回结果")

        cursor.close()
        return df

    except pyodbc.Error as e:
        logging.error(f"执行存储过程失败: {e}")
        raise


def copy_template_to_output(template_file, output_dir):
    """复制模板文件到输出目录，生成带时间戳的新文件名"""
    os.makedirs(output_dir, exist_ok=True)

    # 生成带时间戳的文件名
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"hqmsts01_{timestamp}.CSV"
    output_file = os.path.join(output_dir, filename)

    # 复制模板文件
    shutil.copy2(template_file, output_file)
    logging.info(f"已复制模板到: {output_file}")

    return output_file


def append_to_template(df, output_file):
    """将数据追加到输出文件"""
    if df.empty:
        logging.warning("没有数据需要追加")
        return False, None

    # 读取模板文件的列名（第一行）
    template_df = pd.read_csv(output_file, encoding='utf-8', nrows=0)
    template_columns = list(template_df.columns)

    logging.info(f"模板列数: {len(template_columns)}")
    logging.info(f"数据列数: {len(df.columns)}")

    # 检查数据列是否与模板列匹配
    if list(df.columns) != template_columns:
        logging.warning("警告：数据列与模板列不完全匹配，将尝试按顺序对应")

    # 确保数据列顺序与模板一致
    # 只保留模板中存在的列
    valid_columns = [col for col in template_columns if col in df.columns]
    df = df[valid_columns]

    # 对于模板中有但数据中没有的列，填充空值
    for col in template_columns:
        if col not in df.columns:
            df[col] = ''

    # 按模板列顺序重新排列
    df = df[template_columns]

    # 转换为文本型（字符串）
    df = df.astype(str)

    # 替换 'nan' 和 'None' 为空字符串
    df = df.replace('nan', '')
    df = df.replace('None', '')

    # 追加到输出文件（不写入列名，只追加数据行）
    with open(output_file, 'a', encoding='utf-8', newline='') as f:
        df.to_csv(f, header=False, index=False)

    logging.info(f"数据已追加到: {output_file}")
    return True, output_file


def main():
    """主函数"""
    logger = setup_logging()

    # 解析命令行参数
    parser = argparse.ArgumentParser(description='南乐HQMS上报程序')
    parser.add_argument('--start-date', type=str, help='开始日期 (YYYY-MM-DD)')
    parser.add_argument('--end-date', type=str, help='结束日期 (YYYY-MM-DD)')
    parser.add_argument('--output', type=str, help='输出目录')

    args = parser.parse_args()

    # 确定日期范围（默认上个月）
    if args.start_date and args.end_date:
        start_date = args.start_date
        end_date = args.end_date
    else:
        start_date, end_date = get_last_month_range()

    # 确定输出目录
    output_dir = args.output if args.output else OUTPUT_DIR

    logger.info("=" * 50)
    logger.info("南乐HQMS上报程序启动")
    logger.info(f"存储过程: {STORED_PROCEDURE}")
    logger.info(f"日期范围: {start_date} 至 {end_date}")
    logger.info(f"模板文件: {TEMPLATE_FILE}")
    logger.info(f"输出目录: {output_dir}")
    logger.info("=" * 50)

    try:
        # 检查模板文件是否存在
        if not os.path.exists(TEMPLATE_FILE):
            logger.error(f"模板文件不存在: {TEMPLATE_FILE}")
            return

        # 复制模板文件到输出目录
        output_file = copy_template_to_output(TEMPLATE_FILE, output_dir)

        # 连接数据库
        conn = get_db_connection(DB_CONFIG)

        # 执行存储过程获取数据
        df = execute_stored_procedure(conn, start_date, end_date)

        if df.empty:
            logger.warning("存储过程无返回结果")
            conn.close()
            return

        # 显示数据列
        logger.info(f"数据列: {list(df.columns)[:20]}...")  # 显示前20列

        # 追加数据到输出文件
        success, result_file = append_to_template(df, output_file)

        if success:
            logger.info("=" * 50)
            logger.info("执行完成!")
            logger.info(f"追加记录数: {len(df)}")
            logger.info(f"输出文件: {result_file}")
            logger.info("=" * 50)

        conn.close()

    except Exception as e:
        logger.error(f"程序执行失败: {e}")
        raise


if __name__ == '__main__':
    main()
