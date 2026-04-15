# 南乐HQMS上报程序 - Web界面
# 基于Flask的Web UI，用于配置参数和运行程序

import pyodbc
import pandas as pd
import os
import shutil
import json
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash
from threading import Thread

app = Flask(__name__)
app.secret_key = 'hqms_secret_key_2024'

# ============== 配置区域 ==============
# 获取程序所在目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 默认配置
DEFAULT_CONFIG = {
    'server': '127.0.0.1',
    'database': 'yiyuandb',
    'username': 'founder',
    'password': 'fd',
    'driver': 'ODBC Driver 17 for SQL Server'
}

# 模板文件路径（使用相对路径）
TEMPLATE_FILE = os.path.join(BASE_DIR, '模板', 'hqmsts01.CSV')

# 输出目录
OUTPUT_DIR = os.path.join(BASE_DIR, '输出')

# 配置文件路径
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')

# 日志目录
LOG_DIR = os.path.join(BASE_DIR, 'logs')
# =====================================


def load_config():
    """加载配置文件"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return DEFAULT_CONFIG.copy()


def save_config(config):
    """保存配置文件"""
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)


def get_db_connection(config):
    """建立数据库连接"""
    conn_str = (
        f"DRIVER={{{config['driver']}}};"
        f"SERVER={config['server']};"
        f"DATABASE={config['database']};"
        f"UID={config['username']};"
        f"PWD={config['password']}"
    )
    return pyodbc.connect(conn_str)


def execute_stored_procedure(conn, start_date, end_date, stored_procedure):
    """执行存储过程并获取结果"""
    params = f"'{start_date}', '{end_date}'"
    cursor = conn.cursor()
    cursor.execute(f"EXEC {stored_procedure} {params}")

    all_rows = []
    all_columns = []

    while True:
        try:
            if cursor.description:
                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()
                if rows:
                    all_rows.extend(rows)
                    if not all_columns:
                        all_columns = columns
            if not cursor.nextset():
                break
        except pyodbc.ProgrammingError:
            break

    cursor.close()

    if all_rows and all_columns:
        return pd.DataFrame.from_records(all_rows, columns=all_columns)
    return pd.DataFrame()


def copy_template_to_output(template_file, output_dir):
    """复制模板文件到输出目录"""
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"hqmsts01_{timestamp}.CSV"
    output_file = os.path.join(output_dir, filename)
    shutil.copy2(template_file, output_file)
    return output_file


def append_to_template(df, output_file):
    """将数据追加到输出文件"""
    if df.empty:
        return False

    template_df = pd.read_csv(output_file, encoding='utf-8', nrows=0)
    template_columns = list(template_df.columns)

    valid_columns = [col for col in template_columns if col in df.columns]
    df = df[valid_columns]

    for col in template_columns:
        if col not in df.columns:
            df[col] = ''

    df = df[template_columns]
    df = df.astype(str)
    df = df.replace('nan', '')
    df = df.replace('None', '')

    with open(output_file, 'a', encoding='utf-8', newline='') as f:
        df.to_csv(f, header=False, index=False)

    return True


def run_hqms_task(config, start_date, end_date, stored_procedure, result_holder):
    """后台运行HQMS任务"""
    try:
        log_file = os.path.join(LOG_DIR, 'hqms_web_log.txt')
        os.makedirs(LOG_DIR, exist_ok=True)

        with open(log_file, 'a', encoding='utf-8') as log:
            log.write(f"\n{'='*50}\n")
            log.write(f"开始时间: {datetime.now()}\n")
            log.write(f"日期范围: {start_date} 至 {end_date}\n")
            log.write(f"{'='*50}\n")

            # 连接数据库
            conn = get_db_connection(config)
            log.write("数据库连接成功\n")

            # 执行存储过程
            df = execute_stored_procedure(conn, start_date, end_date, stored_procedure)
            log.write(f"获取到 {len(df)} 条记录\n")

            if df.empty:
                result_holder['status'] = 'warning'
                result_holder['message'] = '存储过程无返回结果'
                conn.close()
                return

            # 复制模板
            output_file = copy_template_to_output(TEMPLATE_FILE, OUTPUT_DIR)
            log.write(f"已复制模板到: {output_file}\n")

            # 追加数据
            success = append_to_template(df, output_file)

            if success:
                result_holder['status'] = 'success'
                result_holder['message'] = f'成功！共 {len(df)} 条记录'
                result_holder['file'] = output_file
                log.write(f"完成！输出文件: {output_file}\n")
            else:
                result_holder['status'] = 'error'
                result_holder['message'] = '数据追加失败'
                log.write("数据追加失败\n")

            conn.close()

    except Exception as e:
        result_holder['status'] = 'error'
        result_holder['message'] = f'执行失败: {str(e)}'
        with open(log_file, 'a', encoding='utf-8') as log:
            log.write(f"错误: {str(e)}\n")


@app.route('/')
def index():
    """首页"""
    config = load_config()
    return render_template('index.html', config=config)


@app.route('/config', methods=['GET', 'POST'])
def config_page():
    """配置页面"""
    config = load_config()

    if request.method == 'POST':
        config['server'] = request.form.get('server', '')
        config['database'] = request.form.get('database', '')
        config['username'] = request.form.get('username', '')
        config['password'] = request.form.get('password', '')
        config['driver'] = request.form.get('driver', 'ODBC Driver 17 for SQL Server')
        save_config(config)
        flash('配置已保存！', 'success')
        return redirect(url_for('index'))

    return render_template('config.html', config=config)


@app.route('/test_connection', methods=['POST'])
def test_connection():
    """测试数据库连接"""
    config = {
        'server': request.form.get('server', ''),
        'database': request.form.get('database', ''),
        'username': request.form.get('username', ''),
        'password': request.form.get('password', ''),
        'driver': request.form.get('driver', 'ODBC Driver 17 for SQL Server')
    }

    try:
        conn = get_db_connection(config)
        conn.close()
        return {'status': 'success', 'message': '数据库连接成功！'}
    except Exception as e:
        return {'status': 'error', 'message': f'连接失败: {str(e)}'}


@app.route('/run', methods=['POST'])
def run_task():
    """运行HQMS任务"""
    config = load_config()

    start_date = request.form.get('start_date', '')
    end_date = request.form.get('end_date', '')
    stored_procedure = request.form.get('stored_procedure', 'yiyuandb..FirstPage_exoprt')

    if not start_date or not end_date:
        flash('请填写开始日期和结束日期！', 'error')
        return redirect(url_for('index'))

    # 检查模板文件
    if not os.path.exists(TEMPLATE_FILE):
        flash(f'模板文件不存在: {TEMPLATE_FILE}', 'error')
        return redirect(url_for('index'))

    # 启动后台任务
    result_holder = {}
    thread = Thread(target=run_hqms_task, args=(config, start_date, end_date, stored_procedure, result_holder))
    thread.start()

    # 等待结果
    thread.join(timeout=60)

    if result_holder.get('status') == 'success':
        flash(f"执行成功！{result_holder['message']}<br>输出文件: {result_holder['file']}", 'success')
    elif result_holder.get('status') == 'warning':
        flash(result_holder['message'], 'warning')
    else:
        flash(result_holder.get('message', '执行失败'), 'error')

    return redirect(url_for('index'))


@app.route('/open-output-dir')
def open_output_dir():
    """打开输出目录"""
    try:
        output_dir = r'D:\codebuddy\输出'
        if os.path.exists(output_dir):
            import subprocess
            subprocess.Popen(f'explorer "{output_dir}"')
            return jsonify({'success': True, 'path': output_dir})
        else:
            return jsonify({'success': False, 'message': '目录不存在'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


if __name__ == '__main__':
    # 创建templates目录
    os.makedirs('templates', exist_ok=True)

    print("=" * 50)
    print("南乐HQMS上报系统 - Web界面")
    print("=" * 50)
    print("访问地址: http://127.0.0.1:5001")
    print("配置页面: http://127.0.0.1:5001/config")
    print("=" * 50)

    app.run(host='0.0.0.0', port=5001, debug=True)
