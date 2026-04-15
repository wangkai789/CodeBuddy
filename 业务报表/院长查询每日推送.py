import pyodbc
import requests
import sqlite3
import sys
import os
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from io import BytesIO

# ============== IRIS ODBC配置 ==============
ODBC_DSN = "IRIS_DHC"
ODBC_USER = "dhsuper"
ODBC_PASSWORD = "o8kfmm#2P"

# ============== 钉钉配置 ==============
DINGTALK_WEBHOOK = "https://oapi.dingtalk.com/robot/send?access_token=d68e1626c9e5689513089f2d67a9b698067442884f342c733b82e953abb71df7"
DINGTALK_CORP_ID = "ding5wrt0wqdqrwzpxpa"
DINGTALK_CORP_SECRET = "1QlNVwS75zFkIcscCeBmpC7mWg0d_4iAzK5hBWYGd22KOKaSGT7lf-oo2ikFH1Jx"

# ============== SQLite配置 ==============
DB_PATH = r'D:\文档\正骨医院\院长查询\data\daily_report.db'

# ============== 连接IRIS数据库 ==============
def connect_iris():
    try:
        connection = pyodbc.connect(
            f'DSN={ODBC_DSN};UID={ODBC_USER};PWD={ODBC_PASSWORD}'
        )
        print(f"成功连接到 IRIS ODBC: {ODBC_DSN}")
        return connection
    except Exception as e:
        print(f"连接失败: {e}")
        return None

# ============== 连接SQLite数据库 ==============
def connect_sqlite():
    try:
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        conn = sqlite3.connect(DB_PATH)
        # 创建表（如果不存在）
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS daily_report (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                query_date TEXT NOT NULL,
                mz_count INTEGER NOT NULL,
                jz_count INTEGER NOT NULL,
                ry_count INTEGER NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(query_date)
            )
        ''')
        conn.commit()
        print(f"成功连接到 SQLite: {DB_PATH}")
        return conn
    except Exception as e:
        print(f"SQLite连接失败: {e}")
        return None

# ============== 从SQLite读取数据 ==============
def read_from_sqlite(sqlite_conn, days=30):
    """读取最近days天的数据"""
    try:
        cursor = sqlite_conn.cursor()
        cursor.execute('''
            SELECT query_date, mz_count, jz_count, ry_count
            FROM daily_report
            ORDER BY query_date DESC
            LIMIT ?
        ''', (days,))
        rows = cursor.fetchall()
        cursor.close()
        
        if not rows:
            print("未找到数据")
            return None
        
        # 反转数据按日期升序排列
        rows = list(reversed(rows))
        dates = [datetime.strptime(row[0], '%Y-%m-%d') for row in rows]
        mz_counts = [row[1] for row in rows]
        jz_counts = [row[2] for row in rows]
        ry_counts = [row[3] for row in rows]
        
        return {
            'dates': dates,
            'mz': mz_counts,
            'jz': jz_counts,
            'ry': ry_counts
        }
    except Exception as e:
        print(f"读取SQLite失败: {e}")
        return None

# ============== 节假日定义 ==============
HOLIDAYS_2026 = {
    datetime(2026, 1, 1): '元旦',
    datetime(2026, 1, 26): '春节',
    datetime(2026, 1, 27): '春节',
    datetime(2026, 1, 28): '春节',
    datetime(2026, 1, 29): '春节',
    datetime(2026, 1, 30): '春节',
    datetime(2026, 1, 31): '春节',
    datetime(2026, 2, 1): '春节',
    datetime(2026, 2, 2): '春节',
    datetime(2026, 2, 3): '春节',
    datetime(2026, 4, 4): '清明',
    datetime(2026, 4, 5): '清明',
    datetime(2026, 4, 6): '清明',
    datetime(2026, 5, 1): '劳动节',
    datetime(2026, 5, 2): '劳动节',
    datetime(2026, 5, 3): '劳动节',
    datetime(2026, 5, 4): '劳动节',
    datetime(2026, 5, 5): '劳动节',
    datetime(2026, 6, 19): '端午节',
    datetime(2026, 6, 20): '端午节',
    datetime(2026, 6, 21): '端午节',
    datetime(2026, 9, 25): '中秋节',
    datetime(2026, 9, 26): '中秋节',
    datetime(2026, 9, 27): '中秋节',
    datetime(2026, 10, 1): '国庆节',
    datetime(2026, 10, 2): '国庆节',
    datetime(2026, 10, 3): '国庆节',
    datetime(2026, 10, 4): '国庆节',
    datetime(2026, 10, 5): '国庆节',
    datetime(2026, 10, 6): '国庆节',
    datetime(2026, 10, 7): '国庆节',
    datetime(2026, 10, 8): '国庆节',
}

# ============== 生成折线图 ==============
def generate_chart(data, save_path=None):
    """生成折线图"""
    if not data:
        print("无数据可绘图")
        return None
    
    # 设置中文字体
    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
    plt.rcParams['axes.unicode_minus'] = False
    
    fig, ax = plt.subplots(figsize=(14, 7))
    
    # 标记节假日背景
    for date in data['dates']:
        if date.date() in [h.date() for h in HOLIDAYS_2026.keys()]:
            ax.axvspan(date - timedelta(hours=12), date + timedelta(hours=12), 
                      alpha=0.2, color='red')
    
    # 绑定数据
    ax.plot(data['dates'], data['mz'], marker='o', label='门诊人次', linewidth=2, markersize=6, color='#1f77b4')
    ax.plot(data['dates'], data['jz'], marker='s', label='急诊人次', linewidth=2, markersize=6, color='#ff7f0e')
    ax.plot(data['dates'], data['ry'], marker='^', label='入院人次', linewidth=2, markersize=6, color='#2ca02c')
    
    # 设置图表格式
    ax.set_xlabel('日期', fontsize=12)
    ax.set_ylabel('人次', fontsize=12)
    ax.set_title('院长查询日报趋势图', fontsize=14, fontweight='bold')
    ax.legend(loc='upper right', fontsize=10)
    ax.grid(True, linestyle='--', alpha=0.7)
    
    # 设置x轴日期格式（日期+星期）
    week_days = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
    labels = []
    for d in data['dates']:
        label = d.strftime('%m-%d') + ' ' + week_days[d.weekday()]
        if d.date() in [h.date() for h in HOLIDAYS_2026.keys()]:
            holiday_name = HOLIDAYS_2026.get(d, '')
            label += f'[{holiday_name}]'
        labels.append(label)
    
    ax.set_xticks(data['dates'])
    ax.set_xticklabels(labels, rotation=45, ha='right')
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(data['dates']) // 10)))
    
    # 在数据点上标注数值
    for x, y in zip(data['dates'], data['mz']):
        ax.annotate(str(y), (x, y), textcoords="offset points", xytext=(0, 5), ha='center', fontsize=8)
    for x, y in zip(data['dates'], data['jz']):
        ax.annotate(str(y), (x, y), textcoords="offset points", xytext=(0, 5), ha='center', fontsize=8)
    for x, y in zip(data['dates'], data['ry']):
        ax.annotate(str(y), (x, y), textcoords="offset points", xytext=(0, 5), ha='center', fontsize=8)
    
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches='tight')
        print(f"图表已保存: {save_path}")
    
    plt.close()
    return save_path

# ============== 保存到SQLite ==============
def save_to_sqlite(sqlite_conn, query_date, mz, jz, ry):
    try:
        cursor = sqlite_conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO daily_report (query_date, mz_count, jz_count, ry_count)
            VALUES (?, ?, ?, ?)
        ''', (query_date, mz, jz, ry))
        sqlite_conn.commit()
        print(f"数据已保存到SQLite: {query_date}")
        return True
    except Exception as e:
        print(f"保存SQLite失败: {e}")
        return False

# ============== 执行查询 ==============
def execute_query(connection, sql):
    try:
        cursor = connection.cursor()
        cursor.execute(sql)
        result = cursor.fetchone()
        cursor.close()
        return result[0] if result else 0
    except Exception as e:
        print(f"查询失败: {e}")
        return 0

# ============== 获取钉钉 Access Token ==============
def get_dingtalk_token():
    """获取钉钉 access_token"""
    url = "https://oapi.dingtalk.com/gettoken"
    params = {
        "appkey": DINGTALK_CORP_ID,
        "appsecret": DINGTALK_CORP_SECRET
    }
    try:
        response = requests.get(url, params=params, timeout=10)
        result = response.json()
        if result.get('errcode') == 0:
            return result.get('access_token')
        else:
            print(f"获取token失败: {result}")
            return None
    except Exception as e:
        print(f"获取token异常: {e}")
        return None

# ============== 发送钉钉消息 ==============
def send_dingtalk(data, query_date):
    message = f"### 院长查询日报\n\n"
    message += f"**查询日期**: {query_date}\n\n"
    message += "| 指标 | 人次 |\n"
    message += "|------|------|\n"
    message += f"| 门诊人次 | {data['门诊']} |\n"
    message += f"| 急诊人次 | {data['急诊']} |\n"
    message += f"| 入院人次 | {data['入院']} |\n"
    
    data_payload = {
        "msgtype": "markdown",
        "markdown": {
            "title": "院长查询日报",
            "text": message
        }
    }
    
    try:
        response = requests.post(DINGTALK_WEBHOOK, json=data_payload, timeout=10)
        if response.json().get('errcode') == 0:
            print("结果已发送到钉钉")
            return True
        else:
            print(f"钉钉发送失败: {response.text}")
            return False
    except Exception as e:
        print(f"钉钉发送异常: {e}")
        return False

# ============== 上传图片到钉钉 ==============
def upload_to_dingtalk(image_path):
    """上传图片到钉钉，获取media_id（使用企业内部API）"""
    token = get_dingtalk_token()
    if not token:
        return None
    
    try:
        with open(image_path, 'rb') as f:
            files = {'media': ('chart.png', f, 'image/png')}
            url = f"https://oapi.dingtalk.com/media/upload?access_token={token}&type=image"
            response = requests.post(url, files=files, timeout=30)
        result = response.json()
        if result.get('errcode') == 0:
            print(f"图片上传成功，media_id: {result.get('media_id')}")
            return result.get('media_id')
        else:
            print(f"图片上传失败: {result}")
            return None
    except Exception as e:
        print(f"图片上传异常: {e}")
        return None

# ============== 发送图片到钉钉（企业内部机器人）==============
def send_image_to_dingtalk(image_path, title="趋势图"):
    """发送图片到钉钉（使用企业内部机器人）"""
    token = get_dingtalk_token()
    if not token:
        return False
    
    # 构建发送消息的URL（使用群机器人的webhook）
    url = DINGTALK_WEBHOOK
    
    # 先上传图片获取media_id
    media_id = upload_to_dingtalk(image_path)
    if not media_id:
        return False
    
    # 使用webhook发送图片消息
    data_payload = {
        "msgtype": "image",
        "image": {
            "media_id": media_id
        }
    }
    
    try:
        response = requests.post(url, json=data_payload, timeout=10)
        if response.json().get('errcode') == 0:
            print("图表已发送到钉钉")
            return True
        else:
            print(f"图片发送失败: {response.text}")
            return False
    except Exception as e:
        print(f"图片发送异常: {e}")
        return False

# ============== 主程序 ==============
if __name__ == "__main__":
    # 命令行参数: 
    #   python 院长查询每日推送.py              # 执行查询推送(默认昨天)
    #   python 院长查询每日推送.py 2026-03-31    # 执行查询推送(指定日期)
    #   python 院长查询每日推送.py chart         # 从SQLite生成图表
    #   python 院长查询每日推送.py chart 30     # 从SQLite生成最近30天图表
    
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg == 'chart':
            # 生成图表模式
            days = int(sys.argv[2]) if len(sys.argv) > 2 else 30
            sqlite_conn = connect_sqlite()
            if sqlite_conn:
                data = read_from_sqlite(sqlite_conn, days)
                sqlite_conn.close()
                if data:
                    save_path = r'D:\文档\正骨医院\院长查询\data\daily_report_chart.png'
                    os.makedirs(os.path.dirname(save_path), exist_ok=True)
                    generate_chart(data, save_path)
                    # 发送到钉钉
                    send_image_to_dingtalk(save_path, f"最近{days}天趋势")
                else:
                    print("无法读取数据生成图表")
            else:
                print("无法连接SQLite数据库")
        else:
            query_date = arg
            print(f"查询日期: {query_date}")
            
            # 连接SQLite
            sqlite_conn = connect_sqlite()
            
            # 连接IRIS数据库
            iris_conn = connect_iris()
            
            if iris_conn:
                # 门诊人次查询
                sql_mz = f"SELECT COUNT(*) FROM PA_Adm WHERE DATEDIFF(day, PAADM_AdmDate, '{query_date}') = 0 AND PAADM_Type IN ('O') AND PAADM_VisitStatus = 'A'"
                mz_count = execute_query(iris_conn, sql_mz)
                print(f"门诊人次: {mz_count}")
                
                # 急诊人次查询
                sql_jz = f"SELECT COUNT(*) FROM PA_Adm WHERE DATEDIFF(day, PAADM_AdmDate, '{query_date}') = 0 AND PAADM_Type IN ('E') AND PAADM_VisitStatus = 'A'"
                jz_count = execute_query(iris_conn, sql_jz)
                print(f"急诊人次: {jz_count}")
                
                # 入院人次查询
                sql_ry = f"SELECT COUNT(*) FROM PA_Adm WHERE DATEDIFF(day, PAADM_AdmDate, '{query_date}') = 0 AND PAADM_Type IN ('I') AND PAADM_VisitStatus = 'A'"
                ry_count = execute_query(iris_conn, sql_ry)
                print(f"入院人次: {ry_count}")
                
                iris_conn.close()
                print("IRIS数据库连接已关闭")
                
                # 保存到SQLite
                if sqlite_conn:
                    save_to_sqlite(sqlite_conn, query_date, mz_count, jz_count, ry_count)
                    sqlite_conn.close()
                
                # 发送到钉钉
                result_data = {
                    '门诊': mz_count,
                    '急诊': jz_count,
                    '入院': ry_count
                }
                send_dingtalk(result_data, query_date)
            else:
                print("无法连接IRIS数据库，程序退出")
    else:
        # 默认执行昨天的查询推送
        query_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        print(f"查询日期: {query_date}")
        
        sqlite_conn = connect_sqlite()
        iris_conn = connect_iris()
        
        if iris_conn:
            sql_mz = f"SELECT COUNT(*) FROM PA_Adm WHERE DATEDIFF(day, PAADM_AdmDate, '{query_date}') = 0 AND PAADM_Type IN ('O') AND PAADM_VisitStatus = 'A'"
            mz_count = execute_query(iris_conn, sql_mz)
            print(f"门诊人次: {mz_count}")
            
            sql_jz = f"SELECT COUNT(*) FROM PA_Adm WHERE DATEDIFF(day, PAADM_AdmDate, '{query_date}') = 0 AND PAADM_Type IN ('E') AND PAADM_VisitStatus = 'A'"
            jz_count = execute_query(iris_conn, sql_jz)
            print(f"急诊人次: {jz_count}")
            
            sql_ry = f"SELECT COUNT(*) FROM PA_Adm WHERE DATEDIFF(day, PAADM_AdmDate, '{query_date}') = 0 AND PAADM_Type IN ('I') AND PAADM_VisitStatus = 'A'"
            ry_count = execute_query(iris_conn, sql_ry)
            print(f"入院人次: {ry_count}")
            
            iris_conn.close()
            print("IRIS数据库连接已关闭")
            
            if sqlite_conn:
                save_to_sqlite(sqlite_conn, query_date, mz_count, jz_count, ry_count)
                sqlite_conn.close()
            
            result_data = {
                '门诊': mz_count,
                '急诊': jz_count,
                '入院': ry_count
            }
            send_dingtalk(result_data, query_date)
        else:
            print("无法连接IRIS数据库，程序退出")