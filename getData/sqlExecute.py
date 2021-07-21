# FILENAME:sqlExecute.py

import pymysql

def sqlConn():
    sqlconn = pymysql.connect(host='127.0.0.1', port=3306, user='root', passwd='root', db='pdf')
    return sqlconn

# 入库文本
def insertText(args):
    # 创建连接
    conn = sqlConn()
    # conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', passwd='root', db='pdf')
    # 创建游标
    cursor = conn.cursor()
    # 执行SQL，并返回收影响行数
    cursor.execute("insert into pdf(text) values (%s)", [args, ])
    # 提交，不然无法保存新建或者修改的数据
    conn.commit()
    # 关闭游标
    cursor.close()
    # 关闭连接
    conn.close()

# 入库图片路径
def insertPic(args):
    # 创建连接
    conn = sqlConn()
    # conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', passwd='root', db='pdf')
    # 创建游标
    cursor = conn.cursor()
    # 执行SQL，并返回收影响行数
    cursor.execute("insert into pic(picPath) values (%s)", [args, ])
    # 提交，不然无法保存新建或者修改的数据
    conn.commit()
    # 关闭游标
    cursor.close()
    # 关闭连接
    conn.close()

# 入库Excel
def insertXls(args):
    # 创建连接
    conn = sqlConn()
    # conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', passwd='root', db='pdf')
    # 创建游标
    cursor = conn.cursor()
    # 执行SQL，并返回收影响行数
    cursor.execute("insert into xls(xlsPath) values (%s)", [args, ])
    # 提交，不然无法保存新建或者修改的数据
    conn.commit()
    # 关闭游标
    cursor.close()
    # 关闭连接
    conn.close()

# 查询到的fileName入库
def insertFileName(args):
    # 创建连接
    conn = sqlConn()
    # conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', passwd='root', db='pdf')
    # 创建游标
    cursor = conn.cursor()
    # 执行SQL，并返回收影响行数
    cursor.execute("insert into filename(fName) values (%s)", [args, ])
    # 提交，不然无法保存新建或者修改的数据
    conn.commit()
    # 关闭游标
    cursor.close()
    # 关闭连接
    conn.close()

def searchInDB(args):
    conn = sqlConn()
    # conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', passwd='root', db='pdf')
    cursor = conn.cursor()
    cursor.execute("select fName from filename where fName=%s", [args, ])
    # 获取第一行数据
    effFile = cursor.fetchone()
    conn.commit()
    cursor.close()
    conn.close()
    return effFile

# 插入子文件夹名称
def insertSubFolder(args):
    # 创建连接
    conn = sqlConn()
    # conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', passwd='root', db='pdf')
    # 创建游标
    cursor = conn.cursor()
    # 执行SQL，并返回收影响行数
    cursor.execute("insert into dirs(dirName) values (%s)", [args, ])
    # 提交，不然无法保存新建或者修改的数据
    conn.commit()
    # 关闭游标
    cursor.close()
    # 关闭连接
    conn.close()

# 查询数据库中子文件夹名
def searchSubFolder(args):
    conn = sqlConn()
    # conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', passwd='root', db='pdf')
    cursor = conn.cursor()
    cursor.execute("select dirName from dirs where dirName=%s", [args, ])
    # 获取第一行数据
    effFile = cursor.fetchone()
    conn.commit()
    cursor.close()
    conn.close()
    return effFile
