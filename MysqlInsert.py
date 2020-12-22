'''
通过xlrd和pymysql库实现将excel的数据导入到mysql数据库中
'''

from datetime import datetime
import pymysql
import xlrd
from xlrd import  xldate_as_tuple

# 文件路径
file = r'D:\\SVNblog\\数据导入模板V1.2.xlsx'
# 通过关键参数进行连接mysql
conn = pymysql.connect(host='localhost', port=3306, user='root', passwd='root', db='test', charset='utf8')
# 创建游标
cur = conn.cursor()

list_list = []
# 打开excel
book = xlrd.open_workbook(file)
# 查找excel中名叫‘xxx’的sheet页
sheet_other = book.sheet_by_name('汇总性数据')
# 获取这个excel中所有的sheet页名称
sheet_all = book.sheet_names()

row_num = sheet_other.nrows  # 获取当前sheet页的行数量

col_num = sheet_other.ncols  # 获取当前sheet页的列数量

try:
    for i in range(1, row_num):
        # 将每一行的数据存入row_data中
        row_data = sheet_other.row_values(i)
        # 将字符串修改成日期格式输出就可以插入数据库的日期格式字段内了
        # 读取excel中日期格式会显示一串数字

        date = datetime(*xldate_as_tuple(row_data[5], 0))

        cell = date.strftime('%y-%m-%d')

        # 向数据库中插入NUll值可以写’None‘就代表NUll
        value = (row_data[0], '', row_data[1], row_data[2], row_data[3], row_data[4], cell, '')

        list_list.append(value)

    insert_sql = 'insert into data_list(yuan_data_table_nbr,data_key,data_type1,date_type2,date_type3,data_type,data_date,data_create_time) values(%s,%s,%s,%s,%s,%s,%s,%s)'

    Delete_sql = 'delete from data_list'

    # 执行sql语句
    cur.execute(Delete_sql)  # 一条一条执行

    cur.executemany(insert_sql, list_list)  # 批量执行
    # 提交语句
    conn.commit()

except Exception as e:
    # 出现异常打印异常信息
    print("[Err]:" + str(e))
    # 出错将执行语句进行回滚
    conn.rollback()

conn.close()
