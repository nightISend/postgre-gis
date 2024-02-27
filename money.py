import psycopg2
import openpyxl
#创星连接行象
conn=psycopg2.connect(database="BigJob",user="postgres",password="Lvu123123",host="localhost",port="5433")#数据库连接属性要改
cur=conn.cursor()#创建指针对象

file_route='f:\pg\大作业\csv'
filename=file_route+'\合计.xlsx'

sql_com='''
select owner.id,owner.name,sum(house_all.sum_money),sum(field_all.sum_money),sum(pond_all.sum_money)
from owner left join house_all on owner.id=house_all.id
left join field_all on owner.id=field_all.id
left join pond_all on owner.id=pond_all.id
group by owner.id
'''

workbook=openpyxl.Workbook()
sheet=workbook.active
sheet.append(['身份证','姓名','房屋补偿款','农田补偿款','池塘补偿款'])
cur.execute(sql_com)
rows=cur.fetchall()
for row in rows:
    sheet.append(row)
workbook.save(filename)

conn.commit()
cur.close()
conn.close()