import psycopg2
import openpyxl
#创星连接行象
conn=psycopg2.connect(database="BigJob",user="postgres",password="Lvu123123",host="localhost",port="5433")#数据库连接属性要改
cur=conn.cursor()#创建指针对象

file_route='f:\pg\大作业\csv'

#建临时表并提取与道路相交的房屋编号的前6位
house_select='''
Create Table house_select(
num varchar(10)
);
INSERT INTO house_select(num)
select distinct left(house_geom.num,6)
from (select ST_union(ST_buffer(geom,width/2)) as m from road) as hcq left join house_geom on ST_intersects(hcq.m,house_geom.geom);
'''

#计算房屋补偿金,编号，户主，面积
house_area='''
  create table house_area(
  num varchar(10),
  id varchar(20),
  area_money float default 0,
  area float
  );
  INSERT INTO house_area(num,id,area,area_money)
  select house.num,house.id,ST_area(house_geom.geom),ST_area(house_geom.geom)*house.floor*house_type.price
  from house_select left join house on house_select.num=left(house.num,6)
  left join house_geom on house_geom.num=house.num
  left join house_type on house.type=house_type.type;
'''

#计算房屋编号，人头费,人头费=人数*该地的补偿金*0.01
house_mem='''
 create table house_mem(
  num varchar(10),
  mem_money int default 0
  );
  INSERT INTO house_mem(num,mem_money)
  select house_area.num,owner.members*house_area.area_money*0.01
  from house_area left join owner on house_area.id=owner.id
'''

#房屋总表插值
house_all='''
CREATE TABLE house_all(
  num varchar(10),
  type varchar(10),
  floor int,
  price int,
  type_name varchar(10),
  area_money float default 0,
  members_price int default 0,
  sum_money float default 0,
  area float,
  geom Geometry(MULTIPOLYGON,32618),
  id varchar(18),
  owner_name varchar(10),
  members int,
  phone varchar(11),
  address varchar(30),
  PRIMARY KEY(num)
);
INSERT INTO house_all(num,type,floor,price,type_name,area_money,members_price,sum_money,area,id,owner_name,members,phone,address,geom)
select  house_area.num,
        house_type.type,
        house.floor,
        house_type.price,
        house_type.name,
        house_area.area_money,
        house_mem.mem_money,
        house_area.area_money+house_mem.mem_money,
        house_area.area,
        owner.id,
        owner.name,
        owner.members,
        owner.phone,
        owner.address,
        house_geom.geom
from house_area left join house_mem on house_mem.num=house_area.num
left join house on house_area.num=house.num
left join owner on house.id=owner.id
left join house_type on house_type.type=house.type
left join house_geom on house_geom.num=house_area.num
'''

#删除临时表
sql_drop='''
drop table house_select;
drop table house_area;
drop table house_mem;
'''

#查看错误数据,以num为主键，显示对应记录id，type或空间数据错误
house_fault='''
create table house_fault(
  num varchar(10),
  id varchar(20),
  type varchar(20),
  geom Geometry(MULTIPOLYGON,32618)
);
insert into house_fault(num,id,type,geom)
select house.num,owner.id,house_type.type,house_geom.geom
from house left join house_type on house.type=house_type.type
left join owner on house.id=owner.id
left join house_geom on house_geom.num=house.num
where owner.id is null or house_type.type is null or house_geom.geom is null;
'''

#查表语句
sql_table='''
SELECT * FROM pg_tables WHERE schemaname = 'public' and tablename='house_all';
'''
sql_fault='''
SELECT * FROM pg_tables WHERE schemaname = 'public' and tablename='house_fault';
'''

#查询field_all表中除geom字段以外的所有字段
sql_select='''
select num,type,type_name,price,floor,area_money,members_price,sum_money,area,id,owner_name,members,phone,address from house_all
'''

cur.execute(house_select)
cur.execute(house_area)
cur.execute(house_mem)

#判断总表是否存在
cur.execute(sql_table)
result=cur.fetchall()
if len(result):
    cur.execute('drop table house_all')
cur.execute(house_all)
cur.execute(sql_drop)

#判断错误表是否存在
cur.execute(sql_fault)
result=cur.fetchall()
if len(result):
    cur.execute('drop table house_fault')
cur.execute(house_fault)

#生成excel表
filename=file_route+'\房屋.xlsx'
workbook=openpyxl.Workbook()
sheet=workbook.active
sheet.append(['编号','房屋类型编号','房屋类型','单价（平方米）','层数','补偿款','人头费','总价','面积','户主身份证','户主','该户人口','联系电话','地址'])
cur.execute(sql_select)
rows=cur.fetchall()
for row in rows:
    sheet.append(row)
workbook.save(filename)

conn.commit()
cur.close()
conn.close()