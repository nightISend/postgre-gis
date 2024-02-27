import psycopg2
import openpyxl
#创星连接行象
conn=psycopg2.connect(database="BigJob",user="postgres",password="Lvu123123",host="localhost",port="5433")#数据库连接属性要改
cur=conn.cursor()#创建指针对象

#excel存储位置
file_route='f:\pg\大作业\csv'

#建临时表并提取与道路相交的农田编号及区域（ok）
field_select='''
Create Table field_select(
  num varchar(10),
  geom Geometry(MULTIPOLYGON,32618)
);
INSERT INTO field_select(num,geom)
select field_geom.num,ST_intersection(hcq.m,field_geom.geom)
from (select ST_union(ST_buffer(geom,width/2)) as m from road) AS hcq left join field_geom on ST_intersects(hcq.m,field_geom.geom);
'''

#计算农田补偿金,编号，面积，几何(ok)
field_area='''
 create table field_area(
    num varchar(10),
    area float,
    id varchar(20),
    area_money float default 0,
    geom Geometry(MULTIPOLYGON,32618),
    PRIMARY KEY(num)
    );
 INSERT INTO field_area(num,id,area,geom,area_money)
 select field_select.num,field.id,ST_area(field_select.geom) as area,field_select.geom,ST_area(field_select.geom)*field_type.price
 from field right join field_select on field.num=field_select.num
 left join field_type on field_type.type=field.type;
'''

#计算房屋人头费,人头费=人数*该地的补偿金*0.01
field_mem='''
 create table field_mem(
  num varchar(10),
  mem_money int default 0
  );
  INSERT INTO field_mem(num,mem_money)
  select field_area.num,owner.members*field_area.area_money*0.01 
  from field_area inner join owner on field_area.id=owner.id
'''

#房屋总表插值(这里的几何只包含算钱的区域)
field_all='''
CREATE TABLE field_all(
  num varchar(10),
  type varchar(10),
  price int,
  type_name varchar(10),
  area_money float default 0,
  members_price int default 0,
  sum_money float default 0,
  area float,
  id varchar(18),
  owner_name varchar(10),
  members int,
  phone varchar(11),
  address varchar(30),
  geom Geometry(MULTIPOLYGON,32618),
  PRIMARY KEY(num)
);
INSERT INTO field_all(num,type,price,type_name,area_money,members_price,sum_money,area,id,owner_name,members,phone,address,geom)
select  field_area.num,
        field_type.type,
        field_type.price,
    field_type.name,
    field_area.area_money,
        field_mem.mem_money,
    field_area.area_money+field_mem.mem_money,
        field_area.area,
    owner.id,
        owner.name,
        owner.members,
        owner.phone,
    owner.address,
     field_area.geom
from field_area left join field_mem on field_mem.num=field_area.num
left join field on field_area.num=field.num
left join owner on field.id=owner.id
left join field_type on field_type.type=field.type
'''

#查看错误数据,以num为主键，显示对应记录id，type或空间数据错误
field_fault='''
create table field_fault(
  num varchar(10),
  id varchar(20),
  type varchar(20),
  geom Geometry(MULTIPOLYGON,32618)
);
insert into field_fault(num,id,type,geom)
select field.num,owner.id,field_type.type,field_geom.geom
from field left join field_type on field.type=field_type.type
left join owner on field.id=owner.id
left join field_geom on field_geom.num=field.num
where owner.id is null or field_type.type is null or field_geom.geom is null;
'''

#删除临时表
sql_drop='''
drop table field_select;
drop table field_area;
drop table field_mem;
'''

#查表语句
sql_all='''
SELECT * FROM pg_tables WHERE schemaname = 'public' and tablename='field_all';
'''
sql_fault='''
SELECT * FROM pg_tables WHERE schemaname = 'public' and tablename='field_fault';
'''

#查询field_all表中除geom字段以外的所有字段
sql_select='''
select num,type,type_name,price,area,area_money,members_price,sum_money,id,owner_name,members,phone,address from field_all
'''

cur.execute(field_select)
cur.execute(field_area)
cur.execute(field_mem)

#判断总表是否存在,并删除临时表
cur.execute(sql_all)
result=cur.fetchall()
if len(result):
    cur.execute('drop table field_all')
cur.execute(field_all)
cur.execute(sql_drop)

#判断错误表是否存在
cur.execute(sql_fault)
result=cur.fetchall()
if len(result):
    cur.execute('drop table field_fault')
cur.execute(field_fault)

#生成excel表
filename=file_route+'\农田.xlsx'
workbook=openpyxl.Workbook()
sheet=workbook.active
sheet.append(['编号','作物类型编号','作物类型','单价（平方米）','面积','土地补偿款','人口补偿','总补偿','户主身份证','户主','该户人口','联系电话','地址'])
cur.execute(sql_select)
rows=cur.fetchall()
for row in rows:
    sheet.append(row)
workbook.save(filename)

conn.commit()
cur.close()
conn.close()


