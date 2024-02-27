import psycopg2
import openpyxl
#创星连接行象
conn=psycopg2.connect(database="BigJob",user="postgres",password="Lvu123123",host="localhost",port="5433")#数据库连接属性要改
cur=conn.cursor()#创建指针对象

#excel存储位置
file_route='''f:\pg\大作业\csv''';

#建临时表并提取与道路相交的农田编号及区域（ok）
pond_select='''
Create Table pond_select(
  num varchar(10),
  geom Geometry(MULTIPOLYGON,32618)
);
INSERT INTO pond_select(num,geom)
select pond_geom.num,ST_intersection(hcq.m,pond_geom.geom)
from (select ST_union(ST_buffer(geom,width/2)) as m from road) AS hcq left join pond_geom on ST_intersects(hcq.m,pond_geom.geom);
'''

#计算农田补偿金,编号，面积，几何(ok)
pond_area='''
 create table pond_area(
    num varchar(10),
    area float,
    id varchar(20),
    area_money float default 0,
    geom Geometry(MULTIPOLYGON,32618),
    PRIMARY KEY(num)
    );
 INSERT INTO pond_area(num,id,area,geom,area_money)
 select pond_select.num,pond.id,ST_area(pond_select.geom) as area,pond_select.geom,ST_area(pond_select.geom)*pond_type.price
 from pond right join pond_select on pond.num=pond_select.num
 left join pond_type on pond_type.type=pond.type;
'''

#计算房屋人头费,人头费=人数*该地的补偿金*0.01
pond_mem='''
 create table pond_mem(
  num varchar(10),
  mem_money int default 0
  );
  INSERT INTO pond_mem(num,mem_money)
  select pond_area.num,owner.members*pond_area.area_money*0.01 
  from pond_area inner join owner on pond_area.id=owner.id
'''

#房屋总表插值(这里的几何只包含算钱的区域)
pond_all='''
CREATE TABLE pond_all(
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
INSERT INTO pond_all(num,type,price,type_name,area_money,members_price,sum_money,area,id,owner_name,members,phone,address,geom)
select  pond_area.num,
        pond_type.type,
        pond_type.price,
        pond_type.name,
        pond_area.area_money,
        pond_mem.mem_money,
        pond_area.area_money+pond_mem.mem_money,
        pond_area.area,
        owner.id,
        owner.name,
        owner.members,
        owner.phone,
        owner.address,
        pond_area.geom
from pond_area left join pond_mem on pond_mem.num=pond_area.num
left join pond on pond_area.num=pond.num
left join owner on pond.id=owner.id
left join pond_type on pond_type.type=pond.type
'''

#查看错误数据,以num为主键，显示对应记录id，type或空间数据错误
pond_fault='''
create table pond_fault(
  num varchar(10),
  id varchar(20),
  type varchar(20),
  geom Geometry(MULTIPOLYGON,32618)
);
insert into pond_fault(num,id,type,geom)
select pond.num,owner.id,pond_type.type,pond_geom.geom
from pond left join pond_type on pond.type=pond_type.type
left join owner on pond.id=owner.id
left join pond_geom on pond_geom.num=pond.num
where owner.id is null or pond_type.type is null or pond_geom.geom is null;
'''

#删除临时表
sql_drop='''
drop table pond_select;
drop table pond_area;
drop table pond_mem;
'''

#查表语句
sql_all='''
SELECT * FROM pg_tables WHERE schemaname = 'public' and tablename='pond_all';
'''
sql_fault='''
SELECT * FROM pg_tables WHERE schemaname = 'public' and tablename='pond_fault';
'''

#查询pond_all表中除geom字段以外的所有字段
sql_select='''
select num,type,type_name,price,area,area_money,members_price,sum_money,id,owner_name,members,phone,address from pond_all
'''

cur.execute(pond_select)
cur.execute(pond_area)
cur.execute(pond_mem)

#判断总表是否存在,并删除临时表
cur.execute(sql_all)
result=cur.fetchall()
if len(result):
    cur.execute('drop table pond_all')
cur.execute(pond_all)
cur.execute(sql_drop)

#判断错误表是否存在
cur.execute(sql_fault)
result=cur.fetchall()
if len(result):
    cur.execute('drop table pond_fault')
cur.execute(pond_fault)

#生成excel表
filename=file_route+'\池塘.xlsx'
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



