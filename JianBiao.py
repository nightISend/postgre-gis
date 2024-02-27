import psycopg2
#创星连接行象
conn=psycopg2.connect(database="BigJob",user="postgres",password="Lvu123123",host="localhost",port="5433")#数据库连接属性要改
cur=conn.cursor()#创建指针对象

sql='''
    CREATE TABLE owner(
        id varchar(18),
        name varchar(10),
        members int,
        phone varchar(11),
        address varchar(30),
        PRIMARY KEY(id)
    );
    CREATE TABLE field(
        num varchar(10),
        id varchar(20),
        type varchar(20),
        PRIMARY KEY(num)
    );
    CREATE TABLE pond(
        num  varchar(10),
        id varchar(20),
        type varchar(20),
        PRIMARY KEY(num)
    );
    CREATE TABLE house(
        num varchar(10),
        id varchar(20),
        floor int,
        type varchar(20),
        PRIMARY KEY(num)
     );
    CREATE TABLE house_type(
        type varchar(20),
        price int,
        name varchar(20),
        PRIMARY KEY(type)
    );
    CREATE TABLE pond_type(
        type varchar(20),
        price int,
        name varchar(20),
        PRIMARY KEY(type)
    );
    CREATE TABLE field_type(
        type varchar(20),
        price int,
        name varchar(20),
        PRIMARY KEY(type)
    );
'''
cur.execute(sql)

conn.commit()
cur.close()
conn.close()