#coding:utf-8
import pymysql

def get_userinfo():
    key_list=['user','password']
    user_list=[]
    data_list=[]
    conn=pymysql.connect(
        port=3307,
        database='phpmyadmin',
        user='root',
        passwd='',
        host='192.168.10.42',
        charset='utf8')

    cur = conn.cursor()
    aa = cur.execute('select user_name,password from ecs_users')
    value = cur.fetchmany(aa)
    for i in value:
        user_list.append(list(i))
        for k in user_list:
            tmp = zip(key_list, k)
        data_list.append(dict(tmp))
    cur.close()
    conn.commit()
    conn.close()
    return value

if __name__=='__main__':
    re=get_userinfo()
    print(re)
