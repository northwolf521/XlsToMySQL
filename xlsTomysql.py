import win32com.client as win32 
import MySQLdb

xl = win32.gencache.EnsureDispatch('Excel.Application')
xlbook = win32.Dispatch('Excel.Application').Workbooks.Open('D:\\xls\\opsdata.xls')
sh = xlbook.Worksheets('sheet1')

dfun = []
jcount = 2
ncount = 1
for ncount in range(1, 65566):
    if sh.Cells(ncount, 1).Value == None:
        break
    else:
        continue

#EXCEL的数据安排为第一行是字段，第二行开始是数据，故从第2行开始循环，将两列数据合并到一个LIST中
#LIST结构为[('a','b'),('c','d')]
for jcount in range(2, ncount):
    dfun.append((sh.Cells(jcount, 1).Value, sh.Cells(jcount, 2).Value))

fo = []
icount = 1
for icount in range(1, 2):
    fo.append((sh.Cells(1, icount).Value, sh.Cells(1, icount + 1).Value))

#打开MYSQL链接
conn = MySQLdb.connect(host='10.10.3.95',user='root',passwd='',db='test')

cursor = conn.cursor()
cursor.execute("create table `test`.`test`(" + fo[0][0] + " varchar(100)," + fo[0][1] + " varchar(100));")
cursor.executemany("""insert into test values(%s, %s);""" ,dfun)

conn.commit()

#执行查询检查结果
count = cursor.execute('select * from test')
print 'has %s record' % count

#重置游标位置
cursor.scroll(0, mode = 'absolute')

#搜取所有结果
results = cursor.fetchall()

#获取MYSQL里的数据字段
fields = cursor.description

#将字段写入到EXCEL新表的第一行
sh2 = xlbook.Worksheets('sheet3')

#清空sheet3
sh2.Cells.Clear
for ifs in range(1, len(fields) + 1):
    sh2.Cells(1, ifs).Value = fields[ifs - 1][0]

ics = 2
jcs = 1
for ics in range(2, len(results) + 2):
    for jcs in range(1, len(fields) + 1):
        sh2.Cells(ics, jcs).Value = results[ics - 2][jcs - 1]

xl.Application.Quit()

cursor.close()

conn.close()
