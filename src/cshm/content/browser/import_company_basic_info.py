# -*- coding: utf-8 -*-
import csv
from sqlalchemy import create_engine

class SqlObj:

    def execSql(self, execStr):
        dbString = 'mysql+mysqldb://cshm:cshm@localhost/cshm?charset=utf8'
        engine = create_engine(dbString, echo=True)

        conn = engine.connect() # DB連線
        execResult = conn.execute(execStr)
        conn.close()
        if execResult.returns_rows:
            return execResult.fetchall()


with open('/home/andy/BGMOPEN1.csv') as file:
    reader = csv.DictReader(file)
    count = 0
    errorCount = 0
    sqlInstance = SqlObj()
    for row in reader:
        if not (row['統一編號'] and row['營業人名稱']):
            continue
        sqlStr = "INSERT INTO company_basic_info (`tax-no`, `company-name`) VALUES ('%s', '%s')" % (row['統一編號'], row['營業人名稱'])
        try:
            sqlInstance.execSql(sqlStr)
            count += 1
            print 'OK, %s' % count
        except:
            errorCount += 1
            print 'Fail, %s' % errorCount
            pass
    print 'FINISH, %s sussess, %s Error' % (count, errorCount)
