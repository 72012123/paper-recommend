import pymysql
import xlwings as xw
import numpy as np
conn = pymysql.connect(host = '127.0.0.1', port = 3306, user = 'root', passwd='', db = 'paper_recommend', charset='utf8')
cursor = conn.cursor()
sql = "SELECT * FROM `paper_keyword` ORDER BY `k_id` ASC"
temp = 0
count = -1
workbook = xw.Book()
sheet = workbook.sheets['工作表1']
keywordCount = np.zeros((294167, 2))#294167 paper_keyword
cursor.execute(sql)
results = cursor.fetchall()
for row in results:
    sid = row[0]
    k_id = row[1]
    p_id = row[2]
    if temp != int(k_id):
        temp = int(k_id)
        count += 1
        keywordCount[count][0] = temp

    keywordCount[count][1] += 1
temp = 1
for i in range(294167):
    if keywordCount[i][1] > 5:
        sheet.cells(temp + 1, 1).value = keywordCount[i][0]
        sheet.cells(temp + 1, 2).value = keywordCount[i][1]
        temp = temp + 1

print(count)
conn.close()