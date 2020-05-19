import pymysql
import xlwings as xw
import numpy as np

conn = pymysql.connect(host = '127.0.0.1', port = 3306, user = 'root', passwd='', db = 'paper_recommend', charset='utf8')
cursor = conn.cursor()

wf = xw.Book()
writeWeight = wf.sheets['工作表1']

paper = [[0 for y in range(3)] for x in range(286740)]
author = [0 for x in range(554689)]
authorWeight = [0 for x in range(554689)]
paperAmount = 0
authorAmount = 0
paperWeight = 0

sql = "SELECT `id`,`citation` FROM `paper_sort_out` ORDER BY `id` ASC"
cursor.execute(sql)
results = cursor.fetchall()
for row in results:
    paper[paperAmount][0] = row[0]
    paper[paperAmount][1] = row[1]
    paperAmount = paperAmount + 1

sql = "SELECT * FROM `author` ORDER BY `a_id` ASC"
cursor.execute(sql)
results = cursor.fetchall()
for row in results:
    author[authorAmount] = row[0]
    authorAmount = authorAmount + 1
print("step1")
for i in range(286740):  # 286740
    tempPaper = paper[i][0]
    authorAmount = 0
    print(i)
    sql = "SELECT * FROM `paper_author` WHERE `p_id` = %s ORDER BY `paper_author`.`a_id` ASC" %\
                       (str(tempPaper))
    cursor.execute(sql)
    results = cursor.fetchall()
    for row in results:
        authorAmount = authorAmount + 1
    paper[i][2] = authorAmount
print("step2")
for i in range(554689):  # 554689
    tempAuthor = author[i]
    print(i)
    sql = "SELECT * FROM `paper_author` WHERE `a_id` = %s ORDER BY `paper_author`.`p_id` ASC" %\
                       (str(tempAuthor))
    cursor.execute(sql)
    results = cursor.fetchall()
    order = 0
    for row in results:
        for j in range(order, 286740):
            if row[2] == paper[j][0]:
                authorWeight[i] = authorWeight[i] + (paper[order][1]/paper[order][2])
                order = j
                break
print("step3")
for i in range(286740):
    tempPaper = paper[i][0]
    print(i)
    sql = "SELECT * FROM `paper_author` WHERE `p_id` = %s ORDER BY `paper_author`.`a_id` ASC" % \
          (str(tempPaper))
    cursor.execute(sql)
    results = cursor.fetchall()
    order = 0
    for row in results:
        for j in range(order, 554689):
            if row[1] == author[j]:
                paperWeight = paperWeight + authorWeight[j]
                order = j
                break
    print(tempPaper, paperWeight + paper[i][1])
    writeWeight.cells(i + 1, 1).value = tempPaper
    writeWeight.cells(i + 1, 2).value = paperWeight + paper[i][1]
    paperWeight = 0
