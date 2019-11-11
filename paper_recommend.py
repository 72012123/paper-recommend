import pymysql
import xlwings as xw
import numpy as np
conn = pymysql.connect(host = '127.0.0.1', port = 3306, user = 'root', passwd='', db = 'paper_recommend', charset='utf8')
cursor = conn.cursor()

wb = xw.Book("author.xlsx")
wc = xw.Book("keyword.xlsx")
wd = xw.Book()

read_author = wb.sheets[0]
read_keyword = wc.sheets[0]
writeKeyword = wd.sheets['工作表1']

author = np.zeros(156384)
author = author.astype(int)
keyword = np.zeros(20343)
keyword = keyword.astype(int)
paper = np.zeros(300)
paper = paper.astype(int)
yearClassification = np.zeros((8, 300))
yearClassification = yearClassification.astype(int)
sortKeyword = np.zeros((8, 20343))
sortKeyword = sortKeyword.astype(int)

a = read_author.range('a2:a156385').value
b = read_keyword.range('a2:a20344').value

for i in range(156384):
    author[i] = a[i]

for i in range(20343):
    keyword[i] = b[i]

for i in range(156384):#依照a_id找到每篇論文的年份並分類
    temp = author[i]
    sql = "SELECT * FROM `paper_author` WHERE `a_id` = %s"%(str(temp))
    cursor.execute(sql)
    results = cursor.fetchall()
    paperAmount = 0

    for row in results:
        paper[paperAmount] = row[2]
        paperAmount = paperAmount + 1

    for j in range(paperAmount):
        tempPaper = paper[j]
        sql2 = "SELECT * FROM `paper_sort_out` WHERE `id` = %s"%(str(tempPaper))
        cursor.execute(sql2)
        result = cursor.fetchone()
        if result is not None:
            record = 0
            while yearClassification[2017 - result[3]][record] != 0 :
                record = record + 1
            yearClassification[2017 - result[3]][record] = paper[j]
            #print(author[i], yearClassification[2017 - result[3]][record], result[3], record)

    for j in range(8):#每個作者所參與的論文轉換成keyword
        for r in range(300):
            if yearClassification[j][r] != 0:
                tempPaper = yearClassification[j][r]
                sql3 = "SELECT * FROM `paper_keyword` WHERE `p_id` = %s ORDER BY `paper_keyword`.`k_id` ASC" %\
                       (str(tempPaper))
                cursor.execute(sql3)
                keywordResult = cursor.fetchall()
                keywordPosition = 0
                for keywordRow in keywordResult:
                    tempKeyword = keywordRow[1]#拿到該篇的keyword
                    while (keywordPosition < 20342) and (keyword[keywordPosition] < int(tempKeyword)):
                        keywordPosition = keywordPosition + 1
                    if keyword[keywordPosition] == int(tempKeyword):
                        sortKeyword[j][keywordPosition] = sortKeyword[j][keywordPosition] + 1
                        #print(2017-j, author[i], yearClassification[j][r], tempKeyword, sortKeyword[j][keywordPosition])
    for j in range(8):#把每個年份最常出現的keyword排序
        for k in range(5):
            maxPosition = 0
            maxKeywordNumber = 0
            for r in range(20343):
                if sortKeyword[j][r] > maxKeywordNumber:
                    maxPosition = r
                    maxKeywordNumber = sortKeyword[j][r]
            if maxKeywordNumber == 0:
                print(2017 - j, author[i], 0, 0)
                writeKeyword.cells(i + 1, (k + 1) + j * 5).value = 0
            else:
                print(2017 - j, author[i], keyword[maxPosition], sortKeyword[j][maxPosition])
                writeKeyword.cells(i + 1, (k + 1) + j * 5).value = keyword[maxPosition]
            sortKeyword[j][maxPosition] = 0
    for j in range(8):
        for r in range(300):
            yearClassification[j][r] = 0
        for k in range(20343):
            sortKeyword[j][k] = 0
conn.close()