import xlwings as xw
import numpy as np
import copy
import itertools
from collections import defaultdict
from operator import itemgetter
import pandas as pd
import io

wb = xw.Book("author.xlsx")
wc = xw.Book("keyword.xlsx")
wd = xw.Book("sequencial.xlsx")
#we = xw.Book()

read_author = wb.sheets[0]
read_keyword = wc.sheets[0]
read_sequence = wd.sheets[0]
#writeKeyword = we.sheets['工作表1']#writeKeyword.cells(x, y).value

dataSize = 156384
keywordSize = 20343
author = np.zeros((2, dataSize))
keyword = np.zeros(keywordSize)
sequence = np.zeros((40, dataSize))
sequence_list = [[0 for y in range(8)] for x in range(dataSize)]
matchweight = [1 for x in range(dataSize)]
matchtable = [[0 for y in range(0)] for x in range(keywordSize)]#每個keyword配對成功的
matchnumber = [0 for x in range(keywordSize)]#每個keyword組合符合的數量
layer = []
layer0 = []
layer1 = []
findkeyword = 0

def delzero(sequenceset):
    result = []
    for i in range(len(sequenceset)):
        temp = []
        for j in range(len(sequenceset[i])):
            if sequenceset[i][j] != 0:
                temp.append(sequenceset[i][j])
        if not temp == []:
            result.append(temp)
    return result


def sequencesuperset(sequence, match):
    copytemp = copy.deepcopy(match)
    for i in range(len(sequence)):
        if(len(match) > 1):
          if set(sequence[i]).issuperset(copytemp[0]) == True:
              copytemp.pop(0)
        else:
          if set(sequence[i]).issuperset(copytemp) == True:
              return i + 1
        if not copytemp:
            return i + 1
    return -1


author = read_author.range('a2:b156385').value
author = [[round(x) for x in y] for y in author]
keyword = read_keyword.range('a2:a20344').value
keyword = [round(x) for x in keyword]
sequence = read_sequence.range('a1:an156384').value
sequence = [[round(x) for x in y] for y in sequence]
for i in range(dataSize):
    sequence_list[i][0] = sequence[i][0:5]
    sequence_list[i][1] = sequence[i][5:10]
    sequence_list[i][2] = sequence[i][10:15]
    sequence_list[i][3] = sequence[i][15:20]
    sequence_list[i][4] = sequence[i][20:25]
    sequence_list[i][5] = sequence[i][25:30]
    sequence_list[i][6] = sequence[i][30:35]
    sequence_list[i][7] = sequence[i][35:40]
    sequence_list[i] = delzero(sequence_list[i])

for i in range(keywordSize):#找出單個keyword位置
    layer.append(keyword[i])
    print(layer)
    for j in range(100):
        findkeyword = sequencesuperset(sequence_list[j], layer)
        if findkeyword < len(sequence_list[j]) and findkeyword > 0:
            for k in range(findkeyword, len(sequence_list[j])):
                for x in sequence_list[j][k]:
                    #把x出現的次數記錄在matchnumber
                    for l in range(keywordSize):
                        if x == keyword[l]:
                            matchnumber[l] = matchnumber[l] + matchweight[j]
    for j in range(keywordSize):
        if matchnumber[j] >= 1:
            matchtable[i].append(keyword[j])
        matchnumber[j] = 0
    layer.pop(0)

#for i in range(keywordSize):
#    if len(matchtable[i]) > 0:
#        print(matchtable[i])
