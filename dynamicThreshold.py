import xlwings as xw
import numpy as np
import copy
import math

wb = xw.Book("author.xlsx")
wc = xw.Book("keyword.xlsx")
wd = xw.Book("sequencial.xlsx")
we = xw.Book()

read_author = wb.sheets[0]
read_keyword = wc.sheets[0]
read_sequence = wd.sheets[0]
writeKeyword = we.sheets['工作表1']#writeKeyword.cells(x, y).value

threshold = 0 #author average = 4.92794
sumthreshold = 0
dataSize = 156384
keywordSize = 20343
author = np.zeros((2, dataSize))
keyword = np.zeros(keywordSize)
sequence = np.zeros((40, dataSize))
sequence_list = [[0 for y in range(8)] for x in range(dataSize)]
matchweight = [0 for x in range(dataSize)]
matchnumber = [0 for x in range(keywordSize)]
layer = []
layer0 = []
layer1 = []
layercompare = []  # 用來確認sequence是否包含keyword
layerTable = []
templayerTable = []
keywordweight = []  # 每個sequence的權重
tempkeywordweight = []
paperkeyword = []  # 跟此keyword相關的keyword在陣列的位置
copylayer = []  # 複製layer的搜尋結果
findkeyword = 0
sumofmatchweight = 0  # sequence的總權重
layertablelen = 0
maxsequencelength = 0  # 找最長的sequence長度
maxsequencesupport = 0  # 找最大的support
keywordorder = 0
sequenceorder = 1
itemposition = 0  # item在frequentitem的位置


def sequencelength(sequence):
  return sum(len(i) for i in sequence)


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
        if set(sequence[i]).issuperset(copytemp[0]):
            copytemp.pop(0)
        if not copytemp:
            return i + 1
    return -1


def finditemposition(item):
    for i in range(len(itemset)):
        if itemset[i] == item:
            return i


def checkitem(item, match):
    for i in range(len(item)):
        if sequencesuperset(item[i], match):
            return True
    return False


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

for i in range(dataSize):
    matchweight[i] = author[i][1] / 294

for i in range(keywordSize):  # 找出單個keyword位置keywordSize
    layer0.append(keyword[i])
    layer.append(layer0)
    layercompare.append(layer0)
    print(layer0)
    for j in range(dataSize):  # dataSize
        findkeyword = sequencesuperset(sequence_list[j], layer)
        if findkeyword > 0:
            paperkeyword.append(j)
            sumthreshold = sumthreshold + matchweight[j]

    if len(paperkeyword) > 0:
        threshold = (sumthreshold/len(paperkeyword))*2.5*math.log(len(paperkeyword))
        
    writeKeyword.cells(i + 1, 1).value = keyword[i]
    writeKeyword.cells(i + 1, 2).value = threshold
    print(keyword[i], sumthreshold, len(paperkeyword), threshold)
    sumthreshold = 0
    maxsequencelength = 0
    maxsequencesupport = 0
    sequenceorder = 1
    layer0.pop()
    maxlayer = []
    layercompare = []
    layer = []
    layerTable = []
    layerTable1 = []
    keywordweight = []
    paperkeyword = []
