import xlwings as xw
import numpy as np
import copy

wb = xw.Book("author.xlsx")
wc = xw.Book("keyword.xlsx")
wd = xw.Book("sequencial.xlsx")
we = xw.Book()

read_author = wb.sheets[0]
read_keyword = wc.sheets[0]
read_sequence = wd.sheets[0]
writeKeyword = we.sheets['工作表1']#writeKeyword.cells(x, y).value

threshold = 0  # 0.03
dataSize = 156384
keywordSize = 20343
author = np.zeros((2, dataSize))
keyword = np.zeros(keywordSize)
sequence = np.zeros((40, dataSize))
sequence_list = [[0 for y in range(8)] for x in range(dataSize)]
matchweight = [0 for x in range(dataSize)]
matchtable = [[0 for y in range(0)] for x in range(keywordSize)]  # 每個keyword配對成功的
matchnumber = [0 for x in range(keywordSize)]  # 每個keyword組合符合的數量
layer = []
layer0 = []
layer1 = []
layer2 = []
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
        if len(match) > 1:
          if set(sequence[i]).issuperset(copytemp[0]):
              copytemp.pop(0)
        else:
          if set(sequence[i]).issuperset(copytemp):
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

for i in range(dataSize):
    matchweight[i] = author[i][1] / 294

for i in range(keywordSize):  # 找出單個keyword位置keywordSize
    layer0.append(keyword[i])
    layer.append(layer0)
    print(layer0)
    for j in range(100):  # dataSize
        findkeyword = sequencesuperset(sequence_list[j], layer0)
        if len(sequence_list[j]) > findkeyword > 0:
            paperkeyword.append(j)
            for k in range(findkeyword, len(sequence_list[j])):
                for x in sequence_list[j][k]:
                    # 把x出現的次數記錄在matchnumber
                    for l in range(keywordSize):  # keywordSize
                        if x == keyword[l]:
                            matchnumber[l] = matchnumber[l] + matchweight[j]
    for j in range(keywordSize):  # keywordSize
        if matchnumber[j] > threshold:
            matchtable[i].append(keyword[j])
        matchnumber[j] = 0

    for x in range(4):
        if (x % 2) == 0:
            for j in range(len(matchtable[i])):
                layer1.append(matchtable[i][j])
                layer.append(layer1)
                #  print(layer)
                for k in range(len(paperkeyword)):
                    if sequencesuperset(sequence_list[paperkeyword[k]], layer) > 0:
                        sumofmatchweight = sumofmatchweight + matchweight[paperkeyword[k]]
                if sumofmatchweight > threshold:
                    copylayer = copy.deepcopy(layer)
                    templayerTable.append(copylayer)
                    tempkeywordweight.append(sumofmatchweight)
                    #  print(copylayer, sumofmatchweight)
                sumofmatchweight = 0
                layer1 = []
                layer.pop(1)
            if not templayerTable:
                break
            layerTable = copy.deepcopy(templayerTable)
            keywordweight = copy.deepcopy(tempkeywordweight)
            templayerTable = []
            tempkeywordweight = []
        else:
            for j in range(len(layerTable)):
                for m in range(len(matchtable[i])):
                    layer = copy.deepcopy(layerTable[j])
                    if layer[(x // 2 + 1)][0] < matchtable[i][m]:
                        layer = copy.deepcopy(layerTable[j])
                        layer[(x // 2 + 1)].append(matchtable[i][m])
                        # print(layer)
                        for k in range(len(paperkeyword)):
                            if sequencesuperset(sequence_list[paperkeyword[k]], layer) > 0:
                                sumofmatchweight = sumofmatchweight + matchweight[paperkeyword[k]]
                        if sumofmatchweight > threshold:
                            copylayer = copy.deepcopy(layer)
                            layerTable.append(copylayer)
                            keywordweight.append(sumofmatchweight)
                            #  print(copylayer, sumofmatchweight)
                        sumofmatchweight = 0

    for x in range(len(layerTable)):
        if sequencelength(layerTable[x]) > maxsequencelength:
            maxsequencelength = sequencelength(layerTable[x])
            maxsequencesupport = keywordweight[x]
            layer = copy.deepcopy(layerTable[x])
        elif sequencelength(layerTable[x]) == maxsequencelength:
            if keywordweight[x] > maxsequencesupport:
                layer = copy.deepcopy(layerTable[x])
                maxsequencesupport = keywordweight[x]
    if maxsequencelength > 0:
        for x in range(len(layer)):
            for y in range(len(layer[x])):
                writeKeyword.cells(keywordorder + 1, sequenceorder).value = layer[x][y]
                sequenceorder = sequenceorder + 1
        keywordorder = keywordorder + 1
        print(layer)

    maxsequencelength = 0
    maxsequencesupport = 0
    sequenceorder = 1
    layer0.pop()
    layer = []
    layerTable = []
    layerTable1 = []
    keywordweight = []
    paperkeyword = []
