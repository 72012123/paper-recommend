import xlwings as xw
import numpy as np
import copy
import math

wa = xw.Book("frequentitem.xlsx")
wb = xw.Book("citation.xlsx")
wc = xw.Book("keyword.xlsx")
wd = xw.Book("sequencial.xlsx")
we = xw.Book("citationthreshold.xlsx")
wf = xw.Book()

read_item = wa.sheets[0]
read_citation = wb.sheets[0]
read_keyword = wc.sheets[0]
read_sequence = wd.sheets[0]
read_threshold = we.sheets[0]
writeKeyword = wf.sheets['工作表1']#writeKeyword.cells(x, y).value

dataSize = 156384
keywordSize = 20343
citation = np.zeros((2, dataSize))
keyword = np.zeros(keywordSize)
sequence = np.zeros((40, dataSize))
sequence_list = [[0 for y in range(8)] for x in range(dataSize)]
matchweight = [0 for x in range(dataSize)]
matchtable = [[0 for y in range(0)] for x in range(keywordSize)]  # 每個keyword配對成功的
matchnumber = [0 for x in range(keywordSize)]  # 每個keyword組合符合的數量
threshold = [0 for x in range(keywordSize)]
maxlayer = []
layer = []
layer0 = []
layer1 = []
layercompare = []  # 用來確認sequence是否包含keyword
layerTable = []
templayerTable = []
keywordweight = []  # 每個sequence的權重
tempkeywordweight = []
paperkeyword = []  # 跟此keyword相關的author在陣列的位置
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


itemset = read_item.range('a1:a20343').value
itemset = [round(x) for x in itemset]
citation = read_citation.range('a1:b156384').value
citation = [[round(x) for x in y] for y in citation]
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

threshold = read_threshold.range('b1:b20343').value
for i in range(keywordSize):
    threshold[i] = threshold[i] * 0.47

for i in range(dataSize):
    if citation[i][1] > 0:
        matchweight[i] = math.log(citation[i][1])

for i in range(keywordSize):  # 找出單個keyword位置keywordSize
    layer0.append(keyword[i])
    matchtable[i].append(keyword[i])
    itemposition = finditemposition(keyword[i])
    layer.append(layer0)
    layercompare.append(layer0)
    print(layer0, threshold[i])
    for j in range(dataSize):  # dataSize
        findkeyword = sequencesuperset(sequence_list[j], layer)
        if findkeyword > 0:
            paperkeyword.append(j)
            for k in range(len(sequence_list[j])):
                for x in sequence_list[j][k]:
                    # 把x出現的次數記錄在matchnumber
                    for l in range(keywordSize):  # keywordSize
                        if x == keyword[l]:
                            matchnumber[l] = matchnumber[l] + matchweight[j]
    for j in range(keywordSize):  # keywordSize
        if matchnumber[j] > threshold[i] and finditemposition(keyword[j]) < itemposition:
            matchtable[i].append(keyword[j])
        matchnumber[j] = 0
    print(len(matchtable[i]), len(paperkeyword))
    layer = []
    for x in range(5):
        if (x % 2) == 0:
            if not layerTable:
                for j in range(len(matchtable[i]) - 1):
                    for m in range(j + 1, len(matchtable[i])):
                        layer1.append(matchtable[i][j])
                        layer1.append(matchtable[i][m])
                        layer.append(layer1)
                        # print(layer)
                        for k in range(len(paperkeyword)):
                            if sequencesuperset(sequence_list[paperkeyword[k]], layer) > 0:
                                sumofmatchweight = sumofmatchweight + matchweight[paperkeyword[k]]
                        if sumofmatchweight > threshold[i]:
                            copylayer = copy.deepcopy(layer)
                            layerTable.append(copylayer)
                            keywordweight.append(sumofmatchweight)
                            # print(copylayer, sumofmatchweight)
                            if sequencelength(copylayer) > maxsequencelength and sequencesuperset(copylayer, layercompare) > 0:
                                maxsequencelength = sequencelength(copylayer)
                                maxsequencesupport = sumofmatchweight
                                maxlayer = copy.deepcopy(copylayer)
                            elif sequencelength(copylayer) == maxsequencelength and sequencesuperset(copylayer, layercompare) > 0:
                                if sumofmatchweight > maxsequencesupport:
                                    maxlayer = copy.deepcopy(copylayer)
                                    maxsequencesupport = sumofmatchweight
                        sumofmatchweight = 0
                        layer1 = []
                        layer.pop()
            else:
                for j in range(len(layerTable)):
                    for m in range(len(matchtable[i])):
                        layer = copy.deepcopy(layerTable[j])
                        if layer[(x // 2)][0] < matchtable[i][m]:
                            layer = copy.deepcopy(layerTable[j])
                            layer[(x // 2)].append(matchtable[i][m])
                            # print(layer)
                            for k in range(len(paperkeyword)):
                                if sequencesuperset(sequence_list[paperkeyword[k]], layer) > 0:
                                    sumofmatchweight = sumofmatchweight + matchweight[paperkeyword[k]]
                            if sumofmatchweight > threshold[i]:
                                copylayer = copy.deepcopy(layer)
                                templayerTable.append(copylayer)
                                tempkeywordweight.append(sumofmatchweight)
                                print(copylayer, sumofmatchweight)
                                if sequencelength(copylayer) > maxsequencelength and sequencesuperset(copylayer, layercompare) > 0:
                                    maxsequencelength = sequencelength(copylayer)
                                    maxsequencesupport = sumofmatchweight
                                    maxlayer = copy.deepcopy(copylayer)
                                elif sequencelength(copylayer) == maxsequencelength and sequencesuperset(copylayer, layercompare) > 0:
                                    if sumofmatchweight > maxsequencesupport:
                                        maxlayer = copy.deepcopy(copylayer)
                                        maxsequencesupport = sumofmatchweight
                            sumofmatchweight = 0
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
                    layer1.append(matchtable[i][m])
                    layer.append(layer1)
                    # print(layer)
                    for k in range(len(paperkeyword)):
                        if sequencesuperset(sequence_list[paperkeyword[k]], layer) > 0:
                            sumofmatchweight = sumofmatchweight + matchweight[paperkeyword[k]]
                        if sumofmatchweight > threshold[i]:
                            copylayer = copy.deepcopy(layer)
                            templayerTable.append(copylayer)
                            tempkeywordweight.append(sumofmatchweight)
                            print(copylayer, sumofmatchweight)
                            if sequencelength(copylayer) > maxsequencelength and sequencesuperset(copylayer, layercompare) > 0:
                                maxsequencelength = sequencelength(copylayer)
                                maxsequencesupport = sumofmatchweight
                                maxlayer = copy.deepcopy(copylayer)
                            elif sequencelength(copylayer) == maxsequencelength and sequencesuperset(copylayer, layercompare) > 0:
                                if sumofmatchweight > maxsequencesupport:
                                    maxlayer = copy.deepcopy(copylayer)
                                    maxsequencesupport = sumofmatchweight
                        sumofmatchweight = 0
                    layer1 = []
            if not templayerTable:
                break
            layerTable = copy.deepcopy(templayerTable)
            keywordweight = copy.deepcopy(tempkeywordweight)
            templayerTable = []
            tempkeywordweight = []

    if maxsequencelength > 0:
        writeKeyword.cells(keywordorder + 1, 1).value = keyword[i]
        for x in range(len(maxlayer)):
            for y in range(len(maxlayer[x])):
                writeKeyword.cells(keywordorder + 1, sequenceorder + 1).value = maxlayer[x][y]
                sequenceorder = sequenceorder + 1
        for x in range(6 - sequencelength(maxlayer)):
            writeKeyword.cells(keywordorder + 1, sequenceorder + x + 1).value = 0
        keywordorder = keywordorder + 1
        print(maxlayer)
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
