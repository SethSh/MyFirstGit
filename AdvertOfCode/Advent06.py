from collections import Counter
import copy
import numpy as np

fileName = 'Day6Data.txt'
dataAsString = np.loadtxt(fileName, delimiter=',', skiprows=0, dtype=str).tolist()

row_count = len(dataAsString)
length = len(dataAsString[0])
print(f"rows {row_count} length {length}")

data = [int(i) for i in dataAsString]
dataDict = Counter(data)
for i in range(9):
    if i not in dataDict:
        dataDict[i] = 0
# print(dataDict)

days = 256
for _ in range(days):
    newDataDict = copy.deepcopy(dataDict)
    for i in range(8):
        newDataDict[i] = dataDict[i+1]
    newDataDict[8] = dataDict[0]
    newDataDict[6] += dataDict[0]
    
    dataDict = newDataDict
    print(dataDict)

print(sum(dataDict.values()))


# for _ in range(days):
#     newData = [i-1 for i in data]
#     isNegativeOne = [x for x in newData if x == -1]
#     if any(isNegativeOne):
#         count = len(isNegativeOne)
#         newEights = [8 for i in range(count)]
#         newData.extend(newEights)
#         newData = [6 if i == -1 else i for i in newData]
#     data = newData
# print (len(data))
    