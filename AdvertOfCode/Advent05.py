import numpy as np

fileName = 'Day5Data.txt'
data = np.loadtxt(fileName, delimiter='->', skiprows=0, dtype=str).tolist()

row_count = len(data)
length = len(data[0])
print(f"rows {row_count} length {length}")

dimensionSize = 1000
countArray  = np.zeros(dimensionSize * dimensionSize)

countArray = []
for i in range(0,dimensionSize):
    countArray.append(np.zeros(dimensionSize))

counter = 0
for row_item in data:
    res = [i.split(',') for i in row_item]
    
    first = [int(i) for i in res[0]]
    x1, y1 = first
    
    second = [int(i) for i in res[1]]
    x2, y2 = second
    #print(f"x1 {x1} and y1 {y1} AND x2 {x2} and y2 {y2}")     

    if x1 == x2:
        miniumumValue = min(y1, y2)
        maximumValue = max(y1, y2)
        for i in range(miniumumValue, maximumValue+1):
            countArray[x1][i] += 1
    else:
        miniumumValue = min(x1, x2)
        maximumValue = max(x1, x2)
        for i in range(miniumumValue, maximumValue+1):
            countArray[i][y1] += 1

total = 0
for i in range(0,dimensionSize):
    total += sum(countArray[i] >= 2)
print(total)


