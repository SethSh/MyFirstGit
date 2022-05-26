import numpy as np
from functools import cache

@cache
def parse_digits(myInteger):
    return [int(a) for a in str(myInteger)]
    
fileName = 'Day9Data.txt'
dataAsString = np.loadtxt(fileName, delimiter=',', skiprows=0, dtype=str).tolist()

row_count = len(dataAsString)
length = len(dataAsString[0])
print(f"rows {row_count} length {length}")

data = [int(i) for i in dataAsString]
# print(data[:2])

max_data = 9
max_row = [max_data for ii in range(length)]



lowestList = []
for i, v in enumerate(data):
    
    above_row = max_row if i == 0 else parse_digits(data[i-1])
    below_row = max_row if i == row_count - 1 else parse_digits(data[i+1])
    
    digits = [int(a) for a in str(v)]
    for j, v2 in enumerate(digits):
        value_left = max_data if j == 0 else digits[j-1]
        value_right = max_data if j == len(digits)-1 else digits[j+1]
        value_above = above_row[j]
        value_below = below_row[j]

        if v2 < value_right and v2 < value_left and v2 < value_above and v2 < value_below:
            lowestList.append(v2)


lowestlistPlusOne = [x+1 for x in lowestList]
print (lowestlistPlusOne)
print(sum(lowestlistPlusOne))    


