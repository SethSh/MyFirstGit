from collections import Counter
import sys
import numpy as np
import statistics
from functools import cache

@cache
def psuedo_factorial(n):
    return n * (n+1) // 2

fileName = 'Day7Data.txt'
dataAsString = np.loadtxt(fileName, delimiter=',', skiprows=0, dtype=str).tolist()

row_count = len(dataAsString)
length = len(dataAsString[0])
print(f"rows {row_count} length {length}")

data = [int(i) for i in dataAsString]
print(data[:5])

dataMedian = statistics.median(data)
distance = [abs(i - dataMedian) for i in data]
print(distance[:5])
print(sum(distance))

print(f"min {min(data)} max {max(data)}")


myDistance = sys.maxsize
for i in range(min(data), max(data)+1):
    totalDistance = sum(psuedo_factorial(abs(item - i)) for item in data)
    if totalDistance < myDistance:
        myDistance = totalDistance

print(myDistance)
# print(sum(distance))