import numpy as np

fileName = 'Day5Data.txt'
data = np.loadtxt(fileName, delimiter='->', skiprows=0, dtype=str).tolist()

row_count = len(data)
length = len(data[0])
print(f"rows {row_count} length {length}")

counter = 0
for row_item in data:
    res = [i.split(',') for i in row_item]
    if counter == 0:
        first = [int(i) for i in res[0]]
        x1, y1 = first
        print(f"x1 {x1} and y1 {y1}") 

        second = [int(i) for i in res[1]]
        x2, y2 = second
        print(f"x2 {x2} and y2 {y2}") 
        counter +=1
