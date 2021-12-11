import numpy as np

filename = 'Day5Data.txt'
data = np.loadtxt(filename, delimiter='->', skiprows=0, dtype=str).tolist()

row_count = len(data)
length = len(data[0])
print(f"rows {row_count} length {length}")

for row_item in data:
    res = [int(i.strip().replace(',','')) for i in row_item]
    print(res)
# print(data[0][0])
# print(data[0][1])
# print(type(data[0]))