import numpy as np

filename = 'Day3Data.txt'
data = np.loadtxt(filename, delimiter=',', skiprows=0, dtype=str).tolist()

row_count = len(data)
length = len(data[0])
print(f"rows {row_count} length {length}")
    
one_counts = [0 for n in range(length)]
for row_item in data:
    
    j = 0
    for item in row_item:
        if item == '1':
            one_counts[j] +=1
        j += 1
print (one_counts)

gammaInBinary = ''
epsilonInBinary = ''
j = 0
for one_count in one_counts:
    if one_count > 500:
        gammaInBinary += '1'
        epsilonInBinary += '0'
    else:
        gammaInBinary += '0'
        epsilonInBinary += '1'
    
    j += 1



gamma = int(gammaInBinary, 2)
epsilon = int(epsilonInBinary, 2)
print(gamma)
print(epsilon)
print(gamma * epsilon)