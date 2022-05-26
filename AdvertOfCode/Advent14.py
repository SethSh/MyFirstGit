import numpy as np
    
fileName = 'Day14Data.txt'
dataAsString = np.loadtxt(fileName, delimiter=',', skiprows=0, dtype=str).tolist()

row_count = len(dataAsString)
full_length = len(dataAsString[0])
print(f"rows {row_count} length {full_length}")

initial_code = dataAsString[0]
mapping_list = dataAsString[1:]


dict = {}
for item in mapping_list:
    dict[item[:2]] = item[-1]

adjusted_code = initial_code
for ii in range(0,10):
    counter = 0
    for jj in range(len(adjusted_code)-2):
        adjusted_counter = jj + counter
        two_letters = adjusted_code[adjusted_counter:(adjusted_counter + 2)]
        if two_letters in dict:
            full_length = len(adjusted_code)
            delta = full_length - adjusted_counter - 1
            adjusted_code = adjusted_code[:adjusted_counter+1] + dict[two_letters] + adjusted_code[full_length - delta:]
            
            counter+=1

uniques = set(adjusted_code)
myMinimum = 10000
myMaximum = 0
for unique in uniques:
    count = adjusted_code.count(unique)
    if count < myMinimum:
        myMinimum=count
    if count > myMaximum:
        myMaximum=count
print(myMaximum -myMinimum)




