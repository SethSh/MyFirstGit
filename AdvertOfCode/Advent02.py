import numpy as np

filename = 'Day2Data.txt'
data = np.loadtxt(filename, delimiter=',', skiprows=0, dtype=str).tolist()

progress = 0
depthProgress = 0
for current in range(len(data)):
    item = data[current]
    parts = item.split(' ')
    
    direction = parts[0]
    magnitude = int(parts[1])
    if direction == 'forward':
        progress +=magnitude
    else:
        if direction == 'down':
            depthProgress +=magnitude
        elif direction == 'up':
            depthProgress -=magnitude
    
print(progress)
print(depthProgress)
print (progress * depthProgress)
    
#part 2
progress = 0
depthProgress = 0
aim = 0
for current in range(len(data)):
    item = data[current]
    parts = item.split(' ')
    
    direction = parts[0]
    magnitude = int(parts[1])
    if direction == 'forward':
        progress +=magnitude
        depthProgress += aim * magnitude
    else:
        if direction == 'down':
            aim +=magnitude
        elif direction == 'up':
            aim -=magnitude
    
print(progress)
print(depthProgress)
print (progress * depthProgress)
