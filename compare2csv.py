import csv
t1 = open('Test 4D29 Step 1.csv', 'r')
t2 = open('Test 4D29 Step 1 generated.csv', 'r')
fileone = t1.readlines()
filetwo = t2.readlines()
t1.close()
t2.close()

outFile = open('update.csv', 'w')
x = 0
for i in fileone:
    if i.strip() != filetwo[x].strip():
        outFile.write(filetwo[x])
    x +=1
outFile.close()