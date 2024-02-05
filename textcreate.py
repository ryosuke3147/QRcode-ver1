f = open("./exam_name.txt", "r", encoding='UTF-8')
datalist = f.readlines()
for data in datalist:
    print(data)
f.close()