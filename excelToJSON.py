from xlrd import open_workbook
import json

data = {}

wb = open_workbook('raw_released_data.xlsx')
for s in wb.sheets():
    print(s.name)
    data[s.name]={}
    i = 0
    while True:
        header = list(filter(lambda x: not x == "", map(lambda x: x.value, s.row(i))))
        print("")
        print("HEADER", header)
        print("")
        i=i+1
        while True:
            try:
                row = list(map(lambda x: x.value, s.row(i)))[0:len(header)]
                i = i+1
                if row == [""]*(len(header)):
                    break
                print(row)
                if row[0] not in data[s.name]:
                    data[s.name][row[0]]={}
                rowData = {}
                for x in range(2,len(header)):
                    rowData[header[x]] = row[x]
                if not row[1] == "":
                    data[s.name][row[0]][row[1]] = rowData
                else:
                    data[s.name][row[0]] = rowData
            except IndexError:
                break
        try:
            s.row(i)
        except IndexError:
            break
with open("data.json", "w") as outputFile:
    json.dump(data, outputFile)
