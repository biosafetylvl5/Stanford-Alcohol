from xlrd import open_workbook
from optparse import OptionParser
import json

parser = OptionParser()
parser.add_option("-f", "--excelfile", dest="filename",
                          help="Input Excel file", metavar="FILE")
(options, args) = parser.parse_args()

data = {}
wb = open_workbook(options.filename)
for s in wb.sheets():
    print("Processing:", s.name)
    data[s.name]={}
    i = 0
    group=0
    while True:
        group=group+1
        header = list(filter(lambda x: not x == "", map(lambda x: x.value, s.row(i))))
        data[s.name][group]={}
        data[s.name][group]["Header"] = header
        i=i+1
        while True:
            try:
                row = list(map(lambda x: x.value, s.row(i)))[0:len(header)]
                i = i+1
                if row == [""]*(len(header)):
                    break
                if row[0] not in data[s.name]:
                    data[s.name][group][row[0]]={}
                rowData = {}
                for x in range(2,len(header)):
                    rowData[header[x]] = row[x]
                if not row[1] == "":
                    data[s.name][group][row[0]][row[1]] = rowData
                else:
                    data[s.name][group][row[0]] = rowData
            except IndexError:
                break
        try:
            s.row(i)
        except IndexError:
            break
with open(options.filename.split(".")[0]+".json", "w") as outputFile:
    outputFile.write(json.dumps(data, indent=4))
