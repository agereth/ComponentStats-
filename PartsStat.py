import os
import csv
import xlwt

dummies = ["", "*", "Logo", "Fuse0R", "Logo", "TESTPOINT", "REFPOINT", "HOLE_METALLED", "DoNotMount"]
components = {}


for file in os.listdir("."):
    if file.split('.')[1] == 'csv':
        f = open(file)
        data = csv.reader(f)
        try:
            data = list(data)
        except UnicodeDecodeError:
            continue
        if 'Value' in data[0]:
            i_value = data[0].index('Value')
        else:
            if ' Value' in data[0]:
                i_value = data[0].index(' Value')
            else:
                print("No value column for BOM %s" % file)
                continue
        for row in data[1:]:
            value = row[i_value]
            if value in components.keys():
                components[value][0]+=1
                components[value][1]+=', %s' % file.split('BOM')[0]
            else:
                components[value] = [1, file.split('BOM')[0]]
data = [(value, count, projects) for (value, (count, projects)) in components.items()]
data = [item for item in data if not item[0] in dummies]
result = []
for value, count, projects in data:
    unique_projects = set(projects.split(', '))
    count -= (len(projects.split(', ')) - len(unique_projects))
    projects = ', '.join(unique_projects)
    result.append((value, count, projects))
result.sort(key=lambda x:x[1], reverse=True)

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Statistics")
sheet1.write(0, 0, "Value")
sheet1.write(0, 1, "Rate")
sheet1.write(0, 2, "Projects")
i = 1
for value, count, projects in result:
    sheet1.write(i, 0, value)
    sheet1.write(i, 1, count)
    sheet1.write(i, 2, projects)
    i+=1
book.save("Components Statistics.xls")

