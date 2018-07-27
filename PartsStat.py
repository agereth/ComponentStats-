import os
import csv
import xlwt

#positions to exclude from analysis
dummies = ["", "*", "Logo", "Fuse0R", "Logo", "TESTPOINT", "REFPOINT", "HOLE_METALLED", "DoNotMount"]
components = {}


def main():
    for BOMfilename in os.listdir("."):
        if BOMfilename.split('.')[1] == 'csv':
            with open(BOMfilename, encoding='utf-8') as f:
                data = csv.reader(f)
                try:
                    data = list(data)
                except UnicodeDecodeError:
                    print("Decode error in %s" % BOMfilename)
                    continue
                csv_headers = data[0]
                headers_corrected = [header.strip().lower() for header in csv_headers]
                try:
                    i_value = headers_corrected.index('value')
                except IndexError:
                    print('No value column in %s BOM' % BOMfilename)
                    continue
                for row in data[1:]:
                    value = row[i_value]
                    if value in components.keys():
                        components[value].append(BOMfilename.split('BOM')[0])
                    else:
                        components[value] = [BOMfilename.split('BOM')[0]]

    data = [item for item in components.items() if not item[0] in dummies]
    result = [(value, len(list(set(projects))), list(set(projects))) for (value, projects) in data]
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

if __name__ == '__main__':
    main()