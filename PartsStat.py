import os
import csv
import xlwt

#positions to exclude from analysis
dummies = ["", "*", "Logo", "Fuse0R", "Logo", "TESTPOINT", "REFPOINT", "HOLE_METALLED", "DoNotMount"]
components = {}


def get_safe_data(BOMfilename:str)->(int, list):
    """
    opens file BOMfilename as csv and returns data (exclude headers) and value column index as any
    if dercode error or no value column returns None
    :param BOMfilename: filename
    :return: value column index, data
    """
    with open(BOMfilename, encoding='utf-8') as f:
        data = csv.reader(f)
        try:
            data = list(data)
        except UnicodeDecodeError:
            print("Decode error in %s" % BOMfilename)
            return None, None
        csv_headers = data[0]
        headers_corrected = [header.strip().lower() for header in csv_headers]
        try:
            return headers_corrected.index('value'), data[1:]
        except IndexError:
            print('No value column in %s BOM' % BOMfilename)
            return None, None


def write_data_to_xls(result: list):
    """
    writes data to xls file
    :param result: data
    :return:
    """
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Statistics")
    sheet1.write(0, 0, "Value")
    sheet1.write(0, 1, "Rate")
    sheet1.write(0, 2, "Projects")
    i = 1
    for value, count, projects in result:
        sheet1.write(i, 0, value)
        sheet1.write(i, 1, count)
        sheet1.write(i, 2, ' '.join(projects))
        i += 1
    book.save("Components Statistics.xls")


def main():
    for BOMfilename in os.listdir("."):
        if BOMfilename.split('.')[1] == 'csv':
            i_value, data = get_safe_data(BOMfilename)
            if data:
                for row in data:
                    value = row[i_value]
                    if value in components.keys():
                        components[value].append(BOMfilename.split('BOM')[0])
                    else:
                        components[value] = [BOMfilename.split('BOM')[0]]
    data = [item for item in components.items() if not item[0] in dummies]
    result = [(value, len(list(set(projects))), list(set(projects))) for (value, projects) in data]
    result.sort(key=lambda x:x[1], reverse=True)
    write_data_to_xls(result)


if __name__ == '__main__':
    main()