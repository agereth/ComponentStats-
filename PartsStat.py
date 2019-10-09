import os
import csv
import xlwt

#positions to exclude from analysis
dummies = ["", "*", "Logo", "Fuse0R", "Logo", "TESTPOINT", "REFPOINT", "HOLE_METALLED", "DoNotMount", "BUTTON", "SWITCH", "ANT_PCB_MONO_2PIN",
           "ANT_GENERAL"]
components = {}
sinonims = [['BLM15AG102SN1', 'BLM15', 'BLM15AG102']]


def get_safe_data(BOMfilename:str)->(int, list):
    """
    opens file BOMfilename as csv and returns data (exclude headers) and value column index as any
    if decode error or no value column returns None
    :param BOMfilename: filename
    :return: value column index, footprint column index, data
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
            return headers_corrected.index('value'), headers_corrected.index('footprint'), data[1:]
        except ValueError:
            print('No value or footprint column in %s BOM' % BOMfilename)
            return None, None, None


def write_data_to_xls(result: list):
    """
    writes data to xls file
    :param result: data
    :return:
    """
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Statistics")
    sheet1.write(0, 0, "Value")
    sheet1.write(0, 1, "Type")
    sheet1.write(0, 2, "Footprint")
    sheet1.write(0, 3, "Rate")
    sheet1.write(0, 4, "Projects")
    i = 1
    for value, count, projects in result:
        try:
            part = ""
            footprint = ""
            newvalue = value
            if any([part in value for part in ("CAP", "RES")]):
                newvalue = value.split(' ')[0]
                part = value.split(' ')[-1].split('_')[0]
                footprint = value.split(' ')[-1].split('_')[1]
            sheet1.write(i, 0, newvalue)
            sheet1.write(i, 1, part)
            sheet1.write(i, 2, footprint)
            sheet1.write(i, 3, count)
            sheet1.write(i, 4, ' '.join(projects))
            i += 1
        except Exception:
            pass
    try:
        book.save("Components Statistics.xls")
    except PermissionError:
        print("File Components Statistics.xls is already opened")


def main():
    for BOMfilename in os.listdir("."):
        if BOMfilename.split('.')[1] == 'csv':
            i_value, i_footprint, data = get_safe_data(BOMfilename)
            if data:
                for row in data:
                    value = row[i_value]
                    footprint = row[i_footprint]
                    if any([part in footprint for part in ['Resistor', 'Capacitor']]):
                         value = value +' ' + footprint.split(':')[1]
                    if value in components.keys():
                        components[value].append(BOMfilename.split('BOM')[0])
                    else:
                        components[value] = [BOMfilename.split('BOM')[0]]
    data = [item for item in components.items() if not item[0].split(' ')[0] in dummies]
    result = [(value, len(list(set(projects))), list(set(projects))) for (value, projects) in data]
    result.sort(key=lambda x:x[1], reverse=True)
    write_data_to_xls(result)


if __name__ == '__main__':
    main()