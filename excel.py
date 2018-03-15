import xlrd
import xlwt
import xdrlib
import sys
import xlutils.copy


if __name__ == "__main__":
    map_dic = {}

    with open('compare.tsv', 'r', encoding='utf-8') as out_reader:
        for line in out_reader.readlines():
            pairs = line.split(',')
            if len(pairs) > 2:
                row_col = (pairs[1].strip(), pairs[2].strip())
                map_dic[pairs[0]] = row_col

    table_file = "D:/tmp/xuejia/template.xls"
    data = xlrd.open_workbook(table_file, formatting_info=True)
    first_sheet = data.sheets()[0]
    second_sheet = data.sheets()[1]

    outweb = xlutils.copy.copy(data)

    danwei = '填报单位：' + '' + '市' + '县（市、区）' + '' + '乡镇（单位）' + '' + '村（社区）'
    xingming = '姓名'
    huji = '户籍'
    id = '身份证'
    chengwei = '称谓'
    data = '出生年月'

    row_num = int(map_dic["填报单位"][0])
    col_num = int(map_dic["填报单位"][1])

    first_sheet.put_cell(int(row_num), int(col_num), 1, danwei, xf_index=0)
    outweb.get_sheet(0).write(row_num, col_num, danwei)

    row_num = int(map_dic["姓名"][0])
    col_num = int(map_dic["姓名"][1])
    first_sheet.put_cell(int(row_num), int(col_num), 1, xingming, xf_index=0)
    outweb.get_sheet(0).write(row_num, col_num, xingming)
    # data.save("./new.xls")
    outweb.save('output2.xls')

    wbook = xlwt.Workbook()
    wsheet = wbook.add_sheet(first_sheet.name)
    style = xlwt.easyxf('align: vertical center, horizontal center')

    for r in range(first_sheet.nrows):
        for c in range(first_sheet.ncols):
            wsheet.write(r, c, first_sheet.cell_value(r, c))

    wbook.save('./output.xls')

    # for i in range(first_sheet.nrows):
    #    for j in range(first_sheet.ncols):
    #        print('r:', i, '\tc:', j, first_sheet.cell(i,j).value)

