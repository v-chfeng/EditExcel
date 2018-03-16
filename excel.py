import xlrd
import xlwt
import sys
import xlutils.copy


def convert_num(alpha):
    if len(alpha) == 1:
        return ord(alpha) - ord('a')
    else:
        return (ord(alpha[0]) - ord('a') + 1) * 26 + ord(alpha[1]) - ord('a')


if __name__ == "__main__":
    first_map_dic = {}
    second_map_dic = {}
    input_map_dic = {}

    with open('compare.tsv', 'r', encoding='utf-8') as out_reader:
        for line in out_reader.readlines():
            page_pairs = line.split(' ')
            if page_pairs[0] == '1':
                pairs = page_pairs[1].strip().split(',')
                if len(pairs) > 2:
                    row_col = pairs[1:]
                    first_map_dic[pairs[0]] = row_col
            else:
                pairs = page_pairs[1].strip().split(',')
                if len(pairs) > 2:
                    row_col = pairs[1:]
                    second_map_dic[pairs[0]] = row_col

    with open('transform.tsv', 'r', encoding='utf-8') as input_reader:
        for line in input_reader.readlines():
            input_pairs = line.split(' ')
            input_map_dic[input_pairs[0]] = input_pairs[1].strip()

    table_file = "template.xls"
    input_file = 'input.xlsx'
    out_path = './outdata/'

    input_data = xlrd.open_workbook(input_file)
    input_sheet = input_data.sheets()[0]
    for input_r in range(input_sheet.nrows):
        out_table = xlrd.open_workbook(table_file, formatting_info=True)
        out_copy = xlutils.copy.copy(out_table)
        table_name = ''
        if input_r > 1:
            for name, first_value in first_map_dic.items():
                # name = first_pair.key
                rc_pairs = first_value
                input_rc = input_map_dic[name]
                input_cols = input_rc.split(',')
                if len(rc_pairs) == 2:
                    if name == '填报单位':
                        tianbiao = '填报单位：{0}市{1}县（市、区）{2}乡镇（单位）{3}村（社区）'
                        shi = input_sheet.cell(input_r, convert_num(input_cols[0])).value
                        xian = input_sheet.cell(input_r, convert_num(input_cols[1])).value
                        xiang = input_sheet.cell(input_r, convert_num(input_cols[2])).value
                        cun = input_sheet.cell(input_r, convert_num(input_cols[3])).value
                        input_value = tianbiao.format(shi,xian, xiang, cun)
                    elif name == '家庭收入':
                        shouru = '家庭收入：{0}  生活情况：{1}'
                        shouruinfo = input_sheet.cell(input_r, convert_num(input_cols[0])).value
                        shenghuo = input_sheet.cell(input_r, convert_num(input_cols[1])).value
                        input_value = shouru.format(shouruinfo, shenghuo)
                    elif name == '现户籍所在地':
                        huji1 = input_sheet.cell(input_r, convert_num(input_cols[0])).value
                        huji2 = input_sheet.cell(input_r, convert_num(input_cols[1])).value
                        huji3 = input_sheet.cell(input_r, convert_num(input_cols[2])).value
                        huji4 = input_sheet.cell(input_r, convert_num(input_cols[3])).value
                        input_value = str(huji1) + str(huji2) + str(huji3) + str(huji4)
                    elif name == '家庭住址':
                        zhuzhi1 = input_sheet.cell(input_r, convert_num(input_cols[0])).value
                        zhuzhi2 = input_sheet.cell(input_r, convert_num(input_cols[1])).value
                        zhuzhi3 = input_sheet.cell(input_r, convert_num(input_cols[2])).value
                        zhuzhi4 = input_sheet.cell(input_r, convert_num(input_cols[3])).value
                        input_value = str(zhuzhi1) + str(zhuzhi2) + str(zhuzhi3) + str(zhuzhi4)
                    elif name == "姓名":
                        input_value = input_sheet.cell(input_r, convert_num(input_cols[0])).value
                        table_name = input_value
                    else:
                        input_value = input_sheet.cell(input_r, convert_num(input_cols[0])).value

                    out_copy.get_sheet(0).write(int(rc_pairs[0]), int(rc_pairs[1]), input_value)
                else:
                    if name == '家庭成员':
                        for turn in range(len(input_cols) // 3):
                            jiatingxingming = input_sheet.cell(input_r, convert_num(input_cols[turn*3 + 0])).value
                            chengwei = input_sheet.cell(input_r, convert_num(input_cols[turn*3+1])).value
                            jiatinggongzuo = input_sheet.cell(input_r,convert_num(input_cols[turn*3+2])).value
                            out_copy.get_sheet(0).write(int(rc_pairs[0])+turn, int(rc_pairs[1]), jiatingxingming)
                            out_copy.get_sheet(0).write(int(rc_pairs[2])+turn, int(rc_pairs[3]), chengwei)
                            out_copy.get_sheet(0).write(int(rc_pairs[4])+turn, int(rc_pairs[5]), jiatinggongzuo)
            for name, second_value in second_map_dic.items():
                # name = second_pair.key
                out_rc = second_value
                input_rc = input_map_dic[name]
                input_cols = input_rc.strip(',').split(',')
                input_value = input_sheet.cell(input_r, convert_num(input_cols[0])).value
                out_copy.get_sheet(1).write(int(out_rc[0]), int(out_rc[1]), input_value,)
            out_copy.save(out_path + table_name + str(input_r - 1) + '.xls')
            # out_table.close()
            # out_copy.close()