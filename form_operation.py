import xlrd
import xlwt
import os

excel_data = []
findData = []
queryIndex = 0


def my_workbook():
    return xlwt.Workbook()


def _my_add_sheet(obj_workbook, sheet_name):
    sheet1 = obj_workbook.add_sheet(sheet_name)
    sheet1.write(0, 0, '姓名')
    sheet1.write(0, 1, '车牌号码')
    sheet1.write(0, 2, '年检到期月份')
    sheet1.write(0, 3, '保险到期日期')
    sheet1.write(0, 4, '投保保险公司')
    sheet1.write(0, 5, '联系电话')
    return sheet1


def add_content(obj_workbook, sheet_name, content_dict_list):
    sheet1 = _my_add_sheet(obj_workbook, sheet_name)
    for i, info in enumerate(content_dict_list):
        info1 = info['姓名']
        info2 = info['车牌号码']
        info3 = info['年检到期月份']
        info4 = info['保险到期日期']
        info5 = info['投保保险公司']
        info6 = info['联系电话']
        sheet1.write(1 + i, 0, info1)
        sheet1.write(1 + i, 1, info2)
        sheet1.write(1 + i, 2, info3)
        sheet1.write(1 + i, 3, info4)
        sheet1.write(1 + i, 4, info5)
        sheet1.write(1 + i, 5, info6)


def create_excel(output_dir, case_name, content_dict_list):
    xls = my_workbook()
    add_content(xls, 'Sheet1', content_dict_list)
    file_path = os.path.join(output_dir, case_name)
    xls.save(file_path)
    return file_path


def open_excel(file_path):
    try:
        workbook = xlrd.open_workbook(file_path, encoding_override='utf-8')
        return workbook
    except Exception as e:
        print(str(e))


filepath1 = 'C:\\vnv-at\\query\\金南检测客户信息登记表.xls'
filepath2 = 'C:\\vnv-at\\query\\2019.xls'
workbook1 = open_excel(filepath1)
# workbook2 = open_excel(filepath2)


def query_register_excel_data(workbook, clue_data):
    """clue_data: list  [2,20210304]"""
    global excel_data
    target_data_dict_list = []
    target_data_dict = {}
    if not excel_data:
        # workbook = xlrd.open_workbook(file_path, encoding_override='utf-8')
        sheet_names = workbook.sheet_names()
        t_table = workbook.sheet_by_index(0)
        title1 = t_table.cell(0, 0).value
        title2 = t_table.cell(0, 1).value
        title3 = t_table.cell(0, 2).value
        title4 = t_table.cell(0, 3).value
        title5 = t_table.cell(0, 4).value
        title6 = t_table.cell(0, 5).value
        target_data_dict[title1] = ''
        target_data_dict[title2] = ''
        target_data_dict[title3] = ''
        target_data_dict[title4] = ''
        target_data_dict[title5] = ''
        target_data_dict[title6] = ''
        target_data_dict['sheet'] = ''
        target_data_dict['row'] = ''
        for sheet in sheet_names:
            table = workbook.sheet_by_name(sheet)
            # ncols = table.ncols
            nrows = table.nrows
            list_member = []
            list_member_group = []
            for n in range(1, nrows):
                if table.cell_value(int(n), 0) == '':
                    print('Invalid values,  The number of valid rows is: %s' % (int(n)))
                    break
                annual_check_expire_month = table.cell(n, 2).value
                if annual_check_expire_month == '':
                    annual_check_expire_month_change = ''
                else:
                    annual_check_expire_month_change = str(int(annual_check_expire_month))
                expiry_date_of_insurance = table.cell(n, 3).value
                if expiry_date_of_insurance == '':
                    expiry_date_of_insurance_change = ''
                else:
                    expiry_date_of_insurance_change = str(int(expiry_date_of_insurance))
                list_member.append(annual_check_expire_month_change)
                list_member.append(expiry_date_of_insurance_change)
                list_member_group.append(list_member)
                list_member = []
                # use target_data to find and output row
            for cnt, lmg in enumerate(list_member_group):
                if lmg == clue_data:
                    tar_row = cnt + 1  # list_member_group is not contain title, so  'cnt' need to add 1
                    target_data_dict[title1] = table.cell(tar_row, 0).value
                    target_data_dict[title2] = table.cell(tar_row, 1).value
                    target_data_dict[title3] = str(int(table.cell(tar_row, 2).value))
                    target_data_dict[title4] = str(int(table.cell(tar_row, 3).value))
                    target_data_dict[title5] = table.cell(tar_row, 4).value
                    target_data_dict[title6] = str(int(table.cell(tar_row, 5).value))
                    target_data_dict['sheet'] = sheet
                    target_data_dict['row'] = tar_row + 1  # excel cnt from 1
                    append_dict = target_data_dict.copy()
                    print('Catch the data: %s, in sheet %s' % (append_dict, sheet))
                    target_data_dict_list.append(append_dict)
        if target_data_dict_list == '':
            print('target data not in this excel, please check it!')
        else:
            return target_data_dict_list


def get_plate_number(sum_data):
    """sum_data: [dict1, dict2,....]"""
    cnt_data = len(sum_data)
    plate_number_lsit = []
    for sd in sum_data:
        plate_number_lsit.append(sd['车牌号码'])
    return plate_number_lsit


def get_plate_number_info(workbook):
    """return about the plate number all info"""
    info_dict = {}
    info_dict_list = []
    # workbook = xlrd.open_workbook(file_path, encoding_override='utf-8')
    sheet = workbook.sheet_by_index(0)
    title = sheet.cell(0, 0).value
    info_dict[title] = ''
    info_dict['row'] = ''
    for cnt, col in enumerate(sheet.col_values(0)):
        info_dict[title] = col
        info_dict['row'] = cnt
        append_info_dict = info_dict.copy()
        info_dict_list.append(append_info_dict)
    return info_dict_list


def get_specified_row(plate_number_list, workbook):
    """allow 'plate_number' multi, format is the list"""
    plate_number_info_dict_list = get_plate_number_info(workbook)
    target_row_list = []
    for pnl in plate_number_list:
        for i in plate_number_info_dict_list:
            if i['车牌号码'] == pnl:
                target_row_list.append(i['row'])
    return target_row_list


def get_specified_info(workbook, row_list):
    """support row list"""
    # workbook = xlrd.open_workbook(file_path, encoding_override='utf-8')
    sheet = workbook.sheet_by_index(0)
    specified_dict_list = []
    for row in row_list:
        specified_dict = {sheet.cell_value(0, 0): sheet.cell_value(row, 0),
                          sheet.cell_value(0, 1): sheet.cell_value(row, 1),
                          sheet.cell_value(0, 2): sheet.cell_value(row, 2),
                          sheet.cell_value(0, 3): sheet.cell_value(row, 3),
                          sheet.cell_value(0, 4): sheet.cell_value(row, 4),
                          sheet.cell_value(0, 5): sheet.cell_value(row, 5),
                          sheet.cell_value(0, 6): sheet.cell_value(row, 6),
                          sheet.cell_value(0, 7): sheet.cell_value(row, 7),
                          sheet.cell_value(0, 8): sheet.cell_value(row, 8),
                          sheet.cell_value(0, 9): sheet.cell_value(row, 9),
                          sheet.cell_value(0, 10): sheet.cell_value(row, 10),
                          sheet.cell_value(0, 11): sheet.cell_value(row, 11),
                          sheet.cell_value(0, 12): sheet.cell_value(row, 12),
                          }
        append_specified_dict = specified_dict.copy()
        specified_dict_list.append(append_specified_dict)
    return specified_dict_list


class StrForm(object):
    def __init__(self, list_dict_data):
        self.data = list_dict_data
        self.str_line = '-------------------------------------------------------------------------------------------'
        self.standard_line_len = 88
        self.str_vertical_line = '|'
        self.str_standard_blank = '  '
        self.str_title1 = '   姓名   '
        self.str_title1_len = 6         # 4 blanks, 2 chinese( total 8 blanks)
        self.str_title1_len_max = 10    # max 12 blank, 6 chinese
        self.str_title2 = '  车牌号码  '
        self.str_title2_len = 8
        self.str_title3 = '  年检到期月份  '
        self.str_title3_len = 10
        self.str_title4 = '  保险到期日期  '
        self.str_title4_len = 10
        self.str_title5 = '  投保保险公司  '
        self.str_title5_len = 10
        self.str_title6 = '  联系电话  '
        self.str_title6_len = 8
        self.cut_times = 0
        self.position_of_cut = []
        self.total_form_row_cnt = 0
        self.form_content_dict = {'row': 0, 'content': ''}
        self.form_content_info = []
        self.all_info = ''

    def make_title1_len(self):
        # get largest str length ( compare with self.str_tile1)
        for j, i in enumerate(self.data):
            name_value = i['姓名']
            str_blank_name_len = len(name_value)*2       # use blanks to compare
            if str_blank_name_len > self.str_title1_len+2:
                if str_blank_name_len > self.str_title1_len_max+2:
                    print('need cut to two lines')
                    cut_value = self.cut_line_max(name_value)
                    self.position_of_cut.append(j)
                    self.str_title1_len = self.str_title1_len_max+2
                    return
        self.str_title1_len = self.str_title1_len_max + 2
        #     else:
        #         print('no cut, but add self.str_title1 len')
        #         self.str_title1_len = str_blank_name_len
        # else:
        #     if self.str_title1_len != 6:
        #         print('no cut, but add self.str_title1 len')
        #     else:
        #         print('use standard title1 len: %s' % self.str_title1_len)

    def make_str_form_row_cnt(self):
        # simple explain: get .  how many newline characters are needed( due to cut row has add \n, so it not
        # include it.
        content_cnt = len(self.data)
        self.total_form_row_cnt = 3 + content_cnt*2  # + self.cut_times

    def assembly_title(self):
        # str_title1_standard = self.str_vertical_line + self.str_standard_blank + \
        #                      self.str_title1 + self.str_standard_blank
        # if len(str_title1_standard)-2 < self.str_title1_len:           # -2   | |
        #     print('need to change title1 length')
        #     str_title1_final = str_title1_standard + ' '*(self.str_title1_len-(len(str_title1_standard)+2))
        # else:
        #     str_title1_final = str_title1_standard
        name_title = '|    姓名    '
        other_title = '|  车牌号码  |  年检到期月份  |  保险到期日期  |  投保保险公司  |  联系电话  |'
        return name_title + other_title

    def assembly_content_cut(self):
        """deal with special row(cut)"""
        if not self.position_of_cut:
            print('No need to assembly cut content')
        else:
            for p in self.position_of_cut:
                cutline_data = self.data[p]
                cut1 = self.data[p]['姓名'][0:6]
                cut2 = self.data[p]['姓名'][6:]
                first_line = self.str_vertical_line + cut1 + self.str_vertical_line + self.str_standard_blank + \
                             cutline_data['车牌号码'] + self.str_standard_blank + self.str_vertical_line + \
                             self.put_value_in_blanks(16, cutline_data['年检到期月份']) + self.str_vertical_line + \
                             self.put_value_in_blanks(16, cutline_data['保险到期日期']) + self.str_vertical_line + ' '*16 + \
                             self.str_vertical_line + cutline_data['联系电话'] + ' ' + self.str_vertical_line
                second_line = self.str_vertical_line + cut2 + ' '*(12-len(cut2)*2) + \
                              '|            |                |                |                |            |' + \
                              '\n' + self.str_line
                join_line = first_line + '\n' + second_line
                self.form_content_dict['row'] = p
                self.form_content_dict['content'] = join_line
                append_dict = self.form_content_dict.copy()
                self.form_content_info.append(append_dict)

    def assembly_content_normal(self):
        all_row = []
        for i in range(len(self.data)):
            all_row.append(i)
        # find cut row, left normal row
        normal_row = [i for i in all_row if i not in self.position_of_cut]
        for n in normal_row:
            data_dict = self.data[n]
            name = data_dict['姓名']
            plate_number = data_dict['车牌号码']
            month = data_dict['年检到期月份']
            date = data_dict['保险到期日期']
            phone = data_dict['联系电话']
            left_blank_len = self.str_title1_len-2-len(name)*2
            content_line = self.str_vertical_line + self.str_standard_blank + name + left_blank_len*' ' + \
                           self.str_vertical_line + self.str_standard_blank + plate_number + self.str_standard_blank + \
                           self.str_vertical_line + self.put_value_in_blanks(16, month) + \
                           self.str_vertical_line + self.put_value_in_blanks(16, date) + \
                           self.str_vertical_line + ' '*16 + self.str_vertical_line + phone + ' ' + \
                           self.str_vertical_line + '\n' + self.str_line
            self.form_content_dict['row'] = n
            self.form_content_dict['content'] = content_line
            append_dict = self.form_content_dict.copy()
            self.form_content_info.append(append_dict)

    def assembly_all_form(self):
        title_info = self.assembly_title()
        self.modify_info_row_rules()
        self.all_info = self.str_line + '\n' + title_info + '\n' + self.str_line
        for i in range(2, self.total_form_row_cnt):
            for j in self.form_content_info:
                if j['row'] == i:
                    self.all_info = self.all_info + '\n' + j['content']

    def modify_info_row_rules(self):
        """--------------------------------------------------------------------------------------------  0   *
            |  姓名      |  车牌号码  |  年检到期月份    |  保险到期日期     |  投保保险公司   |  联系电话    |   1   *
           --------------------------------------------------------------------------------------------  2   *
            |  潘单萍    |  浙B9A1X8  |      2         |  20200604      |                |13777993974 |   3   0
            --------------------------------------------------------------------------------------------
            |慈溪市中凯金 |  浙BUS175  |      2         |  20200604      |                |13906747489 |   4   1
            属          |            |                |                |                |            |
            --------------------------------------------------------------------------------------------
            3 --- 0
            4 --- 1
            5 --- 2
            6 --- 3 """
        for f in self.form_content_info:
            f['row'] = f['row'] + 3

    @staticmethod
    def put_value_in_blanks(blanks_len, value):
        # not for chinese
        value_len = len(value)
        blanks_line = ' '*(blanks_len-value_len)
        blanks_line_len = len(blanks_line)
        half = blanks_line_len//2
        blanks_line_list = list(blanks_line)
        # blanks cut into half
        blanks_cut1 = blanks_line_list[0:half]
        blanks_cut2 = blanks_line_list[half:]
        # blanks_line_list.insert(int(blanks_len//2), value)
        # return ''.join(blanks_line_list)
        return ''.join(blanks_cut1) + value + ''.join(blanks_cut2)

    def cut_line_max(self, line):
        """a Chinese character needs 2 bytes to store
        that meas for name title, max 6 chinese characters
        line_len = len(line)
        line_bytes_len = line_len*2
        max_chinese_cnt = 6
        max_chinese_bytes = 12"""
        cut1 = line[0:6]
        cut2 = line[6:-1]
        self.cut_times = self.cut_times + 1
        return [cut1, cut2]

    def generate_strform(self):
        self.make_title1_len()
        self.make_str_form_row_cnt()
        # self.assembly_title()
        self.assembly_content_cut()
        self.assembly_content_normal()
        self.assembly_all_form()
        return self.all_info


# if __name__ == '__main__':
#     clue_data = ['5', '20200508']
#     test = query_register_excel_data(workbook1, clue_data)
#     Form = StrForm(test)
#     Form.generate_strform()
    # Form.make_title1_len()
    # Form.make_str_form_row_cnt()
    # Form.assembly_title()
    # Form.assembly_content_cut()
    # Form.assembly_content_normal()
    # Form.assembly_all_form()
    # test2 = get_plate_number(test)
    # print(test)
    # test1 = get_plate_number_info(workbook2)
    # test4 = get_specified_row(test2, workbook2)
    # test3 = get_specified_info(workbook2, test4)
    # test5 = get_specified_row('浙B1T0U5', workbook2)
    # print(test3)





