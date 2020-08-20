import tkinter as tk
import os
import form_operation as form

from tkinter import *
from tkinter import scrolledtext
from tkinter import filedialog
from tkinter import messagebox


class MyWin:
    def __init__(self, win):
        self.win = win
        self.title = self.win.title('金南检测客户信息查询V1.5')
        self.size = self.win.minsize(300, 200)
        self.path1 = tk.StringVar()
        self.path1.set('')
        self.path2 = tk.StringVar()
        self.path2.set('')
        self.cnt = tk.StringVar()
        self.cnt.set('')
        self.prompt_str = tk.StringVar()
        self.prompt_str.set('')
        self.info2_cnt = tk.StringVar()
        self.info2_cnt.set('')
        self.current_info2_num = tk.StringVar()
        self.current_info2_num.set('')
        self.dir = ''
        self.output_file1 = None
        self.output_file1_path_ins = tk.StringVar()
        self.output_file1_path_ins.set('')
        self.output_file1_path = ''
        self.data1 = None
        self.data1_1 = None
        self.data2 = None
        self.data2_new_format = []
        self.current_num = 0
        self.entry1 = None
        self.entry2 = None
        self.entry3 = None
        self.labelF1_Window = None
        self.labelF2_Window = None
        self.excel_data1 = None
        self.excel_data2 = None
        self.init_windows()

    def init_windows(self):
        label1 = tk.Label(self.win, text='欢迎使用', font=('microsoft yahei', 15), fg='DarkViolet')
        label1.place(x=80, y=0)
        label2 = tk.Label(self.win, text='文件1所在路径:', padx=30, anchor='w', width=8, height=1)
        label2.place(x=50, y=50)
        label2_1 = tk.Label(self.win, textvariable=self.path1)
        label2_1.place(x=170, y=50)
        label3 = tk.Label(self.win, text='文件2所在路径:', padx=30, anchor='w', width=8, height=1)
        label3.place(x=50, y=80)
        label3_1 = tk.Label(self.win, textvariable=self.path2)
        label3_1.place(x=170, y=80)
        button1 = tk.Button(self.win, text='打开文件', width=8, height=1, command=lambda: self.read_file1())
        button1.place(x=10, y=50)
        button2 = tk.Button(self.win, text='打开文件', width=8, height=1, command=lambda: self.read_file2())
        button2.place(x=10, y=80)
        button3 = tk.Button(self.win, text='退出', width=8, height=1, command=lambda: self.exit_tk())
        button3.place(x=500, y=60)
        button4 = tk.Button(self.win, text='清除数据', width=8, height=1, command=lambda: self.clear_data())
        button4.place(x=580, y=60)

    def exit_tk(self):
        self.win.destroy()

    def _open_file(self):
        # when exe open operation, withdraw windows.
        # self.win.withdraw()
        default_dir = r'文件路径'
        filepath = filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser(default_dir)),
                                              filetypes=[('', '*.xls;*.xlsx')])
        print(filepath)
        (self.dir, filename) = os.path.split(filepath)
        return filepath

    def clear_data(self):
        self.cnt.set('')
        self.entry1.delete(0, 'end')
        self.entry2.delete(0, 'end')
        if self.output_file1_path:
            self.output_file1_path_ins.set('')
        self.labelF1_Window.delete(0.0, 'end')
        if self.excel_data2:
            self.entry3.delete(0, 'end')
            self.labelF2_Window.delete(0.0, 'end')
        self.cnt.set('')
        self.info2_cnt.set('')
        self.current_info2_num.set('')
        self.data1 = None
        self.data1_1 = None
        self.data2_new_format = []

    def read_file1(self):
        filepath = self._open_file()
        if filepath:
            self.path1.set(filepath)  # put "filepath" to the win
            # then show the first stage button.
            self.first_stage_winInit()
            self.parse_info1()
        else:
            print('没有选择任何文件')

    def read_file2(self):
        filepath = self._open_file()
        if filepath:
            self.path2.set(filepath)  # put "filepath" to the win
            # then show the first stage button.
            self.second_stage_winInit()
            self.parse_info2()
        else:
            print('没有选择任何文件')

    def first_stage_winInit(self):
        self.size = self.win.minsize(820, 650)
        input_name_label1 = tk.Label(self.win, text='年检到期月份：').place(x=80, y=150)
        input_name_label2 = tk.Label(self.win, text='保险到期日期：').place(x=80, y=180)
        self.entry1 = tk.Entry(self.win, bg='#ffffff', width=20)
        self.entry1.place(x=160, y=150, anchor=NW)
        self.entry2 = tk.Entry(self.win, bg='#ffffff', width=20)
        self.entry2.place(x=160, y=180, anchor=NW)
        label2 = tk.Label(self.win, text='查询到符合条件的车牌号数量：', padx=30, anchor='w', width=25, height=1)
        label2.place(x=50, y=220)
        label2_1 = tk.Label(self.win, textvariable=self.cnt)
        label2_1.place(x=260, y=220)
        labelF1 = tk.LabelFrame(self.win, text='符合条件的车牌：', padx=10, pady=10)
        labelF1.place(x=50, y=260)
        self.labelF1_Window = scrolledtext.ScrolledText(labelF1, width=100, height=12, padx=10, pady=10, wrap=tk.WORD)
        self.labelF1_Window.grid()
        generate_output_excel_button = tk.Button(self.win, bg='white', text='生成并输出文件', width=15, height=1,
                                                 command=lambda: self.generate_file_show())
        generate_output_excel_button.place(x=50, y=480)
        query_button = tk.Button(self.win, bg='white', text='查询', width=10, height=1,
                                 command=lambda: self.query_and_output_info1(self.excel_data1))
        query_button.place(x=300, y=160, anchor='nw')
        print('hhhhh')

    def first_stage_one_winInit(self):
        label4 = tk.Label(self.win, text='生成文件所在路径：', padx=30, anchor='w', width=25, height=1)
        label4.place(x=170, y=480)
        label4_1 = tk.Label(self.win, textvariable=self.output_file1_path_ins)
        label4_1.place(x=300, y=480)
        open_generate_file_button = tk.Button(self.win, bg='white', text='打开生成的文件', width=15, height=1,
                                              command=lambda: self.open_generate_file())
        open_generate_file_button.place(x=50, y=520)

    def second_stage_winInit(self):
        if self.path1.get() == '':
            tk.messagebox.showinfo('提示', '请先打开文件一')
        else:
            self.size = self.win.minsize(820, 1000)
            input_name_label1 = tk.Label(self.win, text='请输入车牌号：').place(x=50, y=580)
            self.entry3 = tk.Entry(self.win, bg='#ffffff', width=20)
            self.entry3.place(x=160, y=580)
            output_button = tk.Button(self.win, bg='white', text='输出相关信息', width=10,
                                      command=lambda: self.query_and_output_info2(self.excel_data2))
            output_button.place(x=300, y=580, anchor='nw')
            output_button_all = tk.Button(self.win, bg='white', text='导入所有符合条件的车牌号并输出', width=26,
                                          command=lambda: self.import_all_clue_and_output_info())
            output_button_all.place(x=300, y=620, anchor='nw')
            labelF2 = tk.LabelFrame(self.win, text='信息集合：', padx=10, pady=10)
            labelF2.place(x=50, y=650)
            self.labelF2_Window = scrolledtext.ScrolledText(labelF2, width=60, height=15, padx=10, pady=10, wrap=tk.WORD)
            self.labelF2_Window.grid()

    def get_input_conditional1(self):
        value1 = self.entry1.get()
        value2 = self.entry2.get()
        value_list = [value1, value2]
        return value_list

    def get_input_conditional2(self):
        value3 = self.entry3.get()
        value_list = [value3]
        return value_list

    def parse_info1(self):
        file_path = self.path1.get()
        self.excel_data1 = form.open_excel(file_path)

    def parse_info2(self):
        file_path = self.path2.get()
        self.excel_data2 = form.open_excel(file_path)

    def query_and_output_info1(self, workbook):
        self.output_file1_path_ins.set('')
        clue_data = self.get_input_conditional1()
        if not clue_data[0] or not clue_data[1]:
            tk.messagebox.showinfo('提示', '请输入查询条件！')
            return
        self.data1 = form.query_register_excel_data(workbook, clue_data)
        if not self.data1:
            self.labelF1_Window.delete(0.0, 'end')
            self.cnt.set('')
            self.prompt_str.set('')
            self.output_file1_path_ins.set('')
            tk.messagebox.showinfo('提示', '没有符合条件的车牌号！')
            return
        self.info1_cnt(self.data1)
        self.data1_1 = form.get_plate_number(self.data1)
        if len(self.data1) > 6:
            self.prompt_str = tk.StringVar()
            self.prompt_str.set("（提示：如果下方表格查看不便，请生成Excel文件更容易查看！）")
            prompt_label = tk.Label(self.win, textvariable=self.prompt_str)
            prompt_label.place(x=300, y=220)
        form_text = form.StrForm(self.data1).generate_strform()
        self.labelF1_Window.delete(0.0, 'end')
        self.labelF1_Window.insert('end', form_text)
        self.labelF1_Window.see(0.0)
        # self.info1_about_plate_number(self.data1)

    def info1_cnt(self, data):
        info1_cnt = len(data)
        self.cnt.set(info1_cnt)

    def generate_str_form(self):
        str_line = '---------------------------------------------------------------------------------------'
        str_title = '|  姓名  |  车牌号码  |  年检到期月份  |  保险到期日期  |  投保保险公司  |  联系电话  |'
        str_form = str_line + '\n' + str_title + '\n' + str_line + '\n'
        for i in self.data1:
            str_blank1 = '|  '
            str_blank2 = '  |  '
            str_blank3 = '  |'
            str_blank_name = i['姓名']
            str_blank_number = i['车牌号码']
            str_blank_month = i['年检到期月份']
            str_blank_date = i['保险到期日期']
            str_blank_company = i['投保保险公司']
            str_blank_phone = i['联系电话']
            str_blank = str_blank1 + str_blank_name + str_blank2 + str_blank_number + str_blank2 + str_blank_month + \
                        str_blank2 + str_blank_date + str_blank2 + str_blank_company + str_blank2 + str_blank_phone + \
                        str_blank3
            str_blank_form = str_blank + '\n' + str_line + '\n'
            str_form = str_form + str_blank_form
        print(str_form)
        return str_form

    def adjust_output_form(self):
        str_form = self.generate_str_form()
        cnt = int(self.cnt.get())
        pass
        # for i in range(cnt+1):      # +1 : include title

    def info1_about_plate_number(self, data):
        self.data1_1 = form.get_plate_number(data)
        self.labelF1_Window.delete(0.0, 'end')
        self.labelF1_Window.insert('end', str(self.data1_1))
        self.labelF1_Window.see(0.0)

    def generate_file_show(self):
        print('data1" %s' % self.data1)
        if not self.data1:
            tk.messagebox.showinfo('提示', '请先获取相应数据！')
        else:
            self.output_file1_path = form.create_excel(self.dir, '符合条件的车牌相关信息.xls', self.data1)
            self.output_file1_path_ins.set(self.output_file1_path)
            self.first_stage_one_winInit()

    def open_generate_file(self):
        os.startfile(self.output_file1_path)

    def query_and_output_info2(self, workbook):
        clue_data = self.get_input_conditional2()
        if not clue_data[0]:
            tk.messagebox.showinfo('提示', '请先输入车牌号！')
            return
        clue_row = form.get_specified_row(clue_data, workbook)
        if not clue_row:
            tk.messagebox.showinfo('提示', '在文件2中没有找到相应车牌信息')
            self.labelF2_Window.delete(0.0, 'end')
            self.current_info2_num.set('')
            self.info2_cnt.set('')
            return
        self.current_info2_num.set('1')
        self.info2_cnt.set('1')
        self.data2 = form.get_specified_info(workbook, clue_row)
        self.data2_new_format = []    # if first use 'import_all_clue_and_output_info', need clear this data
        self.current_num = 0
        self.modify_info2_format()
        self.info2_about_info_collection(self.data2_new_format[0])

    def info2_about_info_collection(self, data):
        self.labelF2_Window.delete(0.0, 'end')
        self.labelF2_Window.insert('end', data)
        self.labelF2_Window.see(0.0)

    def import_all_clue_and_output_info(self):
        if not self.data1_1:
            tk.messagebox.showinfo('提示', '请先在文件1种查询符合条件的车牌号！')
            self.labelF2_Window.delete(0.0, 'end')
            return
        row_list = form.get_specified_row(self.data1_1, self.excel_data2)
        self.data2 = form.get_specified_info(self.excel_data2, row_list)
        self.modify_info2_format()
        self.info2_cnt.set(len(self.data2))
        self.show_info2_cnt()
        self.show_current_info2_num()
        if len(self.data2) > 1:
            self.last_button_init()
            self.next_button_init()
            self.info2_about_info_collection(self.data2_new_format[0])
        else:
            self.info2_about_info_collection(self.data2_new_format[0])

    def show_info2_cnt(self):
        label5 = tk.Label(self.win, text='信息集合数量：', padx=30, anchor='w', width=20, height=1)
        label5.place(x=500, y=620)
        label5_1 = tk.Label(self.win, textvariable=self.info2_cnt)
        label5_1.place(x=650, y=620)

    def next_button_init(self):
        next_button = tk.Button(self.win, bg='white', text='查看下一个', width=10, height=3,
                                 command=lambda: self.next_button_operation())
        next_button.place(x=700, y=650, anchor='nw')

    def last_button_init(self):
        last_button = tk.Button(self.win, bg='white', text='查看上一个', width=10, height=3,
                                 command=lambda: self.last_button_operation())
        last_button.place(x=560, y=650, anchor='nw')

    @staticmethod
    def modify_format(info):
        new_format_str = ''
        keys = info.keys()
        for k in keys:
            new_format_str = new_format_str + k + ':' + info[k] + '\n'
        return new_format_str

    def modify_info2_format(self):
        for d in self.data2:
            self.data2_new_format.append(self.modify_format(d))

    def next_button_operation(self):
        data2_num = len(self.data2_new_format)
        if not data2_num:
            tk.messagebox.showinfo('提示', '没有数据源！请查询添加')
            return
        if self.current_num + 1 >= data2_num:
            tk.messagebox.showinfo('提示', '查看完毕！')
            return
        else:
            self.current_num = self.current_num + 1
            self.info2_about_info_collection(self.data2_new_format[self.current_num])
            self.show_current_info2_num()

    def last_button_operation(self):
        data2_num = len(self.data2_new_format)
        if not data2_num:
            tk.messagebox.showinfo('提示', '没有数据源！请查询添加')
            return
        if self.current_num == 0:
            tk.messagebox.showinfo('提示', '已经第一个了！')
            return
        else:
            self.current_num = self.current_num - 1
            self.info2_about_info_collection(self.data2_new_format[self.current_num])
            self.show_current_info2_num()

    def show_current_info2_num(self):
        label5 = tk.Label(self.win, text='当前第：    个信息集合', padx=30, anchor='w', width=20, height=1)
        label5.place(x=550, y=750)
        label5_1 = tk.Label(self.win, textvariable=self.current_info2_num)
        label5_1.place(x=625, y=750)
        self.current_info2_num.set(self.current_num + 1)


# if __name__ == '__main__':
#     # clue_data = ['5', '20200506']
#     # test = query_register_excel_data(filepath, clue_data)
#     # test1 = get_plate_number_info(filepath1)
#     # test3 = get_specified_info(workbook2, 15)
#     # print(test3)
#     win = tk.Tk()
#     QUERY = MyWin(win)
#     win.mainloop()
