import tkinter as tk
from Windows import MyWin

if __name__ == '__main__':
    # clue_data = ['5', '20200506']
    # test = query_register_excel_data(filepath, clue_data)
    # test1 = get_plate_number_info(filepath1)
    # test3 = get_specified_info(workbook2, 15)
    # print(test3)
    win = tk.Tk()
    QUERY = MyWin(win)
    win.mainloop()
