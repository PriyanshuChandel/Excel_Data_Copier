from openpyxl import load_workbook  # To manage the excel processing
from warnings import filterwarnings  # To ignore the warnings
from threading import Thread  # To use multiple process threads
from time import time  # To calculate the elapsed time by program
from os.path import join, dirname  # To manage the files directory
from tkinter import Tk, filedialog, Label, Button, Entry  # To create GUI
from datetime import datetime  # Used in logging system

icon_file = join(dirname(__file__), 'Excel.ico')
window = Tk()
window.config(bg='grey')
window.title('Copier - Developed by Priyanshu')
window.minsize(width=800, height=385)
window.maxsize(width=800, height=385)
window.iconbitmap(icon_file)

def threading_btn2():
    thread_btn2 = Thread(target=btn2_func)
    thread_btn2.start()

def btn2_func():
    global excel_file_path_src
    excel_file_path_src = filedialog.askopenfilename(
        filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
    ent2.insert(0, excel_file_path_src)

def threading_btn3():
    thread_btn3 = Thread(target=btn3_func)
    thread_btn3.start()

def btn3_func():
    global excel_file_path_dest
    excel_file_path_dest = filedialog.askopenfilename(
        filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
    ent3.insert(0, excel_file_path_dest)

def threading_btn4():
    thread_btn4 = Thread(target=search_copy_func)
    thread_btn4.start()

def search_copy_func():
    labl12.config(text="I know you are lazy to do this boring copy-paste job. NVM take rest while I complete this task."
                       "\nProcessing...")
    start = time()
    filterwarnings("ignore", category=DeprecationWarning)
    workbook_object_src = load_workbook(excel_file_path_src)
    workbook_object_dest = load_workbook(excel_file_path_dest)
    src_excel_sheet_name = ent4.get()
    dest_excel_sheet_name = ent5.get()
    sheet_obj_src = workbook_object_src.get_sheet_by_name(f'{src_excel_sheet_name}')
    sheet_obj_dest = workbook_object_dest.get_sheet_by_name(f'{dest_excel_sheet_name}')
    search_column_src = int(ent6.get())
    search_column_dest = int(ent7.get())
    src_column_list = list()
    dest_column_list = list()
    [src_column_list.append(int(item)) for item in ent9.get().split(',')]
    [dest_column_list.append(int(item)) for item in ent10.get().split(',')]
    no_of_max_rows_src = sheet_obj_src.max_row
    no_of_max_rows_dest = sheet_obj_dest.max_row
    file_handler = open(f"logs_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt", 'a')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} {excel_file_path_src} selected as source file \n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} {excel_file_path_dest} selected as destination file \n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} {sheet_obj_src} selected as source sheet \n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} {sheet_obj_dest} selected as destination sheet \n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} column {search_column_src} is source reference\n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} column {search_column_dest} is destination reference \n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} {src_column_list} are the source columns\n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} {dest_column_list} are the destination columns\n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} {no_of_max_rows_src} are the total source rows\n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} {no_of_max_rows_dest} are the total destination rows\n')
    reference_cell_value_for_which_data_copied = list()
    for row_index_src in range(2, no_of_max_rows_src + 1):
        try:
            if not sheet_obj_src.row_dimensions[row_index_src].hidden:
                cell_src_value = sheet_obj_src.cell(row=row_index_src, column=search_column_src).value
                file_handler.write(f'{datetime.now().replace(microsecond=0)} {cell_src_value} is source cell value\n')
                for row_index_dest in range(2, no_of_max_rows_dest + 1):
                    try:
                        if not sheet_obj_dest.row_dimensions[row_index_dest].hidden:
                            cell_dest_value = sheet_obj_dest.cell(row=row_index_dest, column=search_column_dest).value
                            file_handler.write(f'{datetime.now().replace(microsecond=0)} {cell_dest_value} is '
                                               f'destination cell value\n')
                            file_handler.write(f'{datetime.now().replace(microsecond=0)} comparing source cell value '
                                               f'[{cell_src_value}] with destination cell value [{cell_dest_value}]\n')
                            if str(cell_src_value) in str(cell_dest_value):
                                reference_cell_value_for_which_data_copied.append(cell_dest_value)
                                for column_src, column_dest in zip(src_column_list, dest_column_list):
                                    temp = sheet_obj_dest.cell(row=row_index_dest, column=int(column_dest)).value = \
                                        sheet_obj_src.cell(row=row_index_src, column=int(column_src)).value
                                    workbook_object_dest.save(str(excel_file_path_dest))
                                    file_handler.write(f'{datetime.now().replace(microsecond=0)} match found, [{temp}] '
                                                       f'from source cell [column_row] [{column_src}_{row_index_src}] '
                                                       f'copied in destination cell [column_row] '
                                                       f'[{column_dest}_{row_index_dest}]\n')
                                    temp = ''
                                break
                    except Exception as e:
                        file_handler.write(f'{datetime.now().replace(microsecond=0)} [ERROR1] [{e}]\n')
            continue
        except Exception as e:
            file_handler.write(f'{datetime.now().replace(microsecond=0)} [ERROR2] [{e}]\n')
    end = time()
    file_handler.write(f'{datetime.now().replace(microsecond=0)} Data for {len(reference_cell_value_for_which_data_copied)} '
                       f'cells whose values are [{reference_cell_value_for_which_data_copied}] respectively copied\n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} TASK COMPLETED]\n')
    file_handler.write(f'{datetime.now().replace(microsecond=0)} ELAPSED TIME TO COMPLETE THIS TASK IS '
                       f'{((end - start) / 60)} Minutes\n')
    file_handler.close()
    labl12.config(text='Task completed, check logs file for the status')

labl1 = Label(window, text='Excel Data Copier', font=(None, 17, 'bold'), bg='grey').place(x=285, y=1)
labl2 = Label(window, text='Select the source excel file', font=(None, 9, 'bold'), bg='grey')
labl2.place(x=0, y=40)
ent2 = Entry(window, bd=4, width=47, bg='lavender')
ent2.place(x=180, y=40)
btn2 = Button(window, text='...', command=threading_btn2, bg='green')
btn2.place(x=480, y=38)
labl3 = Label(window, text='Select the target excel file', font=(None, 9, 'bold'), bg='grey')
labl3.place(x=0, y=70)
ent3 = Entry(window, bd=4, width=47, bg='lavender')
ent3.place(x=180, y=70)
btn3 = Button(window, text='...', command=threading_btn3, bg='green')
btn3.place(x=480, y=68)
labl4 = Label(window, text='Enter the name of sheet from source excel file', font=(None, 9, 'bold'), bg='grey')
labl4.place(x=0, y=100)
ent4 = Entry(window, bd=4, width=47, bg='lavender')
ent4.place(x=270, y=100)
labl5 = Label(window, text='Enter the name of sheet from target excel file', font=(None, 9, 'bold'), bg='grey')
labl5.place(x=0, y=130)
ent5 = Entry(window, bd=4, width=47, bg='lavender')
ent5.place(x=270, y=130)
labl6 = Label(window, text='Enter the source reference column index(1 for A, 2 for B ...n for N)'
              , font=(None, 9, 'bold'), bg='grey')
labl6.place(x=0, y=160)
ent6 = Entry(window, bd=4, width=47, bg='lavender')
ent6.place(x=370, y=160)
labl7 = Label(window, text='Enter the target reference column index(1 for A, 2 for B ...n for N)',
              font=(None, 9, 'bold'), bg='grey')
labl7.place(x=0, y=190)
ent7 = Entry(window, bd=4, width=47, bg='lavender')
ent7.place(x=370, y=190)
labl9 = Label(window,
              text='Enter the index of source columns from which data to be copied. Enter like 1,2...n for A,'
                   'B..N', font=(None, 9, 'bold'), bg='grey')
labl9.place(x=0, y=220)
ent9 = Entry(window, bd=4, width=45, bg='lavender')
ent9.place(x=514, y=220)
labl10 = Label(window,
               text='Enter the index of target columns into which data to be copied. Enter like 1,2...n for '
                    'A,B..N', font=(None, 9, 'bold'), bg='grey')
labl10.place(x=0, y=250)
ent10 = Entry(window, bd=4, width=47, bg='lavender')
ent10.place(x=502, y=250)
btn4 = Button(window, text='Submit', font=(None, 9, 'bold'), command=threading_btn4, bg='green')
btn4.place(x=350, y=280)
labl11 = Label(window, font=(None, 9, 'bold'), text="Note: Program may not work well will excel extensions .xlsm, "
                                                    ".xlsb, .xltx, .xltm and .xlt. Let the program finish completely "
                                                    "and do not close it \nwhile it is on progress, otherwise TARGET "
                                                    "EXCEL FILE WILL CORRUPT.",bg='grey', fg = 'orange',
               justify='left').place(x=3,y=310)
labl12 = Label(window, bg='grey',font=(None, 9, 'bold'),justify='left')
labl12.place(x=0, y=345)
window.mainloop()