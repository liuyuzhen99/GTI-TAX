import tkinter as tk
from readExcelClass import ReadExcelClass
from budgetClass import BudgetClass
from budgetClass import MonthlyBudgetClass
# from readerClass import ReaderClass
from thinkcellClass import ThinkcellClass


def on_button_click1():
    text = entry1.get()
    if not text:
        label1.config(text="excel file path cannot be empty!")
    else:
        label1.config(text=text)


def on_button_click2():
    text = entry2.get()
    if not text:
        label2.config(text="ppt template file path cannot be empty!")
    else:
        label2.config(text=text)


def on_button_click3():
    text1 = entry1.get()
    text2 = entry2.get()
    if not text1 or not text2:
        label3.config(text="excel or ppt file path cannot be empty!")
    else:
        '''------------------------------'''
        # 主函数开始
        df_class = ReadExcelClass(entry1.get())
        df_class.read_excel_file()

        budget_class = BudgetClass(df_class)
        budget_class.get_monthly_budget()

        thinkcell_class = ThinkcellClass(entry2.get(), budget_class)
        thinkcell_class.add_template()
        thinkcell_class.add_chart()
        thinkcell_class.save_file()
        # 主函数结束
        '''------------------------------'''
        label3.config(text="ppt created successfully! " + thinkcell_class.filename)


# 创建主窗口
window = tk.Tk()
window.title("excel2ppt")

# 创建标签
label1 = tk.Label(window, text="pls input excel file path ")
label1.pack()

# 创建文本框
entry1 = tk.Entry(window)
entry1.pack()

# 创建按钮
button1 = tk.Button(window, text="confirm", command=on_button_click1)
button1.pack()

# 创建一个Frame用于分割
separator_frame1 = tk.Frame(window, height=20)
separator_frame1.pack()

# 创建标签
label2 = tk.Label(window, text="pls input ppt template file path ")
label2.pack()

# 创建文本框
entry2 = tk.Entry(window)
entry2.pack()

# 创建按钮
button2 = tk.Button(window, text="confirm", command=on_button_click2)
button2.pack()

# 创建一个Frame用于分割
separator_frame2 = tk.Frame(window, height=20)
separator_frame2.pack()

# 创建按钮
button3 = tk.Button(window, text="submit", command=on_button_click3)
button3.pack()

# 创建标签
label3 = tk.Label(window, text="")
label3.pack()

# 进入著事件循环
window.mainloop()
