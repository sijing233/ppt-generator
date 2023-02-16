import ttkbootstrap as tk
# import tkinter
from tkinter import filedialog
from tkinter import messagebox

from ttkbootstrap.dialogs import MessageDialog

from run import rungen

md_file_path = ""
model_name = ""

root = tk.Window(
    title="PPT生成器",  # 设置窗口的标题
    themename="litera",  # 设置主题
    size=(1066, 600),  # 窗口的大小
    position=(100, 100),  # 窗口所在的位置
    minsize=(0, 0),  # 窗口的最小宽高
    maxsize=(1920, 1080),  # 窗口的最大宽高
    resizable=None,  # 设置窗口是否可以更改大小
    alpha=1.0,  # 设置窗口的透明度(0.0完全透明）
)

max_w, max_h = root.maxsize()
root.resizable(width=False, height=False)

# 标签组件
label = tk.Label(root, text='选择你的markdown文件: ', font=('微软雅黑', 14), width=200)
label.place(x=50, y=100)

entry_text = tk.StringVar()
entry = tk.Entry(root, textvariable=entry_text, font=('FangSong', 10), width=50)
entry.place(x=350, y=105)


def get_path():
    path = filedialog.askopenfilename(title='请选择文件')
    global md_file_path
    md_file_path = path
    entry_text.set(path)


button = tk.Button(root, text='选择路径', command=get_path)
button.place(x=850, y=105)

model_label = tk.Label(root, text='请选择模板', width=200, font=('微软雅黑', 14))
model_label.place(x=50, y=200)

xVariable = tk.StringVar()

com = tk.Combobox(root, textvariable=xVariable)
com.place(x=350, y=200)
# 在这里增加下拉选项
com["value"] = ("通用")
com.current(0)


def xFunc(event):
    global model_name
    model_name = com.get()


com.bind("<<ComboboxSelected>>", xFunc)  # #给下拉菜单绑定事件


def gen_ppt_main():
    path_list = md_file_path.split('.')[0] + ".pptx"
    out_md_file_path = path_list
    rungen.run_gen(model_name, out_md_file_path, md_file_path)
    md = MessageDialog("生成成功")
    md.show()


run_button = tk.Button(root, text='生成PPT', command=gen_ppt_main)
run_button.place(x=550, y=505)
root.mainloop()

# pyinstaller -F -i D:\FFOutput\sijing233.ico .\ppt生成器.py
#