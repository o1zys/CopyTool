from tkinter import *
import main
import error
import global_var as gl

def do():
    error.Error.set_code(-1, "")
    hintLabel["text"] = ""
    main.execute(gl.file_a, gl.file_b, v1.get(), v2.get())
    error_code = error.Error.get_code()
    if error_code == -1:
        hintLabel["text"] = "Success!"
        hintLabel["fg"] = "green"
    else:
        hintLabel["text"] = "[Error " + str(error_code) + "] " + error.Error.get_info(error_code)
        hintLabel["fg"] = "red"


root = Tk()
root.title('Copy Tool')
root.columnconfigure(1, weight=1)
# root.iconbitmap(".\\icon.ico")
try:
    gl.set_var_from_config()
except Exception as e:
    tkinter.messagebox.showerror('错误', '读取配置文件config.txt失败\n' + str(e))

v1 = IntVar()
v2 = IntVar()

Checkbutton(root, text='Match SheetID (use with caution)', variable=v1, onvalue=1, offvalue=0).grid(row=0, column=0, sticky=SW)
Checkbutton(root, text='Ignore comments', variable=v2, onvalue=1, offvalue=0).grid(row=1, column=0, sticky=SW)


Button(root, text="Copy", command=do).grid(row=2, columnspan=3)
hintLabel = Label(root, text="")
hintLabel.grid(row=5, columnspan=3)


Label(root, text="Copyright © 2018 Oizys, All Rights Reserved").grid(row=9, columnspan=3, sticky=SW)
root.mainloop()
