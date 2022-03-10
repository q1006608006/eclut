import tkinter
import tkinter.filedialog
from tkinter import messagebox
from tkinter import *
from tkinter import scrolledtext
import os
import json

import xlsnest

if __name__ != '__main__':
    os._exit(0) 

root = os.path.dirname(os.path.realpath(sys.argv[0]))
module_path = root + '/modules'
output = root + '/output/'

if not os.path.exists(module_path):
    os.mkdir(module_path)
if not os.path.exists(output):
    os.mkdir(output)

class xls_app_frame:

    width = 600
    height = 480
    max_width = 1200
    max_heigth = 960
    min_width = 300
    min_height = 240

    text_font = ('宋体', 12, 'normal')
    bt_font = ('宋体', 10, 'normal')

    def __init__(self):
        root = tkinter.Tk()
        self.__root__ = root
        root.title('Excel表换转换工具')
        self.set_size()

        def select_file(accept):
            fn = self.open_xls_file()
            if fn:
                accept(fn)
        
        # 菜单
        menu = Menu(root)
        root['menu'] = menu
        func_menu = Menu(menu,tearoff = False)
        menu.add_cascade(label = '其他功能',menu = func_menu)
        
        # 添加模板
        def open_explore(path):
            os.system("start {}".format(path))

        func_menu.add_command(label='打开模板文件夹',command = lambda : open_explore(module_path))
        func_menu.add_command(label='打开输出文件夹',command = lambda : open_explore(output))

        
        # 选择模板区域
        mod_frame = Frame(root)
        mod_frame.grid(row=0,column=0,rowspan=3,columnspan=5,padx=30,pady=20)

        label =  Label(mod_frame,text = '模板文件:',font = self.text_font)
        label.pack(side = LEFT)

        entry = Entry(mod_frame,width = 50)
        entry.pack(side=LEFT,padx = 10)
        entry_val = StringVar()
        self._get_mod_path = entry_val.get
        entry['textvariable'] = entry_val

        bt = Button(mod_frame,font = self.bt_font,text='浏览',command = lambda : select_file(entry_val.set))
        bt.pack(side=RIGHT)
        
        # 选择数据区域
        data_frame = Frame(root)
        data_frame.grid(row=4,column=0,rowspan=10,columnspan=5,padx=30,pady=10)

        data_text = scrolledtext.ScrolledText(data_frame,width=66,height=22)
        data_text.pack(side=BOTTOM,padx=30)

        def get_text_val():
            return data_text.get(1.0,END)

        self._get_data = get_text_val

        def read_mod():
            mod_fn = self._get_mod_path()
            if mod_fn:
                data_fns = self.open_xls_files(initialdir=os.path.join(os.path.expanduser("~"), 'Desktop'))
                for data_fn in data_fns:
                    try:
                        esmod = xlsnest.read_xls_mod(mod_fn)
                    except Exception:
                        tkinter.messagebox.showerror('提示','无法读取模板文件，请重新选择模板')
                        return
                    
                    merge_inindex = self._get_inindex_val()
                    data = esmod.load(data_fn,inindex = merge_inindex)
                    ori_data = self._get_data()
                    ori_data = ori_data.replace('\r','').replace('\n','')
                    
                    if ori_data and not ori_data == '':
                        try:   
                            ori_data = json.loads(ori_data)
                        except json.decoder.JSONDecodeError:
                            tkinter.messagebox.showerror('提示','解析文本框数据失败，请重试')
                            return
                        
                        merge_index = self._get_index_val()
                        try:
                            data = xlsnest.merged_defines(ori_data,data,merge_index)
                        except xlsnest.EsMergedException as err:
                            tkinter.messagebox.showerror('警告',err.msg)
                            return
                    
                    data_text.delete(0.0,END)
                    data_text.insert(0.0,json.dumps(data,ensure_ascii = False,indent = 1))
            else:
                tkinter.messagebox.showinfo('提示','请先选择模板')

        Label(data_frame).pack(side=LEFT,padx=10)
        Button(data_frame,font = self.bt_font,text='读取',command = read_mod).pack(side=LEFT,padx=5,pady=5)
        Button(data_frame,font = self.bt_font,text='清空',command = lambda : data_text.delete(0.0,END)).pack(side=LEFT,padx=5,pady=5)

        index_var = StringVar()
        inindex_var = StringVar()

        Label(data_frame,text='    关联值:',font = self.text_font).pack(side = LEFT)
        Entry(data_frame,textvariable=index_var,width=10).pack(side=LEFT,padx=5,pady=5)
        Label(data_frame,text='  sheet关联值:',font = self.text_font).pack(side = LEFT)
        Entry(data_frame,textvariable=inindex_var,width=10).pack(side=LEFT,padx=5,pady=5)

        self._get_index_val = index_var.get
        self._get_inindex_val = inindex_var.get

        # 导出数据
        out_frame = Frame(root)
        out_frame.grid(row = 15)

        def data_out_put():
            fn = self.open_xls_file()
            if fn:
                try:
                    out_mod = xlsnest.read_xls_mod(fn)
                except Exception:
                    tkinter.messagebox.showerror('提示','模板文件不可用，请重新选择')
                    return
                try:
                    data = json.loads(self._get_data())
                except Exception:
                    tkinter.messagebox.showerror('提示','无法使用当前数据，请保证数据完整性')
                    return

                try:
                    out_path = output + os.path.split(fn)[1]
                    out_mod.write(data,out_path)
                except FileNotFoundError as fnf:
                    tkinter.messagebox.showerror("警告",fnf.strerror)
                    return
                except Exception:
                    tkinter.messagebox.showerror('警告','写入失败，请查看日志排查')
                    return
                
                tkinter.messagebox.showinfo('提示','导出成功')
                os.system("start {}".format(output))

        Label(out_frame).pack(side=LEFT,padx=195)
        def remove_blank():
            data = self._get_data()
            try:
                data = json.loads(data)
                data = xlsnest.remove_blank_row(data)
            except xlsnest.EsMergedException as eserr:
                tkinter.messagebox.showerror('警告',eserr.msg)
            except Exception:
                tkinter.messagebox.showerror('警告','数据载入失败')
            
            data_text.delete(0.0,END)
            data_text.insert(0.0,json.dumps(data,ensure_ascii = False,indent = 1))

        Button(out_frame,text='清理空行',command=remove_blank).pack(side=LEFT,padx = 5)
        Button(out_frame,text='导出数据',command=data_out_put).pack(side=RIGHT)

    def set_size(self,width = width,height = height,max_width = max_width,max_heigth = max_heigth,min_width = min_width,min_height = min_height):
        root = self.__root__
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()

        size = '%dx%d+%d+%d' % (width, height, (screenwidth -width)/2, (screenheight - height)/2)
        root.geometry(size)
        root.maxsize(max_width,max_heigth)
        root.minsize(min_width,min_height)

    def start(self):
        self.__root__.mainloop()

    def open_xls_file(self,initialdir = module_path):
        filename = tkinter.filedialog.askopenfile(
            defaultextension='.xls',                #默认文件的扩展名
            filetypes=[('excel Files', '*.xls')],   #设置文件类型下拉菜单里的的选项
            initialdir=initialdir,                 #对话框中默认的路径
            parent=self.__root__,                   #父对话框(由哪个窗口弹出就在哪个上端)
            title="打开"                            #弹出对话框的标题
        )

        return filename.name if filename else None

    def open_xls_files(self,initialdir = module_path):
        fnl = tkinter.filedialog.askopenfilenames(
            defaultextension='.xls',                #默认文件的扩展名
            filetypes=[('excel Files', '*.xls;*.xlsx')],   #设置文件类型下拉菜单里的的选项
            initialdir=initialdir,                 #对话框中默认的路径
            parent=self.__root__,                   #父对话框(由哪个窗口弹出就在哪个上端)
            title="选择文件"                         #弹出对话框的标题
        )

        return list(fnl)

app = xls_app_frame()
app.start()