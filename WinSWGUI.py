import platform
import ctypes, os, time
from tkinter import *
from tkinter.filedialog import askopenfilename
import xml.etree.ElementTree as ET
from tkinter import messagebox
import subprocess
import tkinter as tk
from tkinter import simpledialog
from qq import img
import base64

class WinSWGUI:
    #查看系统构架
    def checksysplatform(self):
        sysInfo = platform.architecture()
        print(sysInfo)
        return sysInfo[0]
    
    #判断管理员权限，0为无权限，1为已获得权限
    def isadmin(self):
        return ctypes.windll.shell32.IsUserAnAdmin()
        
    #判断winsw是否存在
    def check_files(self, path):
        #获取当前运行程序所在路径
        #curr_path = os.getcwd()
        #遍历文件夹
        for dir in os.listdir(path):
            dir = os.path.join(path, dir)
            #print(dir)
            #判断是否为文件夹
            if os.path.isdir(dir):
                self.check_files(dir)
            else:
                filename = str(os.path.basename(dir))
                filename = filename.lower()
                print(filename)
                sysinfo = str(self.checksysplatform())[0,1]
                if "exe" and "winsw" and sysinfo in filename:
                    return dir

class GUI:
    def __init__(self):
        self.tk = Tk()
        self.tk.title("AddService")
        self.tk.geometry("280x260")
        self.set_ui()
        
        self.setpermissionText()
        self.set_icon()
        
        #初始化服务状态
        self.check_service_status()
        
        #初始化菜单栏
        self.init_menu()
        
        #可执行文件路径（需要安装为服务的软件）
        self.ptah_excutable = ""
        self.main_exe = "WinSW.exe"
        self.main_exe_xml = "WinSW.xml"
    
    def run(self):
        self.tk.mainloop()

    def __del__(self):
        self.tk.destroy()

    #初始化UI
    def set_ui(self):
        #权限状态   
        self.label_permission_status = Label(self.tk, text="权限状态")
        self.label_permission_status.grid(column=0, row=0)
        
        #服务状态   
        self.service_status_label = Label(self.tk, text="服务状态")
        self.service_status_label.grid(column=1, row=0)
        
        #已选择文件路径
        self.label_file_path = Label(self.tk, text="已选择文件")
        self.label_file_path.grid(column=0, row=1, columnspan = 3)
        
        #选择文件按钮
        self.select_file_button = Button(self.tk, text="选择文件",command=self.select_file)
        self.select_file_button.grid(column=0, row=2, columnspan = 3, pady = 10)
        
        #label_id   
        self.label_id = Label(self.tk, text="id：")
        self.label_id.grid(column=0, row=3, ipadx = 0)
        #id
        self.entry_id = Entry(self.tk, width=20)
        self.entry_id.grid(column=1, row=3)
        
        #label_name   
        self.label_name = Label(self.tk, text="name：")
        self.label_name.grid(column=0, row=4)
        #name
        self.entry_name = Entry(self.tk, width=20)
        self.entry_name.grid(column=1, row=4)
        
        #label_dercription  
        self.label_dercription = Label(self.tk, text="dercription：")
        self.label_dercription.grid(column=0, row=5)       
        #dercription
        self.entry_description = Entry(self.tk, width=20, xscrollcommand = 1)
        self.entry_description.grid(column=1, row=5) 
        
        #安装服务按键
        self.install_button = Button(self.tk, text="安装为服务",command=self.install_service)
        self.install_button.grid(column=0, row=7, columnspan = 2, pady = 10)
        
        #卸载服务按键
        self.uninstall_button = Button(self.tk, text="卸载服务")
        self.uninstall_button.grid(column=0, row=8, columnspan = 2)
        
        #绑定卸载按钮的左右键单击事件
        self.uninstall_button.bind("<Button-1>", self.uninstall_service)  # 绑定左键单击事件
        self.uninstall_button.bind("<Button-3>", self.uninstall_service_others)  # 绑定右键单击事件
        
    #设置图标
    def set_icon(self):
        print("开始设置图标")
        # 将import进来的icon.py里的数据转换成临时文件tmp.ico，作为图标
        tmp = open("tmp.ico","wb+")  
        tmp.write(base64.b64decode(img))#写入到临时文件中
        tmp.close()
        self.tk.iconbitmap("tmp.ico") #设置图标
        os.remove("tmp.ico") 
    
    #初始化菜单栏
    def init_menu(self):
        #　创建一个菜单栏，这里我们可以把它理解成一个容器，在窗口的上方
        menubar = tk.Menu(self.tk)
        #　定义一个空的菜单单元
        filemenu = tk.Menu(menubar, tearoff=0)  # tearoff意为下拉
        #　将上面定义的空菜单命名为`File`，放在菜单栏中，就是装入那个容器中
        menubar.add_cascade(label='菜单', menu=filemenu)
        #　在`文件`中加入`新建`的小菜单，即我们平时看到的下拉菜单，每一个小菜单对应命令操作。
        #　如果点击这些单元, 就会触发`do_job`的功能
        filemenu.add_command(label='切换为x86程序', command=self.change_main_exe)
        filemenu.add_command(label='帮助', command=self.help)
        filemenu.add_command(label='关于', command=self.about)
        # 分隔线
        filemenu.add_separator()
        filemenu.add_command(label='退出', command=self.quit)
        self.tk.config(menu=menubar)
    
    #销毁界面
    def quit(self):
        self.tk.destroy()
        
    #安装为服务
    def install_service(self):
        SW = WinSWGUI()
        if SW.isadmin():
            service_id = self.entry_id.get()
            service_name = self.entry_name.get()
            service_dersciption = self.entry_description.get()
            
            if  service_id and  service_dersciption and  service_name:
                self.save_config(service_id,  service_dersciption,  service_name)
                result = subprocess.run(['CMD', '/C',self.main_exe , 'install', self.main_exe_xml], shell=True, capture_output=True, text=True)
                if "successfully\" in result.stdout":
                    self.check_service_status()
                    messagebox.showinfo('安装系统服务', result.stdout)
                    time.sleep(2)
                    service_name = self.get_service_id()
                    resultstart = subprocess.run(['CMD', '/C', str('net start ' + service_name)], shell=True, capture_output=True, text=True)
                    self.check_service_status()
                    if resultstart.stdout:
                        messagebox.showinfo('启动系统服务', resultstart.stdout)
                    else:
                        messagebox.showinfo('提示', "服务安装成功但启动失败，请手动启动")
                else:
                    self.check_service_status()
                    messagebox.showinfo('安装系统服务', result.stdout)
            else:
                messagebox.showinfo('提示', '请将参数填写完整')
        else:
            messagebox.showinfo('错误', '无权限，请以管理员权限运行本程序')
    
    #卸载服务    
    def uninstall_service(self, event):
        # 读取id元素的文本内容
        service_id = self.get_service_id()
        if service_id:
            SW = WinSWGUI()
            if SW.isadmin():
                result = subprocess.run(['CMD', '/C', self.main_exe, 'uninstall', self.main_exe_xml], shell=True, capture_output=True, text=True)
                self.check_service_status()
                messagebox.showinfo('卸载系统服务', result.stdout)
                if "successfully" in result.stdout:
                    time.sleep(1)
                    resultstop = subprocess.run(['CMD', '/C', str('net stop '+ service_id)], shell=True, capture_output=True, text=True)
                    self.check_service_status()
                    if resultstop.stdout:
                        messagebox.showinfo('卸载系统服务', resultstop.stdout)
                    else:
                        messagebox.showinfo('提示', "服务卸载成功但还需要手动停止一次")
                else:
                    self.check_service_status()
                    messagebox.showinfo('卸载系统服务', result.stdout)
            else:
                messagebox.showinfo('提示', '未以管理员权限运行，请使用管理员权限运行')
        else:
            messagebox.showinfo('错误', '为能在服务中找到服务名，请查看帮助！')
    
    #卸载其他服务
    def uninstall_service_others(self, event):
        SW = WinSWGUI()
        if SW.isadmin():
            service_id = self.input_popup()
            print(service_id)
            if service_id:
                result = subprocess.run(['CMD', '/C', self.main_exe, 'uninstall', self.main_exe_xml], shell=True, capture_output=True, text=True)
                self.check_service_status()
                messagebox.showinfo('卸载系统服务', result.stdout)
                if "successfully" in result.stdout:
                    time.sleep(1)
                    resultstop = subprocess.run(['CMD', '/C', str('net stop '+ service_id)], shell=True, capture_output=True, text=True)
                    self.check_service_status()
                    messagebox.showinfo('停止系统服务', resultstop.stdout)
                else:
                    self.check_service_status()
                    messagebox.showinfo('卸载系统服务', result.stdout)          
            else:
                messagebox.showinfo('提示', '服务名称不能为空！')
        else:
             messagebox.showinfo('提示', '未以管理员权限运行，请使用管理员权限运行')
        
        #请求用户输入的弹窗
    def input_popup(self):
        result = simpledialog.askstring("卸载服务", "请输入需要卸载的服务名称:")
        if result:
            return result
        else:
            return False
    
        #查询是否已获取权限
    def setpermissionText(self):
        WS = WinSWGUI()
        if WS.isadmin():
            self.label_permission_status.configure(text="权限状态：管理员")
            print("1111")
        else:
            self.label_permission_status.configure(text="权限状态：无权限")
        
        
    #保存配置文件
    def save_config(self, service_id, service_name, service_dersciption):
        # 解析XML文件
        tree = ET.parse('WinSW.xml')
        root = tree.getroot()
        
        tmp_id = root.find('id')
        tmp_id.text = str(service_id)
        
        tmp_name = root.find('name')
        tmp_name.text = str(service_name)
        
        tmp_dersciption = root.find('description')
        tmp_dersciption.text = str(service_dersciption)
        
        print(self.label_file_path.cget('text'))
        tmp_excutable = root.find('executable')
        tmp_excutable.text = str(self.ptah_excutable)
    
        tree.write('WinSW.xml')
        
        
        #选择文件
    def select_file(self):
        file_path =askopenfilename(filetypes=[('应用程序','exe')])
        print('已选择文件地址：'+file_path)
        if self.is_chinese(file_path):
            self.label_file_path.configure(text="软件路径不可包含中文", fg = 'red')         
        elif file_path:
            excutable_name = os.path.basename(file_path)
            self.label_file_path.configure(text='已选择：' + excutable_name, fg = 'black')
            self.ptah_excutable = file_path
            #自动填写配置
            self.entry_id.insert(0,excutable_name.split('.')[0])
            self.entry_name.insert(0,excutable_name.split('.')[0])
            self.entry_description.insert(0,'a service for ' + excutable_name.split('.')[0])
        
        
        #检查整个字符串是否包含中文    
    def is_chinese(self, string):
        for ch in string:
            if '\u4e00' <= ch <= '\u9fa5':
                return True
        return False
    
        
        #查询服务状态
    def check_service_status(self):
        result = subprocess.run(['CMD', '/C', 'sc', 'query', self.get_service_id()], shell=True, capture_output=True, text=True)
        if 'SERVICE_NAME' in result.stdout and self.get_service_id() in result.stdout:
            if 'STATE' in result.stdout and 'RUNNING' in result.stdout:
                self.service_status_label.config(text='服务状态：运行中')
            else:
                self.service_status_label.config(text='服务状态：未运行')
        else:
            self.service_status_label.config(text='服务状态：未安装')
                
        #获取服务ID
    def get_service_id(self):
        # 解析XML文件
        tree = ET.parse('WinSW.xml')
        root = tree.getroot()
    
        # 读取id元素的文本内容
        service_id = root.find('id').text
        return service_id
        
        #修改winsw的程序和配置文件为x86构架    
    def change_main_exe(self):
        SW =WinSWGUI()
        sysinfo = SW.checksysplatform
        self.main_exe = 'WinSWx86.exe'
        self.main_exe_xml = 'WinSWx86.xml'
        messagebox.showinfo('提示', '主程序已切换为x86，重启可切换为x64')
            
        #帮助
    def help(self):
        messagebox.showinfo("帮助", "1.如果系统为32位可切换至x64程序。\n2.卸载程序只能直接卸载上一个程序\n3.鼠标右键单击卸载服务可填写服务名卸载对应服务")
        
        #关于
    def about(self):
        messagebox.showinfo("关于", "PWOER BY Tate Luo, CORE BY WINSW FROM GITHUB")
            
    #运行
gui = GUI()
gui.run()