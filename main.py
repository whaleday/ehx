import os, sys, time, win32api, winreg, tempfile, tkinter, wget, zipfile
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.common import NoSuchElementException
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
default_font = ("微软雅黑", 10) #ttk全局字体样式
base_time = 0.5 #sleep时间，网络不佳时应上调该时间
edge_flag = False
edge_path = "msedgedriver.exe"
def barProgress(current, total, width): #进度条更新
    percentage = int(current / total * 100)
    bar_var.set(percentage)
    label_var.set("正在下载webdriver %" + str(percentage))
    window.update()
def tkClose():
    window.withdraw()
    if edge_flag:
        browser.quit()
    sys.exit()
def tkClick():
    if username.get() != "" and password.get() != "":
        window.withdraw()
        window.quit()
def getVersion(file): #获取driver版本
    file_version = win32api.GetFileVersionInfo(file, os.sep)
    ms = file_version["FileVersionMS"]
    ls = file_version["FileVersionLS"]
    return f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
def tkPosition(root): #窗口居中
    X = int((root.winfo_screenwidth() - root.winfo_reqwidth()) / 2)
    Y = int((root.winfo_screenheight() - root.winfo_reqheight()) / 2)
    window.geometry(f"+{X}+{Y}")
def browserWait():
    i = 1
    while i <= 4:
        try:
            browser.implicitly_wait(2* i * base_time)
            break
        except NoSuchElementException as err:
            i *= 2
window = tkinter.Tk()
window.attributes("-topmost", True)
window.withdraw()
del_var = tkinter.BooleanVar()
try:
    edge_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Edge\\BLBeacon") #从注册表获取edge版本信息
    edge_version = winreg.QueryValueEx(edge_key, "version")[0]
    if not os.path.exists(edge_path) or getVersion(edge_path) != edge_version: #下载driver
        down_url = "https://msedgedriver.azureedge.net/" + edge_version + "/edgedriver_win64.zip"
        tempdir = tempfile.gettempdir()
        driver_path = tempdir + "/edgedriver_win64.zip"
        window.deiconify()
        label_var = tkinter.StringVar()
        bar_var = tkinter.IntVar()
        down_label = ttk.Label(window, font=default_font, textvariable=label_var)
        down_bar = ttk.Progressbar(window, length=250, maximum=100, variable=bar_var)
        down_label.grid(row=0, column=0, pady=(1, 3))
        down_bar.grid(row=1, column=0, padx=5, pady=(0, 5))
        window.overrideredirect(True)
        tkPosition(window)
        wget.download(down_url, tempdir, bar=barProgress)
        window.withdraw()
        down_label.grid_forget()
        down_bar.grid_forget()
        if os.path.exists(edge_path):
            os.remove(edge_path)
        driver_zip = zipfile.ZipFile(driver_path)
        driver_zip.extract(edge_path, ".")
        driver_zip.close()
        os.remove(driver_path)
    if not os.path.exists("userdata"): #登录部分
        window.title("提交登录信息")
        ttk.Label(window, text="用户名：", font=default_font).grid(row=0, column=0, padx=(5, 0), pady=(7, 5))
        ttk.Label(window, text="密码：", font=default_font).grid(row=1, column=0, padx=(5, 0))
        username = ttk.Entry(window, font=default_font)
        password = ttk.Entry(window, font=default_font)
        username.grid(row=0, column=1, columnspan=2, pady=(7, 5), ipady=2)
        password.grid(row=1, column=1, columnspan=2, ipady=2)
        chkbox_style = ttk.Style()
        chkbox_style.configure("custom.TCheckbutton", font=default_font)
        info_var = tkinter.BooleanVar()
        info_chkbox = ttk.Checkbutton(window, text="保存信息", onvalue=True, offvalue=False, variable=info_var, takefocus=False, style="custom.TCheckbutton")
        info_chkbox.grid(row=0, column=3, padx=(5, 0), pady=(1, 0), sticky="w")
        del_var.set(True)
        del_chkbox = ttk.Checkbutton(window, text="保留driver", onvalue=True, offvalue=False, variable=del_var, takefocus=False, style="custom.TCheckbutton")
        del_chkbox.grid(row=1, column=3, padx=(5, 5), pady=(1, 0), sticky="w")
        btn_style = ttk.Style()
        btn_style.configure("custom.TButton", font=default_font)
        btn = ttk.Button(window, text="提交", command=tkClick, width=8, takefocus=False, style="custom.TButton")
        btn.grid(row=2, column=1, padx=(45, 0), pady=(5, 7))
        window.protocol("WM_DELETE_WINDOW", tkClose)
        window.resizable(False, False)
        window.deiconify()
        window.overrideredirect(False)
        tkPosition(window)
        window.mainloop()
        if info_var.get():
            file = open("userdata", "w")
            file.write(f"{username.get()}\n{password.get()}")
            file.close()
    edge_opt = Options() #更改edge的启动配置，防止与当前用户配置冲突
    edge_opt.add_argument("--profile-directory=driver_tmp_prof")
    browser = webdriver.Edge(service=Service(edge_path), options=edge_opt)
    edge_flag = True
    browser.get("https://www.ehuixue.cn/")
    browserWait()
    browser.find_element("xpath", "//span[@onclick='login()']").click() #自动填写登录信息
    browser.switch_to.frame(browser.find_element("id", "layui-layer-iframe100001"))
    if os.path.exists("userdata"):
        file = open("userdata", "r")
        data = file.read().split("\n")
        browser.find_element("id", "account1").send_keys(data[0])
        browser.find_element("id", "pwd").send_keys(data[1])
        file.close()
    else:
        browser.find_element("id", "account1").send_keys(username.get())
        browser.find_element("id", "pwd").send_keys(password.get())
    browser.switch_to.default_content()
    browserWait()
    while len(browser.find_elements("class name", "loginstyle")):
        time.sleep(base_time)
    browser.get("https://www.ehuixue.cn/index/personal/mystudy") #选课
    while True:
        while True:
            handles = browser.window_handles
            if len(handles) == 2:
                browser.switch_to.window(handles[1])
                break
            time.sleep(base_time)
        href = browser.execute_script("return location.href;")
        browser.get("https://www.ehuixue.cn/index/study/inclass.html?cid=" + href.split("cid=")[1])
        browserWait()
        icon_list = browser.find_elements("xpath", "//span[@class='video']|//span[@class='pdf']") #根据小图标判断任务类型-视频/文档
        listlen = len(icon_list)
        video_list = [e.find_element("xpath", "../../td/div") for e in icon_list]
        dot_list = [e.find_element("xpath", "./span[2]") for e in video_list] #小蓝点列表，用于判断课程完成情况
        count = 0
        while count < listlen:
            while count < listlen and dot_list[count].get_attribute("class") == "lr_status1":
                count += 1
            if count < listlen:
                ActionChains(browser).move_to_element(video_list[count]).click().perform()
                if icon_list[count].get_attribute("class") == "video": #刷视频部分
                    time.sleep(2 * base_time)
                    browser.find_element("id", "playercontainer").click()
                    while browser.execute_script("return isNaN(document.querySelector('video').duration);"): #等待视频加载
                        time.sleep(base_time)
                    time.sleep(2 * base_time)
                    browser.execute_script("document.querySelector('video').currentTime = document.querySelector('video').duration - 8;") #**核心代码**
                    time.sleep(7.5)
                    studyend = browser.find_element("class name", "studyend")
                    while studyend.value_of_css_property("display") != "block":
                        time.sleep(base_time)
                    time.sleep(2 * base_time)
                else:
                    browser.switch_to.frame(browser.find_element("id", "spdf"))
                    WebDriverWait(browser, 10).until(
                        expected_conditions.presence_of_element_located(("class name", "canvasWrapper"))) #等待pdf画布加载
                    time.sleep(base_time)
                    pgdown_btn = browser.find_element("class name", "pageDown") #根据翻页键是否可用判断pdf是否完成阅读
                    pdf_viewer = browser.find_element("id", "viewerContainer")
                    scr_origin = ScrollOrigin.from_element(pdf_viewer)
                    while pgdown_btn.is_enabled(): #模拟鼠标滚动翻pdf
                        ActionChains(browser).scroll_from_origin(scr_origin, 0, 2000).perform()
                        time.sleep(2 * base_time)
                    browser.switch_to.default_content()
                count += 1
        msg_ret = tkinter.messagebox.askyesno("提示", "该课程已完成，是否继续其它课程？")
        if msg_ret:
            browser.close()
            browser.switch_to.window(handles[0])
        else:
            break
except Exception as errinfo:
    tkinter.messagebox.showwarning("提示", "程序错误！\n" + str(errinfo).split("\n  (Session")[0])
finally:
    window.destroy()
    if edge_flag:
        browser.quit()
    if os.path.exists(edge_path) and not del_var.get():
        os.remove(edge_path)