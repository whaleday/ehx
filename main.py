import os, sys, time, win32api, winreg, tempfile, tkinter, wget, zipfile
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
default_font = ("微软雅黑", 10)
edge_flag = False
def barProgress(current, total, width):
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
def getVersion(file):
    info = win32api.GetFileVersionInfo(file, os.sep)
    ms = info["FileVersionMS"]
    ls = info["FileVersionLS"]
    return f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
def tkPosition(root):
    X = int((root.winfo_screenwidth() - root.winfo_reqwidth()) / 2)
    Y = int((root.winfo_screenheight() - root.winfo_reqheight()) / 2)
    window.geometry(f"+{X}+{Y}")
window = tkinter.Tk()
window.attributes("-topmost", True)
window.withdraw()
try:
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Edge\\BLBeacon")
    edge_version = winreg.QueryValueEx(key, "version")[0]
    if not os.path.exists("msedgedriver.exe") or getVersion("msedgedriver.exe") != edge_version:
        url = "https://msedgedriver.azureedge.net/" + edge_version + "/edgedriver_win64.zip"
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
        wget.download(url, tempdir, bar=barProgress)
        window.withdraw()
        down_label.grid_forget()
        down_bar.grid_forget()
        if os.path.exists("msedgedriver.exe"):
            os.remove("msedgedriver.exe")
        driver_zip = zipfile.ZipFile(driver_path)
        driver_zip.extract("msedgedriver.exe", ".")
        driver_zip.close()
        os.remove(driver_path)
    if not os.path.exists("userdata"):
        window.title("提交登录信息")
        ttk.Label(window, text="用户名：", font=default_font).grid(row=0, column=0, padx=(5, 0), pady=(7, 5))
        ttk.Label(window, text="密码：", font=default_font).grid(row=1, column=0, padx=(5, 0))
        username = ttk.Entry(window, font=default_font)
        password = ttk.Entry(window, font=default_font)
        username.grid(row=0, column=1, columnspan=2, pady=(7, 5), ipady=2)
        password.grid(row=1, column=1, columnspan=2, ipady=2)
        check_var = tkinter.BooleanVar()
        check_style = ttk.Style()
        check_style.configure("custom.TCheckbutton", font=default_font)
        checkbox = ttk.Checkbutton(window, text="保存信息", onvalue=True, offvalue=False, variable=check_var, takefocus=False, style="custom.TCheckbutton")
        checkbox.grid(row=0, column=3, rowspan=2, padx=(3, 3), pady=(1, 0))
        button_style = ttk.Style()
        button_style.configure("custom.TButton", font=default_font)
        btn = ttk.Button(window, text="提交", command=tkClick, width=8, takefocus=False, style="custom.TButton")
        btn.grid(row=2, column=1, padx=(45, 0), pady=(5, 7))
        window.protocol("WM_DELETE_WINDOW", tkClose)
        window.resizable(False, False)
        window.deiconify()
        window.overrideredirect(False)
        tkPosition(window)
        window.mainloop()
        if check_var.get():
            file = open("userdata", "w")
            file.write(f"{username.get()}\n{password.get()}")
            file.close()
    edge_opt = Options()
    edge_opt.add_argument("--profile-directory=wdProf")
    browser = webdriver.Edge(service=Service("msedgedriver.exe"), options=edge_opt)
    edge_flag = True
    browser.get("https://www.ehuixue.cn/")
    browser.implicitly_wait(5)
    browser.find_element("xpath", "//span[@onclick='login()']").click()
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
    browser.implicitly_wait(0)
    while len(browser.find_elements("class name", "loginstyle")):
        time.sleep(0.5)
    browser.get("https://www.ehuixue.cn/index/personal/mystudy")
    while True:
        while True:
            handles = browser.window_handles
            if len(handles) == 2:
                browser.switch_to.window(handles[1])
                break
            time.sleep(0.5)
        href = browser.execute_script("return location.href;")
        browser.get("https://www.ehuixue.cn/index/study/inclass.html?cid=" + href.split("cid=")[1])
        browser.implicitly_wait(5)
        icon_list = browser.find_elements("xpath", "//span[@class='video']|//span[@class='pdf']")
        listlen = len(icon_list)
        video_list = [e.find_element("xpath", "../../td/div") for e in icon_list]
        dot_list = [e.find_element("xpath", "./span[2]") for e in video_list]
        count = 0
        while count < listlen:
            while count < listlen and dot_list[count].get_attribute("class") == "lr_status1":
                count += 1
            if count < listlen:
                ActionChains(browser).move_to_element(video_list[count]).click().perform()
                if icon_list[count].get_attribute("class") == "video":
                    time.sleep(1)
                    browser.find_element("id", "playercontainer").click()
                    while browser.execute_script("return isNaN(document.querySelector('video').duration);"):
                        time.sleep(0.5)
                    time.sleep(1)
                    browser.execute_script("document.querySelector('video').currentTime = document.querySelector('video').duration - 8;")
                    time.sleep(7.5)
                    study_end = browser.find_element("class name", "studyend")
                    while study_end.value_of_css_property("display") != "block":
                        time.sleep(0.5)
                    time.sleep(1)
                else:
                    browser.switch_to.frame(browser.find_element("id", "spdf"))
                    WebDriverWait(browser, 10).until(
                        expected_conditions.presence_of_element_located(("class name", "canvasWrapper")))
                    time.sleep(0.5)
                    pgdown_btn = browser.find_element("class name", "pageDown")
                    pdf_viewer = browser.find_element("id", "viewerContainer")
                    sc_origin = ScrollOrigin.from_element(pdf_viewer)
                    while pgdown_btn.is_enabled():
                        ActionChains(browser).scroll_from_origin(sc_origin, 0, 2000).perform()
                        time.sleep(1)
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