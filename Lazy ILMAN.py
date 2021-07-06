image_dict = dict()
file_name = str()
xW = float()
yH = float()

def find_insert(hwp, direction, name, msg):
    import os

    global xW,yH

    hwp.MovePos(2, 0, 0)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    option = hwp.HParameterSet.HFindReplace
    option.FindString = msg
    option.UseWildCards = 1
    option.IgnoreMessage = 1
    option.Direction = hwp.FindDir("Forward")
    option.FindType = False

    while(1):
        if hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
            print("INSERT IMG : ", name, ", ", msg, ", ", xW, ", ", yH)
            
            hwp.InsertPicture(os.path.join(direction, name), Embedded=True, Width=xW, Height=yH, sizeoption=1)
        else:
            break


def start():
    import win32com.client as win32

    global image_dict, window, file_name, xW, yH

    if not bool(file_name):
        from tkinter import messagebox as msg
        msg.showwarning('메시지 알림', 'HWP 파일을 넣어주세요!!')
        if not bool(image_dict):  
            msg.showwarning('메시지 알림', '이미지 파일을 넣어주세요!!') 
        if not bool(xW) or not bool(yH):
            msg.showwarning('메시지 알림', '이미지 크기를 설정하세요!!') 
            return
        return
    else:
        from tkinter import messagebox as msg
        if not bool(image_dict):  
            msg.showwarning('메시지 알림', '이미지 파일을 넣어주세요!!')
            if not bool(xW) or not bool(yH):
                msg.showwarning('메시지 알림', '이미지 크기를 설정하세요!!') 
                return 
            return

    window.destroy()
    
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.Open(file_name, "HWP", "forceopen:true")

    print(image_dict)
    img_direction = image_dict['direction']
    img_list = image_dict['name']

    for index, img_name in enumerate(img_list):
        find_insert(hwp, img_direction, img_name, img_name)
        find_insert(hwp, img_direction, img_name, "$"+str(index+1)+"$")
    # hwp.Clear(3)
    # hwp.Quit()


'''
def input_name():
    global entry

    name = str(entry.get())
    entry.delete(0, len(name))

    start(name)
'''

def get_xy():
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.title("이미지 크기 조정")
    root.geometry("250x100") 
    root.resizable(False, False)  
    #root.iconbitmap('./.ico/ilman.ico')

    tk.Label(root, text="너비(mm) : ", width=15).grid(row=0, column=0, pady=5)
    Xentry = tk.Entry(root, width=15)
    Xentry.grid(row=0, column=1, pady=5)

    tk.Label(root, text="높이(mm) : ", width=15).grid(row=1, column=0)
    Yentry = tk.Entry(root, width=15)
    Yentry.grid(row=1, column=1)

    def get_WH():
        global xW, yH
        xW = Xentry.get()
        yH = Yentry.get()

        try:
            if int(xW) or float(xW):
                print("okok")
                root.destroy()
            elif int(yH) or float(yH):
                print("okok")
                root.destroy()
        except(ValueError):
            import tkinter.messagebox
            tkinter.messagebox.showerror(    "메시지 알림", "  0보다 큰 숫자를 입력하세요!!\n\n    -----다시 시도하세요-----")
    
    button = tk.Button (root, text='확인',command=get_WH, width=20)
    button.grid(row=2, column=0, columnspan=2, pady=10)
    
    root.mainloop()


def get_imgList():
    from tkinter.filedialog import askopenfilenames

    global image_dict

    imagelist = askopenfilenames(
        title='이미지를 선택하세요', filetypes=[('all files', '.*')])

    for i in imagelist:
        if i[-4:] == '.jpg':
            pass
        elif i[-4:] == '.png':
            pass
        else:
            import tkinter.messagebox
            tkinter.messagebox.showerror(
                "메시지 알림", "  jpg 혹은 png 파일만 가능합니다!!\n\n    -----다시 시도하세요-----")
            return

    try:
        direc_list = imagelist[0].rsplit("/", maxsplit=1)[0]
        image_list = [i.rsplit("/", maxsplit=1)[1] for i in imagelist]
        image_dict = {'direction': direc_list, 'name': image_list}

        from tkinter import messagebox
        messagebox.showinfo(title="이미지 목록", message=str(image_list)+"\n")

    except IndexError:
        return


def get_file():
    from tkinter.filedialog import askopenfilenames

    global file_name

    hwpFile = askopenfilenames(title='한글 파일을 선택하세요', filetypes=[
                               ('hwp files', '*.hwp')])

    try:
        file_name = hwpFile[0]
    except IndexError:
        return


if __name__ == '__main__':
    import tkinter as tk
    
    window = tk.Tk()
    window.title("Lazy ILMAN")
    window.geometry("400x300+50+50")
    window.resizable(False, False)
    #window.iconbitmap('./ilman.ico')

    text = tk.Label(window, text="이 프로그램은 아래한글에 사진을 \n자동으로 첨부해주는 자동문서화 프로그램입니다.\n")
    text.pack(pady="10")

    btn1 = tk.Button(window, text="hwp 파일 선택", command=get_file, width=30)
    btn1.pack(side="top", pady="5")

    btn2 = tk.Button(window, text="이미지 선택", command=get_imgList, width=30)
    btn2.pack(side="top", pady="5")

    btn3 = tk.Button(window, text="이미지 크기 설정", command=get_xy, width=30)
    btn3.pack(side="top", pady="5")

    btn4 = tk.Button(window, text="실행하기", command=start, width=30, height=2)
    btn4.pack(side="top", pady="5")

    text2 = tk.Label(window, text="만든 사람 : Youtube 일만이", height=2)
    text2.pack(side="top", pady="10")

    window.mainloop()
