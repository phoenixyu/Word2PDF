#-*- coding:utf-8 -*-
# Requires Office Word
# Requires python for win32 extension

import os
from win32com import client as wc 
import win32file
import win32con

def word2pdf(word, pwd, in_file):
    try:
        doc = word.Documents.Open(pwd+"/"+in_file) 
    except:
        print("%s open Error!" % (in_file))
        return False
    else:
        try:
            doc.SaveAs(pwd+"/pdfs/"+in_file.split(".")[0]+".pdf", 17)
        except:
            print("%s convert Error!" % (in_file))
            return False
        return True
        doc.Close() 

def checkExtensiton(pwd, in_file):
    name = in_file.split(".")[-1]
    file_flag = win32file.GetFileAttributesW(pwd+"/"+in_file)
    is_hiden = file_flag & win32con.FILE_ATTRIBUTE_HIDDEN
    
    if is_hiden == False:
        if name == "doc" or name == "docx":
            return True
        else:
            return False
    else:
        return False

if __name__ == '__main__':
    # Open Word
    try:
        word = wc.Dispatch("Word.Application") 
    except:
        print("Please install or check Microsoft Word!")
    else:
        pwd = os.getcwd()
        if os.path.exists(pwd+"/pdfs") is not True:
            os.mkdir(pwd+"/pdfs")
        else:
            file_list = os.listdir(pwd)
            for item_name in file_list:
                flag = checkExtensiton(pwd, item_name)
                if flag:
                    print("Converting %s now, please wait." % (item_name))
                    res = word2pdf(word, pwd, item_name)
                    if res:
                        print("Convert Success!")
                    else:
                        print("Convert Failed! An Error occurred!")