import win32print
import win32api

def printer_run(file_name):
    #for fn in ['1.txt', '2.txt', '3.txt', '4.docx']:
        win32api.ShellExecute(0,\
                              'print',\
                              file_name,\
                              win32print.GetDefaultPrinterW(),\
                              ".",0)
