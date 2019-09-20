import printer_api as pa
import os 

def listdir(path, list_name): 
    for file in os.listdir(path): 
        file_path = os.path.join(path, file) 
        if os.path.isdir(file_path): 
            listdir(file_path, list_name) 
        elif os.path.splitext(file_path)[1]=='.docx': 
            list_name.append(file_path)

    return list_name

list_filename=[]
listdir('bill',list_filename)

for name in list_filename:
    print('Printer ' ,name.split(".")[0].split("_")[1])
    #pa.printer_run(name)
#print (listdir('bill',list_filename))
