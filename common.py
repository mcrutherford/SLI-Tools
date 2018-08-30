"""

@file: common.py
@author: Mark Rutherford [ mcr5801 ]
@created: 8/29/2018 9:06 PM
@description:
    

"""
import zipfile
import tkinter as tk
from tkinter import filedialog
import os
import re


def main():
    root = tk.Tk()
    root.withdraw()

    input_file_path = filedialog.askopenfilename()
    if input_file_path == '':
        return
    output_file_path = filedialog.askdirectory()
    if output_file_path == '':
        return

    zip_ref = zipfile.ZipFile(input_file_path, 'r')
    zip_ref.extractall(output_file_path)
    zip_ref.close()

    for filename in os.listdir(output_file_path):
        if filename.endswith(".zip"):
            pattern = re.compile("\d*-\d*\s-\s(\w*),\s(\w*)\s-\s(.*)\.zip")
            matches = pattern.match(filename)
            if matches != None:
                last = matches[1]
                first = matches[2]
                zipname = matches[3]
                # print("[" + last + "][" + first + "][" + labname+"]")
                zip_ref = zipfile.ZipFile(output_file_path+"/"+filename, 'r')
                zip_ref.extractall(output_file_path + "/" + last + "_" + first)
                zip_ref.close()

                os.remove(output_file_path+"/"+filename)

if __name__ == "__main__":
    main()
    print("done")
