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

import parsejson
import parsexl


def main():
    rubric_data = parsejson.parse_json('rubric.json')

    root = tk.Tk()
    root.withdraw()

    input_file_path = filedialog.askopenfilename()
    if input_file_path == '':
        return
    output_file_path = os.path.dirname(os.path.abspath(input_file_path))

    zip_ref = zipfile.ZipFile(input_file_path, 'r')
    zip_ref.extractall(output_file_path)
    zip_ref.close()

    for filename in os.listdir(output_file_path):
        if filename.endswith(".zip"):
            pattern = re.compile("\d*-\d*\s-\s(\w*),\s(\w*)\s-\s(.*)\.zip")
            matches = pattern.match(filename)
            if matches != None:
                easy_access_vars = {}
                easy_access_vars['lastname'] = matches[1]
                easy_access_vars['first'] = matches[2]
                easy_access_vars['zipname'] = matches[3]
                zip_ref = zipfile.ZipFile(output_file_path+"/"+filename, 'r')
                print(output_file_path + "/" + easy_access_vars['lastname'] + "_" + easy_access_vars['first'])
                zip_ref.extractall(output_file_path + "/" + easy_access_vars['lastname'] + "_" + easy_access_vars['first'])
                zip_ref.close()

                os.remove(output_file_path+"/"+filename)

                easy_access_vars['filenames'] = []
                easy_access_vars['filelocations'] = []
                for fname in os.listdir(output_file_path + "/" + easy_access_vars['lastname'] + "_" + easy_access_vars['first']):
                    if fname.endswith('.py'):
                        easy_access_vars['filenames'].append(fname)
                        easy_access_vars['filelocations'].append(output_file_path + \
                                                                 "/" + easy_access_vars['lastname'] + "_" + easy_access_vars['first'] + \
                                                                 "/" + fname)

                parsexl.create_individual_xl(rubric_data, output_file_path + "/" + easy_access_vars['lastname'] + "_" + easy_access_vars['first'], easy_access_vars)


if __name__ == "__main__":
    main()
    print("done")
