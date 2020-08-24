"""

@file: UnpackSubmissions.py
@author: Mark Rutherford [ mcr5801 ]
@created: 8/29/2018 9:06 PM
@description:
    Unpack student submissions to individual folders with personal xlsx grading files.

    Validated for Python 3.8, but should work for Python 3.x. Tkinter should come preinstalled with Python, unless you
    decided to not install it during the python install.

"""
import openpyxl  # pip install openpyxl
import openpyxl.utils  # pip install openpyxl
import zipfile
import tkinter as tk
from tkinter import filedialog
import os
import re
import json


def create_individual_xl(rubric_data, location, easy_access_vars):
    """
    Create an individual rubric xlsx for the student, and add automatic deductions.

    :param rubric_data: The loaded json data from the rubric
    :param location: The file in which to save the rubric
    :param easy_access_vars: Variables to be used for automatic deductions in the rubric
    """
    wb = openpyxl.Workbook()
    ws = wb['Sheet']

    ws.column_dimensions[openpyxl.utils.get_column_letter(1)].width = 30
    ws.column_dimensions[openpyxl.utils.get_column_letter(2)].width = 14
    ws.column_dimensions[openpyxl.utils.get_column_letter(3)].width = 10
    ws.column_dimensions[openpyxl.utils.get_column_letter(4)].width = 70

    col_names = ['Requirement', 'Points available', 'Deductions', 'Comments']
    for i in range(len(col_names)):
        ws.cell(1, i+1).value = col_names[i]
    row = 1
    for i, key in enumerate(rubric_data['requirements'].keys()):
        row = i+2
        ws.cell(row, 1).value = key
        ws.cell(row, 2).value = rubric_data['requirements'][key]["value"]

        # If there are automatic deductions, execute the code
        if "auto deductions" in rubric_data['requirements'][key]:
            deduct_sum = 0
            comment = ""
            for autodeduct in rubric_data['requirements'][key]['auto deductions']:
                # print(autodeduct['command'])
                if not eval(autodeduct['command']):
                    deduct_sum += autodeduct['penalty']
                    comment += ' -' + str(autodeduct['penalty']) + 'p: ' + autodeduct['desc'] + '\n'
            ws.cell(row, 3).value = deduct_sum
            ws.cell(row, 4).value = comment
    ws.cell(row+1, 1).value = "Additional Comments"

    wb.save(os.path.join(location, 'Rubric.xlsx'))


def parse_json(file_location):
    """
    Load a json file into a dictionary.

    :param file_location: The file to load
    :return: A dictionary of the loaded data
    """
    file = open(file_location, 'r')
    json_data = json.loads(file.read())
    file.close()

    return json_data


def main():
    if not os.path.exists('rubric.json'):
        print("No rubric.json file found. Aborting execution.")
        return
    rubric_data = parse_json('rubric.json')

    # Get the input and output file paths
    root = tk.Tk()
    root.withdraw()
    input_file_path = filedialog.askopenfilename()
    if input_file_path == '':
        return
    output_file_path = os.path.dirname(os.path.abspath(input_file_path))

    # Extract the main zip file
    zip_ref = zipfile.ZipFile(input_file_path, 'r')
    zip_ref.extractall(output_file_path)
    zip_ref.close()

    # Extract each student's zip file if possible
    for filename in os.listdir(output_file_path):
        if filename == 'index.html' or filename == os.path.basename(input_file_path):
            # these files are not student zip files, and won't be extracted
            continue
        try:
            # Identify the student's name from the mycourses file naming scheme
            pattern = re.compile("\d*-\d*\s-\s([\w|\s]*),\s(\w*)\s-\s(.*)\.zip")
            matches = pattern.match(filename)
            if not matches:
                raise zipfile.BadZipFile
            else:
                # Extract the student's zip file
                student_name = matches[1] + "_" + matches[2]
                print("Extracting:", student_name)
                zip_ref = zipfile.ZipFile(os.path.join(output_file_path, filename), 'r')
                zip_ref.extractall(os.path.join(output_file_path, student_name))
                zip_ref.close()

                # Remove the student's extracted zip file
                os.remove(os.path.join(output_file_path, filename))

                # Load variables for automatic grading later
                easy_access_vars = {
                    'zipname': matches[3] + '.zip',  # zip file name... Ex. "lab01.zip"
                    'filenames': os.listdir(os.path.join(output_file_path, student_name)),  # all file names... Ex. ["hw01.py", "debug.py", "QuestionAnswers.pdf"]
                }
                create_individual_xl(rubric_data, os.path.join(output_file_path, student_name), easy_access_vars)
        except zipfile.BadZipFile:  # if a file wasn't able to be extracted, log the error
            print("> Failed to extract", filename)


if __name__ == "__main__":
    main()
    print("\ndone")
