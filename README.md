# SLI-Tools

### Python requirements

These scripts have been validated for Python 3.8, though any Python 3.x version should work. They 
rely on Tkinter, which should be included in your python install by default, unless you specifically 
chose to not install it.

### Dependencies
openpyxl - This library is not default in a python installation, and can be added via pip with the command "pip 
install openpyxl"

### How to grade using SLI-Tools
1. Create a grading rubric to be used and name it rubric.json in the same folder as 
[UnpackSubmissions.py](./UnpackSubmissions.py). There is a provided example [ExampleRubric.json](./ExampleRubric.json) 
from an assignment a previous year
2. Download the .zip of students' assignments from MyCourses and place into an empty folder
3. Run [UnpackSubmissions.py](./UnpackSubmissions.py) and select the downloaded .zip
4. Grade the assignments by marking how many points are deducted in each section and add comments. No value in the 
deduction column is counted as no point loss, comments are optional
5. Run [CombineGrades.py](./CombineGrades.py) and select the main folder. It will combine all grading 
sheets into one xlsx named AllStudents.xlsx
6. Copy the grade and comments into the MyCourses grading page. NOTE: You may need to double click 
each comment in the master xlsx in order to properly copy the formatting (such as line breaks). 
This is a limitation of excel.