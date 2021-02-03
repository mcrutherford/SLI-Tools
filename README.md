# SLI-Tools

### Python requirements
These scripts have been validated for Python 3.8, though any Python 3.x version should work. They have been tested on 
Windows and Linux (Ubuntu).

### Dependencies
- openpyxl - This library is not included in a standard python installation, and can be added via pip with the command 
  '*pip install openpyxl*' or by running '*python [setup.py](./setup.py) install*'
- tkinter - This should be included in Windows python installs. If running on linux, you may need to run '*apt-get install 
  python3-tk*'

### How to grade using SLI-Tools
1. Create a json grading rubric. There is a provided example [ExampleRubric.json](./ExampleRubric.json) from an 
   assignment a previous year
2. Download the .zip of student assignments from MyCourses and place into an empty folder
3. Run [UnpackSubmissions.py](./UnpackSubmissions.py) and select the downloaded .zip. Then select the json rubric you 
   would like to grade against
4. Grade the assignments by marking how many points are deducted in each requirement and add comments. No value in the 
   deduction column is counted as no point loss, comments are optional
5. Run [CombineGrades.py](./CombineGrades.py) and select the main folder. It will combine all grading sheets into one 
   Excel file named AllStudents.xlsx
6. Copy the grade and comments into the MyCourses grading page. NOTE: You may need to double-click each comment in the 
   master xlsx in order to properly copy the formatting (such as line breaks). This is a limitation of Excel.