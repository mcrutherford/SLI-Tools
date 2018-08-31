"""

@file: parsejson.py
@author: Mark Rutherford [ mcr5801 ]
@created: 8/30/2018 8:54 AM
@description:
    

"""

from datetime import datetime
import json


def parse_json(file_location):
    file = open(file_location, 'r')
    json_data = json.loads(file.read())
    file.close()

    json_data['due date']['time'] = datetime.strptime(json_data['due date']['time'], '%b %d, %Y %I:%M %p')

    return json_data

def main():
    pass


if __name__ == "__main__":
    main()
    print("done")
