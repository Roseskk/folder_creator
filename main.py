import pydrive2
import traceback
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from openpyxl import load_workbook

import utils as utl

from pathlib import Path


gauth = GoogleAuth()
gauth.LocalWebserverAuth()


drive = GoogleDrive(gauth)

# Create path choice
path_to_file = input('Path to file: ')
try:
    if Path(path_to_file).exists():
        data = load_workbook(f'{Path(path_to_file)}')

        drive_id = input('Write id of the GoogleDrive -->:')

        for name in data.sheetnames:
            print(f'Group {name}')

        group = input('Name of group ->: ')

        while group not in data.sheetnames:
            group = input('Name of group ->: ')

        get_list = utl.name_list_data(data[f'{group}'])
        print(get_list)
        for name in get_list:
            metadata = {
                'parents': [
                    {"id": f'{drive_id}'}
                ],
                'title': f'{name}',
                'mimeType': 'application/vnd.google-apps.folder'
            }
            # Create file
            file = drive.CreateFile(metadata=metadata)
            # permission = file.GetPermissions('1VWKbAHt8MM1vQX-9xCt5xUERdVmh3mK5')
            # print(permission)
            file.Upload()
            permission = file.InsertPermission({
                'type': 'user',
                'value': 'mixan9000@gmail.com',
                'role': 'reader'})
            print('.', end='')
    else:
        print('Incorrect path')
    # file1 = drive.CreateFile({'id':'Hello.txt'})
    # Fetch permissions.
    # permissions = file1.GetPermissions()
    # print(permissions)

except Exception as e:
    print('Wrong ID of Google Drive!', traceback.print_exc())


# /Users/xxrose/PycharmProjects/toggl/firstCourse.xlsx
# 176ZDkmDvQ-KY8Zcx7nBcWLdR22_F0Kg_




