def name_list_data(data):
    name_list = []
    for i in range(2, data.max_row+1):
        name = data.cell(row=i, column=2).value
        email = data.cell(row=i, column=3).value
        if name is None:
            break
        else:
            stud_permission = {'Name': f'{name}', "Email": f'{email}'}
            name_list.append(stud_permission)
    return name_list


def list_folder(drive, parent):
    filelist = []
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % parent}).GetList()
    for f in file_list:
        if f['mimeType'] == 'application/vnd.google-apps.folder':  # if folder
            filelist.append({"id": f['id'], "title": f['title'], })
        else:
            pass
    return filelist


def list_course(title_handler, profile, drive):
    title_handler_id = str()

    for name_profile in profile:
        if title_handler in name_profile.values():
            title_handler = name_profile
            title_handler_id = name_profile.get('id')

    course = list_folder(drive, title_handler_id)

    return course


def list_group_with_id(title_handler, course, drive):
    title_handler_id = str()

    for name_group in course:
        if title_handler in name_group.values():
            title_handler = name_group
            title_handler_id = name_group.get('id')

    group = list_folder(drive, title_handler_id)

    return {'group_id': f'{group}', 'title_id': f'{title_handler_id}'}


def create_group_folder(course_id, folder_name, drive):
    metadata = {
        'parents': [
            {"id": f'{course_id}'}
        ],
        'title': f'{folder_name}',
        'mimeType': 'application/vnd.google-apps.folder'
    }
    # Создание папки
    file = drive.CreateFile(metadata=metadata)
    while file.Upload():
        print('.', end='')


def create_student_folder(group_id, folder_name, data, drive):
    get_list = name_list_data(data[f'{folder_name}'])

    for credentials in get_list:

        name = credentials.get('Name')
        email = credentials.get('Email')

        metadata = {
            'parents': [
                {"id": f'{group_id}'}
            ],
            'title': f'{name}',
            'mimeType': 'application/vnd.google-apps.folder'
        }

        student = drive.CreateFile(metadata=metadata)

        while student.Upload():
            print('.', end='')

        student.InsertPermission({
                'type': 'user',
                'value': f'{email}',
                'role': 'writer'})

        print('.', end='')
