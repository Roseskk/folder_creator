def name_list_data(data):
    name_list = []
    for i in range(2, data.max_row):
        name = data.cell(row=i, column=2).value
        if name is None:
            break
        else:
            name_list.append(name)
    return name_list
