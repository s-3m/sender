import openpyxl


def get_data_dict(path):
    # C:\Users\samarenkora\Downloads\cert\VirtQuest.xlsx
    workbook = openpyxl.load_workbook(filename=f'{path}')
    sheet = workbook['Лист1']
    return {i[1].value: i[0].value for i in sheet.rows if i[1].value != 'ФИО'}

