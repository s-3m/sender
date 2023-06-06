import threading

import PySimpleGUI as sg

from sender import main


sg.theme('DarkTeal9')

btn_layout = [
    [sg.Button('Начать рассылку', key='start', button_color='green', size=(40, 2))],
]

layout = [
    [sg.Text('Данные отправителя рассылки:', font='Times_New_Roman 14')],
    [sg.Text('С какого ящика отправляем:', size=(22, 1)), sg.InputText(key='sender')],
    [sg.Text('Пароль:', size=(22, 1)), sg.InputText(key='password')],
    [sg.HorizontalSeparator(pad=(0, 20))],

    [sg.Text('Сообщение', font='Times_New_Roman 14')],
    [sg.Text('Тема сообщения:', size=(22, 1)), sg.InputText(key='subject')],
    [sg.Text('Текст сообщения:', size=(22, 1)), sg.Multiline(size=(43, 5), key='textbox')],

    [sg.Text('"Уважаемый(ая) <Имя>!" будет добавлено автоматически в начале текста. Вводите текст без этой фразы.', font='Times_New_Roman 8', text_color='grey')],
    [sg.HorizontalSeparator(pad=(0, 20))],

    [sg.Text('Файл с данными:', font='Times_New_Roman 14')],
    [sg.FileBrowse('Выбрать файл с данными', key='file'), sg.Text(key='file_path', pad=(0, 20))],

    [sg.Text('Название столбца с email получателя:', size=(22, 2)), sg.InputText(key='email_column')],
    [sg.Text('Название столбца с ФИО получателя:', size=(22, 2)), sg.InputText(key='name_column')],
    [sg.Text('ФИО в столбце обязательно должно иметь последовательность Ф -> И -> О', font='Times_New_Roman 8', text_color='grey')],
    [sg.HorizontalSeparator(pad=(0, 20))],

    [sg.Text('Папка с файлами для отправки:', font='Times_New_Roman 14')],
    [sg.Text('Файлы в папке должны быть названы в соответствии со столбцом ФИО из файла, который вы указали выше!', font='Times_New_Roman 8', text_color='grey')],
    [sg.FolderBrowse('Выбрать папку', key='folder'), sg.Text(key='file_path', pad=(0, 20))],

    [sg.Column(layout=btn_layout, element_justification='center', justification='center')],

    [sg.Text('Статус отправки:', font='Times_New_Roman 8', text_color='grey'), sg.Text('подготовка', key='status', font='Times_New_Roman 8', text_color='grey')],
    [sg.ProgressBar(max_value=20, orientation='h', size=(50, 10), key='progress')],
    [sg.Text('Отправлено:'), sg.Text(key='was_send')],
]

column_layout = [
    [sg.Column(layout=layout, scrollable=True, vertical_scroll_only=True, size=(None, 600))]
]

window = sg.Window('Рассылка', column_layout, finalize=True, size=(650, 700), resizable=True)
column_layout[0][0].expand(True, True)
progress_value = 0.1
window['progress'].update(progress_value)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Выход':
        break
    if event == 'wrong_name':
        window.extend_layout(column_layout[0][0], [[sg.Text('Сообщения не были отправлены:', size=(22, 2)), sg.Multiline(size=(50, 5), key='wrong_list_names', pad=(0, 20), auto_size_text=True)]])
        window['wrong_list_names'].update(values[event])
        window.size = (600, 900)
        sg.popup('Не все сообщения были отправлены.\nПосмотрите отчёт внизу основного окна.', title='Ошибка отправки', )
    if event == 'data_file_error':
        sg.popup(values[event], title='Ошибка')
    if event == 'wrong_server_login':
        sg.popup('Ошибка авторизации на сервере.\nПроверьте логин и пароль.', title='Ошибка авторизации')
        window['start'].update(disabled=False, button_color='green')
        window['status'].update('подготовка')
        window['was_send'].update('0')
    if event == 'start':
        if values['folder'] == '':
            sg.popup('Папка с файлами не выбрана.')
        elif values['email_column'] == '' or values['name_column'] == '':
            sg.popup('Заполните наименование колонок.')
        else:
            threading.Thread(target=main, args=(values, window), daemon=True).start()
            window['start'].update(disabled=True, button_color='grey')
    window['file_path'].update(values['file'])

window.close()
