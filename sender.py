import smtplib
import time

import pandas as pd
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart


def send_email(message, name, to, data_dict, window):
    sender = data_dict['sender']
    psw = data_dict['password']

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    # try:
    server.login(sender, psw)
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = to
    msg['Subject'] = data_dict['subject']

    msg.attach(MIMEText(message))
    # try:
    with open(f'{data_dict["folder"]}\\{name}.pdf', 'rb') as f:
        file = MIMEApplication(f.read())
    # except FileNotFoundError as _ex:
    #     print(f'Файла {name}.pdf не существует')
    #     window.write_event_value('wrong_file_attach', name)
    #     return

    file.add_header('content-disposition', 'attachment', filename=f'{name}.pdf')
    msg.attach(file)

    server.sendmail(sender, to, msg.as_string())

    # except Exception as _ex:
    #     window.write_event_value('wrong_server_login')
    #     return


def main(data_dict, window):
    try:
        data_frame = pd.read_excel(data_dict['file'], index_col=False)
    except FileNotFoundError:
        window.write_event_value('data_file_error', 'Выберите файл с данными.')
        window['start'].update(disabled=False, button_color='green')
        exit()
    d = data_frame.filter([data_dict['email_column'].strip(), data_dict['name_column'].strip()])
    progress_value = 0
    progress_count = 0
    error_names = ''

    for i, r in d.iterrows():
        try:
            message = f'Уважаемый(ая) {str(r[1]).split(" ")[1]}!\n{data_dict["textbox"]}'
            window['status'].update('идёт отправка')
            window['was_send'].update(f'{progress_count} из {len(d.values)}')
            send_email(message, r[1], r[0], data_dict, window)

        except IndexError:
            error_names += f'{r[1]} - неправильный формат ФИО\n'
        except FileNotFoundError as _ex:
            error_names += f'{r[1]}.pdf - такого файла нет в директории\n'
        except Exception as _ex:
            window.write_event_value('wrong_server_login', 'err')
            exit()

        progress_value += 20 / len(d.values)
        window['progress'].update(progress_value)
        progress_count += 1

    if error_names != '':
        window['was_send'].update(background_color='red')
        window['was_send'].update('Не все сообщения были отправлены. Детали ниже.')
        window['status'].update('Завершено с ошибками!', text_color='red')
    else:
        window['was_send'].update('Cообщения отправлены!')
        window['status'].update('Завершено!')

    window['start'].update(disabled=False, button_color='green')
    window.write_event_value('wrong_name', error_names)
