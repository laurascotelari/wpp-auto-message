import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

#ler planilha e guardar as informacoes relevantes
file_name = 'contatos_whatsapp.xlsx'
work_book = openpyxl.load_workbook(file_name)

#especificando o nome da planilha que iremos usar
spreadsheet_name = 'Página1'
spreadsheet = work_book[spreadsheet_name]

for line in spreadsheet.iter_rows(min_row=2):
    name = line[0].value
    number = str(int(line[1].value))
    reason = line[2].value
    message = f"Oi, {name}! Se você recebeu essa mensagem é porque você é {reason} :)"

    #criar links personalizados
    link_message_wpp = f'https://web.whatsapp.com/send?phone={number}&text={quote(message)}'
    print(link_message_wpp)

    #é preciso estar logado no whatsapp web
    webbrowser.open(link_message_wpp)

