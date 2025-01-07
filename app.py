import openpyxl
#ler planilha e guardar as informacoes relevantes
file_name = 'contatos_whatsapp.xlsx'
work_book = openpyxl.load_workbook(file_name)

#especificando o nome da planilha que iremos usar
spreadsheet_name = 'Página1'
spreadsheet = work_book[spreadsheet_name]

for line in spreadsheet.iter_rows(min_row=2):
    name = line[0].value
    number = str(line[1].value)
    reason = line[2].value

    print(name)
    print(number)
    print(reason)