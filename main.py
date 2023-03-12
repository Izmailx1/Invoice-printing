import sys
import docx
from docx2pdf import convert
import win32print
import win32api
import PySimpleGUI as sg
from datetime import date
from docx.shared import Inches


def doc_edit(items_list):

    res = 0
    doc = docx.Document()
    doc.add_paragraph(f'Накладная от {date.today()}')

    table = doc.add_table(rows=len(items_list) + 1, cols=5)
    table.style = 'Table Grid'

    widths = (Inches(0.3), Inches(6), Inches(0.8), Inches(0.6), Inches(1.5))
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    cell = table.cell(0, 0)
    cell.width = Inches(0.1)
    cell.text = '№'
    cell = table.cell(0, 1)
    cell.width = Inches(2)
    cell.text = 'Наименование'
    cell = table.cell(0, 2)
    cell.text = 'Кол-во'
    cell = table.cell(0, 3)
    cell.text = 'Цена'
    cell = table.cell(0, 4)
    cell.text = 'Сумма'

    for row in range(1, len(items_list) + 1):
        cell = table.cell(row, 0)
        cell.text = str(row)
        cell = table.cell(row, 1)
        cell.text = items_list[row - 1][0]
        cell = table.cell(row, 2)
        cell.text = str(items_list[row - 1][1])
        cell = table.cell(row, 3)
        cell.text = str(items_list[row - 1][2])
        cell = table.cell(row, 4)
        cell.text = str(int(items_list[row - 1][1]) * int(items_list[row - 1][2]))
        res += int(items_list[row - 1][1]) * int(items_list[row - 1][2])

    doc.add_paragraph(f'Итого: {res}руб')
    doc.save('files/doc.docx')

def print_document(num_of_lists: str):
    convert('files/doc.docx')
    name = 'HP LaserJet Professional M1132 MFP'
    mode = 1
    input_pdf = 'files\doc.pdf'
    win32print.SetDefaultPrinterW(name)
    win32print.SetDefaultPrinter(name)
    printdefaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
    handle = win32print.OpenPrinter(name, printdefaults)
    level = 2
    attributes = win32print.GetPrinter(handle, level)
    attributes['pDevMode'].Duplex = mode

    win32print.SetPrinter(handle, level, attributes, 0)
    win32print.GetPrinter(handle, level)['pDevMode'].Duplex

    for i in range(int(num_of_lists)):
        win32api.ShellExecute(1, 'print', input_pdf, '.', '/manualstoprint', 0)

    win32print.ClosePrinter(handle)


with open('files/list_items.txt', 'r', encoding='UTF-8') as f:
    list_item = f.read()
list_items = list_item.split('\n')

print_items = []
res_s = ''

layout = [
    [sg.Text('Выбери наименование', size=(20, 1), font='Lucida', justification='left')],
    [sg.Combo(list_items, default_value=list_items[0], key='_list_items_')],
    [sg.Text('Введи количество', size=(20, 1), font='Lucida', justification='left')],
    [sg.InputText(key='_quantity_')],
    [sg.Text('Введи цену', size=(20, 1), font='Lucida', justification='left')],
    [sg.InputText(key='_price_')],
    [sg.Button('Добавить в накладную', font=('Times New Roman', 12), key='_add_')],
    [sg.Text('Ваша накладная', size=(20, 1), font='Lucida', justification='left')],
    [sg.Output(size=(88, 20), key='_output_')],
    [sg.Button('Сбросить', font=('Times New Roman', 12), key='_reset_'),
     sg.Button('Печатать накладную', font=('Times New Roman', 12), key='_print_')]
]


window = sg.Window('File Compare', layout)

while True:
    event, values = window.read()
    if event in (None, 'Exit', 'Cancel'):
        break
    if event == '_add_':
        pos = []
        pos.append(values['_list_items_'])
        s = 'Наим.: '
        s += pos[-1]
        s += ' '
        pos.append(values['_quantity_'])
        s += 'Кол-во: '
        s += pos[-1]
        s += ' '
        pos.append(values['_price_'])
        s += 'Цена: '
        s += pos[-1]
        print_items.append(pos)
        window['_list_items_'].update('')
        window['_quantity_'].update('')
        window['_price_'].update('')
        res_s += s
        res_s += '\n'
        print(s)

    if event == '_reset_':
        window['_output_'].update('')
        print_items = []
        res_s = ''

    if event == '_print_':
        doc_edit(print_items)
        sg.popup('Печать упаковочного листа', res_s)
        print_document(1)

