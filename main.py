from PyPDF2 import PdfReader
import pandas
import glob
import os
import re


def iter_file(path_to_read, dt):
    for filename in glob.glob(os.path.join(path_to_read, '*.pdf')):
        page = read_file(filename)
        dt = extract_data_pdf(page, dt)
    return dt


def read_file(filename):
    with open(os.path.join(os.getcwd(), filename), 'rb') as fin:
        lt = list()
        for line in PdfReader(fin)._get_page(0).extract_text():
            lt.append(line)
        page = ''.join(lt)
    return page


def extract_data_pdf(page, dt):  #Нужна обработка ошибки NoneType на append
    s_temp = {}
    contract_num = re.search('\d{3}/\d{4}/\d{2}', page)
    contract_date = re.search('\d{2}(\s|.)\w{2,8}(\s|.)\d{4}', page)
    dt['Дата договора'].append(str(contract_date[0]))
    dt['Номер договора'].append(str(contract_num[0]))
    return dt


def save_data(dt, path_to_save):  #Создать добавление новых строк в файл
    return pandas.DataFrame(dt).to_excel(path_to_save)


def main(path_to_read, path_to_save):
    dt = {'Дата договора': [], 'Номер договора': []}
    iter_file(path_to_read, dt)
    return save_data(dt, path_to_save)


if __name__ == '__main__':
    ptr = r'D:\mypy\PDFConv2\testpdf'
    pts = r'D:\example.xlsx'
    main(ptr, pts)
