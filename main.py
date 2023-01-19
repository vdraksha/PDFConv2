from PyPDF2 import PdfReader
import pandas
import glob
import os
import re


def iter_file(path_to_read, dt):
    for filename in glob.glob(os.path.join(path_to_read, '*.pdf')):
        read_file(filename, dt)
    return dt


def read_file(filename, dt):
    with open(os.path.join(os.getcwd(), filename), 'rb') as fin:
        reader = PdfReader(fin)
        page = reader.pages[0].extract_text()
        extract_data_pdf(page, dt)
    return dt


def extract_data_pdf(page, dt):
    s_temp = {'Дата договора': r'\d{2}( \w{2,8} |\.\d{2}\.)\d{2,4}',
              'Номер договора': r'\d{3}/\d{4}/\d{2}',
              'Контрагент': r'(Общество с ограниченной ответственностью|ООО) \"\w+\"'}  #Не находит, хотя re рабочее
    for key in s_temp:
        val = re.search(s_temp[key], page)
        if val is None:
            val = 'Not Found'
            dt[key].append(val)
        else:
            dt[key].append(val[0])
    return dt


def save_data(dt, path_to_save):
    df_new = pandas.DataFrame(dt)
    df_old = pandas.read_excel(path_to_save)
    df = pandas.concat([df_old, df_new], ignore_index=True)
    return df.to_excel(path_to_save, index=False)


def main(path_to_read, path_to_save):
    dt = {'Дата договора': [], 'Номер договора': [], 'Контрагент': []}
    iter_file(path_to_read, dt)
    return save_data(dt, path_to_save)


if __name__ == '__main__':
    ptr = r'D:\mypy\PDFConv2\testpdf'
    pts = r'D:\mypy\PDFConv2\testpdf\example.xlsx'
    main(ptr, pts)
