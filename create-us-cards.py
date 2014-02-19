import os
from xlsxwriter.workbook import Workbook
import datetime

def main():
    file_name = prepare_output_file(None, 'xlsx')

    workbook = Workbook(file_name)
    worksheet = workbook.add_worksheet()

    text = 'Time at which file was generated: '+ datetime.datetime.now().__str__()

    worksheet.write(0,0,text)

    worksheet.merge_range(2,0,2,1,'MMF:')
    worksheet.merge_range(2,2,2,3,'Feature:')
    worksheet.merge_range(2,4,2,5,'Projet:')

    worksheet.merge_range(3,0,3,3,'')
    worksheet.merge_range(3,4,3,5,'Taille:')

    worksheet.merge_range(4,0,4,5,'Titre US')

    worksheet.merge_range(5,0,5,1,'Date backlog')
    worksheet.merge_range(5,2,5,3,'Date dev')
    worksheet.merge_range(5,4,5,5,'Date done')

    workbook.close()

    print(text)


def prepare_output_file(output_file, extension):
    file_name = None
    if output_file is not None:
        file_name = output_file
    else:
        file_name = 'output.' + extension

    output_path = 'output'
    if not os.path.isdir(output_path):
        os.makedirs(output_path, exist_ok=True)
    file_name = os.path.join(output_path, file_name)

    if os.path.isfile(file_name):
        os.remove(file_name)

    return file_name


if __name__ == "__main__":
    main()