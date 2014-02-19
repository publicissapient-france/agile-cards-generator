import os
from xlsxwriter.workbook import Workbook
import datetime


def write_us_card(worksheet, starting_row=2, starting_column=2):
    worksheet.merge_range(starting_row, 0, starting_row, 1, 'MMF:')
    worksheet.merge_range(starting_row, 2, starting_row, 3, 'Feature:')
    worksheet.merge_range(starting_row, 4, starting_row, 5, 'Projet:')

    worksheet.merge_range(starting_row + 1, 0, starting_row + 1, 3, '')
    worksheet.merge_range(starting_row + 1, 4, starting_row + 1, 5, 'Taille:')

    worksheet.merge_range(starting_row + 2, 0, starting_row + 2, 5, 'Titre US')

    worksheet.merge_range(starting_row + 3, 0, starting_row + 3, 1, 'Date backlog')
    worksheet.merge_range(starting_row + 3, 2, starting_row + 3, 3, 'Date dev')
    worksheet.merge_range(starting_row + 3, 4, starting_row + 3, 5, 'Date done')


def main():
    file_name = prepare_output_file(None, 'xlsx')

    workbook = Workbook(file_name)
    worksheet = workbook.add_worksheet()

    text = 'Time at which file was generated: '+ datetime.datetime.now().__str__()

    worksheet.write(0,0,text)

    write_us_card(worksheet)

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