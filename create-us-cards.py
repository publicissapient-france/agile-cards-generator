import os
from xlsxwriter.workbook import Workbook
import datetime


def write_us_card(worksheet, starting_row=2, starting_column=0):
    worksheet.merge_range(starting_row, starting_column, starting_row, starting_column + 1, 'MMF:')
    worksheet.merge_range(starting_row, starting_column + 2, starting_row, starting_column + 3, 'Feature:')
    worksheet.merge_range(starting_row, starting_column + 4, starting_row, starting_column + 5, 'Projet:')

    worksheet.merge_range(starting_row + 1, starting_column + 0, starting_row + 1, starting_column + 3, '')
    worksheet.merge_range(starting_row + 1, starting_column + 4, starting_row + 1, starting_column + 5, 'Taille:')

    worksheet.merge_range(starting_row + 2, starting_column, starting_row + 2, starting_column + 5, 'Titre US')

    worksheet.merge_range(starting_row + 3, starting_column, starting_row + 3, starting_column + 1, 'Date backlog')
    worksheet.merge_range(starting_row + 3, starting_column + 2, starting_row + 3, starting_column + 3, 'Date dev')
    worksheet.merge_range(starting_row + 3, starting_column + 4, starting_row + 3, starting_column + 5, 'Date done')


def main():
    file_name = prepare_output_file(None, 'xlsx')

    workbook = Workbook(file_name)
    worksheet = workbook.add_worksheet()

    text = 'Time at which file was generated: '+ datetime.datetime.now().__str__()

    worksheet.write(0,0,text)

    write_us_card(worksheet)

    write_us_card(worksheet, starting_column=7)

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