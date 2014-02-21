import os
from openpyxl.cell import coordinate_from_string, column_index_from_string

import openpyxl
import datetime
import csv


def write_us_card(worksheet, card, starting_row=0, starting_column=0):
    starting_row = starting_row + 2

    worksheet.merge_range(starting_row, starting_column, starting_row, starting_column + 1, card.mmf)
    worksheet.merge_range(starting_row, starting_column + 2, starting_row, starting_column + 3, card.feature)
    worksheet.merge_range(starting_row, starting_column + 4, starting_row, starting_column + 5, card.project)

    worksheet.merge_range(starting_row + 1, starting_column + 0, starting_row + 1, starting_column + 3, '')
    worksheet.merge_range(starting_row + 1, starting_column + 4, starting_row + 1, starting_column + 5, card.size)

    worksheet.merge_range(starting_row + 2, starting_column, starting_row + 2, starting_column + 5, card.title)

    worksheet.merge_range(starting_row + 3, starting_column, starting_row + 3, starting_column + 1, card.date_backlog)
    worksheet.merge_range(starting_row + 3, starting_column + 2, starting_row + 3, starting_column + 3, card.date_dev)
    worksheet.merge_range(starting_row + 3, starting_column + 4, starting_row + 3, starting_column + 5, card.date_done)

class USCard():

    def __init__(self, mmf='MMF:', feature='Feature:', project='Projet:', size='Taille', title='Titre de la US',
                          date_backlog='Date backlog:', date_dev='Date dev:', date_done='Date done'):
        self.mmf = mmf
        self.feature = feature
        self.project = project
        self.size = size
        self.title = title
        self.date_backlog = date_backlog
        self.date_dev = date_dev
        self.date_done = date_done


def load_cards():
    cards = []

    file_name = os.path.join('test-input-file', 'cards.csv')
    with open(file_name, encoding='utf8') as csv_file:
        csv_reader = csv.reader(csv_file)
        for row in csv_reader:
            new_card = USCard(mmf=row[0], feature=row[1], project=row[2], size=row[3], title=row[4],
                              date_backlog=row[5], date_dev=row[6], date_done=row[7])

            cards.append(new_card)
    return cards


def get_mins_maxs_from_range(range_string):
    min_col, min_row = coordinate_from_string(range_string.split(':')[0])
    max_col, max_row = coordinate_from_string(range_string.split(':')[1])
    min_col = column_index_from_string(min_col)
    max_col = column_index_from_string(max_col)
    return (min_col, min_row, max_col, max_row)


def duplicate_cell_with_offset(cell, worksheet=None, row=0, column=0):
    if not worksheet:
        worksheet = cell.parent

    new_cell = cell.offset(row=row, column=column)
    new_cell.value = cell.value
        # used info on https://groups.google.com/forum/#!topic/openpyxl-users/s27khYlovwU
    worksheet._styles[new_cell.address] = cell.style

    for range_string in worksheet._merged_cells:
        if cell.address == range_string.split(':')[0]:

            min_col, min_row, max_col, max_row = get_mins_maxs_from_range(range_string)
            rows_in_range = max_row - min_row + 1
            columns_in_range = max_col - min_col + 1
            worksheet.merge_cells('%s:%s' % (new_cell.address,
                                             new_cell.offset(row=rows_in_range - 1,
                                                             column=columns_in_range - 1).address))

            # For some reason need also to apply style to each of the merged cells
            for r_offset in range(rows_in_range):
                for c_offset in range(columns_in_range):
                    worksheet._styles[new_cell.offset(row=r_offset, column=c_offset).address] = cell.style


    return new_cell


def main():
    output_file_name = prepare_output_file(None, 'xlsx')

    input_file_name = os.path.join('input', 'input.xlsx')

    my_workbook = openpyxl.load_workbook(input_file_name)

    my_worksheet = my_workbook.get_sheet_by_name(my_workbook.get_sheet_names()[0])

    for row in range(4):
        for column in range(6):
            my_cell = my_worksheet.cell(row=row, column=column)
            duplicate_cell_with_offset(my_cell, row=0, column=7)

    my_workbook.save(output_file_name)

    text = 'Time at which file was generated: '+ datetime.datetime.now().__str__()

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