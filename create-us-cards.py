import os
from xlsxwriter.workbook import Workbook
import datetime


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

def main():
    file_name = prepare_output_file(None, 'xlsx')

    workbook = Workbook(file_name)
    worksheet = workbook.add_worksheet()

    text = 'Time at which file was generated: '+ datetime.datetime.now().__str__()

    worksheet.write(0,0,text)

    num_of_cards = 5
    cards = []

    for i in range(num_of_cards):
        new_card = USCard(title='Titre US ' + i.__str__())
        cards.append(new_card)

    row = 0
    cards_per_line = 2
    card_position_on_line = 0

    for card in cards:
        write_us_card(worksheet, card, starting_row=row, starting_column=card_position_on_line * 7)
        card_position_on_line += 1

        if card_position_on_line == cards_per_line:
            card_position_on_line = 0
            row += 5

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