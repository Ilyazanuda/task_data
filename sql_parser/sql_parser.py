from io import StringIO
import argparse
import csv
import time
import re
import os


def get_args():
    parser = argparse.ArgumentParser(description='EXCEL to CSV parser')
    parser.add_argument('-src', '--source', metavar='<src-filename>',
                        help='Specifies the absolute path to the source file. Default path - "excel/data.xlsx"')
    parser.add_argument('-dst', '--destination', metavar='<dst-filename>',
                        help='Specifies the absolute path to the output file')

    debug = parser.add_mutually_exclusive_group()

    debug.add_argument('-d', '--debug',
                       action='store_true',
                       help='use option to get feedback about completing')

    return parser.parse_args()


def write_row(row, valid, destination, mode='a'):
    if valid:
        with open(destination + '.csv', mode, newline='', encoding='ANSI') as csvfile:
            csv_writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_writer.writerow(row)
    else:
        with open(destination + '_bad.csv', mode, newline='', encoding='ANSI') as csvfile:
            csv_writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_writer.writerow(row)


def get_validated_row(row):
    print(row)
    return 0, 1


def processing(source, destination):
    selected_columns = {0: 'user_id',
                        1: 'name',
                        2: 'username',
                        3: 'password',
                        4: 'usermail',
                        5: 'permission',
                        6: 'sex',
                        7: 'country',
                        8: 'birth'}
    sep = '|'

    result_header = ['name', 'username', 'user_ID', 'usermail', 'country', 'user_additional_info']
    write_row(result_header, False, destination, 'w')
    write_row(result_header, True, destination, 'w')

    with open(source, 'r', encoding='ANSI') as file:
        for row in file:
            if re.search(r'^INSERT.*VALUES', row):
                continue

            line = re.sub(r'^\(|\),$|\);$|\t', '', row).strip()

            csv_row = list(*csv.reader(StringIO(line), delimiter=',', quotechar="'", skipinitialspace=True))

            csv_row = [[j, None][j in ('', '0')] for j in [re.sub(r"^'|'$|\t", '', i.strip()) for i in csv_row]]

            # каличный отлов приколов с quotechar
            if len(csv_row) != 9:
                temp_row = list()

                for num, cell in enumerate(csv_row):
                    if cell is not None and ',' in cell:
                        for i in cell.split(','):
                            temp_row.append(i)
                        continue
                    temp_row.append(cell)

                csv_row = temp_row

            raw_row = {selected_columns[num]: cell for num, cell in enumerate(csv_row) if num in selected_columns}
            raw_row, valid = get_validated_row(raw_row)
            # clear_row = get_clear_row(raw_row, sep)
            # write_row(clear_row, valid, destination)


def main():
    source, destination, debug = get_args().source, get_args().destination, get_args().debug
    start = time.time()

    if not source:
        source = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                              'sql', 'data.sql')

    if destination:
        destination = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   destination.split('.')[0])

    if not destination and source:
        destination = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   os.path.basename(source).split('.')[0] + '_result')

    if destination and source:
        try:
            processing(source, destination)

            print(f'CSV is ready. Path: {destination + ".csv"}')

            if debug:
                print(f'Time spent to execution: {time.time() - start} sec.')

        except FileNotFoundError:
            print('Src file not found')


if __name__ == '__main__':
    main()
