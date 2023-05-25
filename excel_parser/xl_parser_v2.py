import argparse
import openpyxl
import time
import csv
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
        with open(destination + '.csv', mode, newline='') as csvfile:
            csv_writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_writer.writerow(row)
    else:
        with open(destination + '_bad.csv', mode, newline='') as csvfile:
            csv_writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_writer.writerow(row)


def get_clear_row(row, sep):

    def get_normalized_address(address_row):
        # variant 2
        if re.search(r',+', address_row):
            split_pattern = r"^([A-Za-z.\s']*\d[A-Za-z.\d\s']*)?,?\s?([A-Z'\sa-z]*?)?,?\s?([A-Z]{2})?\s?([\d-]*)?$"
            match = re.search(split_pattern, address_row)

            if match:
                address = match.group(1) if match.group(1) else None
                city = match.group(2) if match.group(2) else None
                state = match.group(3) if match.group(3) else None
                #zipcode = match.group(4) if match.group(4) else None
                return address, city, state

        return address_row, None, None

        # variant 1
        # try:
        #     state_pattern = r'([A-Za-z]+)'
        #
        #     address, city, state = [_.strip() for _ in address_row.split(',')]
        #
        #     match = re.search(state_pattern, state)
        #     state = match.group(1) if match else None
        #
        #     return address, city, state
        #
        # except ValueError:
        #     state_pattern = r',\s*([A-Za-z]+)(?:\s|\-)\d+'
        #     match = re.search(state_pattern, address_row)
        #
        #     if match:
        #         return address_row, None, match.group(1)
        #     else:
        #         return address_row, None, None

    def get_normalized_mobile_number(number):
        digits_only = re.sub(r'\D', '', number)

        if len(digits_only) == 11:
            formatted_number = re.sub(r'(1)(\d{3})(\d{3})(\d{4})', r'\1-\2-\3-\4', digits_only)
        else:
            formatted_number = re.sub(r'(\d{3})(\d{3})(\d{4})', r'\1-\2-\3', digits_only)

        return formatted_number

    clear_row = ['' for i in range(8)]
    clear_row[0] = row['first_name']
    clear_row[1], clear_row[3], clear_row[4] = get_normalized_address(row['address'])
    clear_row[2] = ' '.join([row['first_name'], row['last_name']])
    clear_row[5] = row['zip']
    clear_row[6] = get_normalized_mobile_number(row['mobile_number'])
    clear_row[7] = sep.join([f'ssn:{row["ssn"]}',
                             f'company:{row["company"]}',
                             f'department:{row["department"]}',
                             f'position:{row["position"]}'])
    return clear_row


def get_validated_row(row):
    pattern_ssn = r'^(?:\d[-.]?){2}\d[-.]?(?:\d[-.]?){4}\d[-.]?\d$'
    pattern_name = r"^(?!.*[A-Z]{3})(?!.*[A-Z].*[A-Z].*[A-Z])(?:[A-Z][a-z']* ?)+$"
    pattern_mobile = r'^[1]?\d{10}$'

    ssn = bool(re.match(pattern_ssn, row['ssn']))
    first_name = bool(re.match(pattern_name, row['first_name']))
    last_name = bool(re.match(pattern_name, row['last_name']))
    mobile = bool(re.match(pattern_mobile, re.sub(r'\(|\)|-|\.|\s', '', row['mobile_number'])))

    return row, all([ssn, first_name, last_name, mobile])


def process_sheet(sheet, destination):
    selected_columns = {0: 'first_name',
                        1: 'last_name',
                        2: 'ssn',
                        3: 'address',
                        4: 'company',
                        5: 'department',
                        6: 'position',
                        7: 'zip',
                        8: 'mobile_number'}
    sep = '|'
    result_header = ['name', 'address', 'user_fullname', 'city', 'state', 'zip', 'tel', 'user_additional_info']
    write_row(result_header, False, destination, 'w')
    write_row(result_header, True, destination, 'w')

    for row in sheet.iter_rows(min_row=2):
        # dict with elements of row
        raw_row = {selected_columns[num]: cell.value for num, cell in enumerate(row) if num in selected_columns}
        raw_row, valid = get_validated_row(raw_row)
        clear_row = get_clear_row(raw_row, sep)
        write_row(clear_row, valid, destination)


def main():
    source, destination, debug = get_args().source, get_args().destination, get_args().debug
    start = time.time()

    if not source:
        source = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                              'excel', 'data.xlsx')

    if destination:
        destination = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   destination.split('.')[0])

    if not destination and source:
        destination = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   os.path.basename(source).split('.')[0] + '_result')

    if destination and source:
        try:
            sheet = openpyxl.load_workbook(source).active
            process_sheet(sheet, destination)

            print(f'CSV is ready. Path: {destination + ".csv"}')

            if debug:
                print(f'Time spent to execution: {time.time() - start} sec.')

        except FileNotFoundError:
            print('Src file not found')


if __name__ == '__main__':
    main()
