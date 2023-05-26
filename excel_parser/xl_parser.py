import pandas as pd
import argparse
import time
import re
import os


def get_args():
    parser = argparse.ArgumentParser(description='EXCEL to CSV parser')
    parser.add_argument('-src', '--source', metavar='<src-filename>',
                        help='Specifies the absolute path to the source file. Default path - "excel/data.xlsx"')
    parser.add_argument('-dst', '--destination', metavar='<dst-filename>',
                        help='Specifies the absolute path to the output file.')

    debug = parser.add_mutually_exclusive_group()

    debug.add_argument('-d', '--debug',
                       action='store_true',
                       help='use option to get feedback about completing')

    return parser.parse_args()


def df_to_csv(df, destination, selected_columns):
    clear_df = df[df['valid'].astype(bool) == True]
    bad_df = df[df['valid'].astype(bool) == False]
    clear_df[selected_columns].to_csv(f"{destination}.csv", index=False, header=True)
    bad_df[selected_columns].to_csv(f"{destination}_bad.csv", index=False, header=True)


def compile_additional_info(row):
    sep = '|'
    additional_info = list()
    additional_info.append(f"ssn:{str(row['SSN'])}" if str(row['SSN']) else None)
    additional_info.append(f"company:{str(row['Company'])}" if str(row['Company']) else None)
    additional_info.append(f"department:{str(row['Department'])}" if str(row['Department']) else None)
    additional_info.append(f"position:{str(row['Position'])}" if str(row['Position']) else None)
    additional_info = [info for info in additional_info if info is not None]

    return sep.join(additional_info)


def normalize_mobile_number(row):
    digits_only = re.sub(r'\D', '', row)

    if len(digits_only) == 11:
        formatted_number = re.sub(r'(1)(\d{3})(\d{3})(\d{4})', r'\1-\2-\3-\4', digits_only)
    else:
        formatted_number = re.sub(r'(\d{3})(\d{3})(\d{4})', r'\1-\2-\3', digits_only)

    return formatted_number


def validation_process(row):
    pattern_ssn = r'^(?:\d[-.]?){2}\d[-.]?(?:\d[-.]?){4}\d[-.]?\d$'
    # pattern_name = r"^(?!.*[A-Z]{3})(?!.*[A-Z].*[A-Z].*[A-Z])(?:[A-Z][a-z']* ?)+$"
    pattern_name = r"^([A-Za-z\s\.',-]*)$"
    pattern_mobile = r'^[1]?\d{10}$'

    ssn = bool(re.match(pattern_ssn, row['SSN'])) if row['SSN'] else True
    first_name = bool(re.match(pattern_name, row['First Name'])) if row['First Name'] else True
    last_name = bool(re.match(pattern_name, row['Last Name'])) if row['Last Name'] else True
    mobile = bool(re.match(pattern_mobile, re.sub(r'\(|\)|-|\.|\s', '', row['Mobile number']))) \
        if row['Mobile number'] else True

    return all([ssn, first_name, last_name, mobile])


def split_address(row):
    # variant 2
    if re.search(r',+', row):
        split_pattern = r"^([A-Za-z.\s']*\d[A-Za-z.\d\s']*)?,?\s?([A-Z'\sa-z]*?)?,?\s?([A-Z]{2})?\s?([\d-]*)?$"
        match = re.search(split_pattern, row)

        if match:
            address = match.group(1) if match.group(1) else None
            city = match.group(2) if match.group(2) else None
            state = match.group(3) if match.group(3) else None
            # zipcode = match.group(4) if match.group(4) else None
            return address, city, state

    return row, None, None

    # variant 1
    # try:
    #     state_pattern = r'([A-Za-z]+)'
    #     address, city, state = [_.strip() for _ in row.split(',')]
    #     match = re.search(state_pattern, state)
    #     state = match.group(1) if match else None
    #
    #     return address, city, state
    #
    # except ValueError:
    #     state_pattern = r',\s*([A-Za-z]+)(?:\s|\-)\d+'
    #     match = re.search(state_pattern, row)
    #
    #     if match:
    #         return row, None, match.group(1)
    #     else:
    #         return row, None, None


def processing(df, source, destination):
    df['valid'] = df.apply(validation_process, axis=1)
    df['user_additional_info'] = df.apply(compile_additional_info, axis=1)
    df['user_fullname'] = df.apply(lambda row: f"{str(row['First Name'])} {str(row['Last Name'])}", axis=1)
    df['name'] = os.path.basename(source)
    df[['address', 'city', 'state']] = df['Address'].apply(split_address).apply(pd.Series)
    df['Mobile number'] = df['Mobile number'].apply(normalize_mobile_number)

    ready_df = pd.DataFrame({'name': df['name'],
                             'address': df['address'],
                             'user_fullname': df['user_fullname'],
                             'city': df['city'],
                             'state': df['state'],
                             'zip': df['Zip'],
                             'tel': df['Mobile number'],
                             'user_additional_info': df['user_additional_info'],
                             'valid': df['valid']})

    selected_columns = ['name',
                        'address',
                        'user_fullname',
                        'city',
                        'state',
                        'zip',
                        'tel',
                        'user_additional_info']

    df_to_csv(ready_df, destination, selected_columns)


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
            sheet_name = pd.ExcelFile(source).sheet_names[0]
            df = pd.read_excel(io=source, sheet_name=sheet_name)
            processing(df, source, destination)

            print(f'CSV is ready. Path: {destination + ".csv"}')

            if debug:
                print(f'Time spent to execution: {time.time() - start} sec.')

        except FileNotFoundError:
            print('Src file not found')


if __name__ == '__main__':
    main()
