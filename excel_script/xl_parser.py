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
                        help='Specifies the absolute path to the output file')

    debug = parser.add_mutually_exclusive_group()

    debug.add_argument('-d', '--debug',
                       action='store_true',
                       help='use option to get feedback about completing')

    return parser.parse_args()


def df_to_csv(df, destination, selected_columns):
    clear_df = df[df['valid'].astype(bool) == True]
    clear_df[selected_columns].to_csv(f"{destination}.csv", index=False, header=True)

    bad_df = df[df['valid'].astype(bool) == False]
    bad_df[selected_columns].to_csv(f"{destination}_bad.csv", index=False, header=True)


def compile_additional_info(row, sep):
    ssn = str(row['SSN'])
    company = str(row['Company'])
    department = str(row['Department'])
    position = str(row['Position'])

    return sep.join([f'ssn:{ssn}', f'company:{company}', f'department:{department}', f'position:{position}'])


def normalize_mobile_number(row):
    digits_only = re.sub(r'\D', '', row)
    formatted_number = re.sub(r'(\d{3})(\d{3})(\d{4})', r'\1-\2-\3', digits_only)

    return formatted_number


def validation_process(row):
    pattern_ssn = r'^(?:\d[-.]?){2}\d[-.]?(?:\d[-.]?){4}\d[-.]?\d$'
    pattern_name = r"^(?!.*[A-Z]{3})(?!.*[A-Z].*[A-Z].*[A-Z])(?:[A-Z][a-z']* ?)+$"
    pattern_mobile = r'^[1]?\d{10}$'

    ssn = bool(re.match(pattern_ssn, row['SSN']))
    first_name = bool(re.match(pattern_name, row['First Name']))
    last_name = bool(re.match(pattern_name, row['Last Name']))
    mobile = bool(re.match(pattern_mobile, re.sub(r'\(|\)|-|\.|\s', '', row['Mobile number'])))

    return all([ssn, first_name, last_name, mobile])


def split_address(row):
    try:
        address, city, state = [_.strip() for _ in row.split(',')]
        state_pattern = r'([A-Za-z]+)'
        match = re.search(state_pattern, state.strip())
        state = match.group(1) if match else None

        return address, city, state

    except ValueError:
        state_pattern = r',\s*([A-Za-z]+)(?:\s|\-)\d+'
        match = re.search(state_pattern, row)

        if match:
            return row, None, match.group(1)
        else:
            return row, None, None


def processing(df, destination, sep):
    df['valid'] = df.apply(validation_process, axis=1)
    df['user_additional_info'] = df.apply(compile_additional_info, args=(sep,), axis=1)
    df['user_fullname'] = df.apply(lambda row: f"{str(row['First Name'])} {str(row['Last Name'])}", axis=1)
    df[['address', 'city', 'state']] = df['Address'].apply(split_address).apply(pd.Series)
    df['Mobile number'] = df['Mobile number'].apply(normalize_mobile_number)

    ready_df = pd.DataFrame({'name': df['First Name'],
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
            sep = '|'
            sheet_name = pd.ExcelFile(source).sheet_names[0]
            df = pd.read_excel(io=source, sheet_name=sheet_name)
            processing(df, destination, sep)

            print(f'CSV is ready. Path: {destination + ".csv"}')

            if debug:
                print(f'Time spent to execution: {time.time() - start} sec.')

        except FileNotFoundError:
            print('Src file not found')


if __name__ == '__main__':
    main()
