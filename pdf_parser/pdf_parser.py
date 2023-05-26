# import pdfplumber
from datetime import datetime
from tabula import read_pdf
import pandas as pd
import numpy as np
import argparse
import time
import re
import os


def get_args():
    parser = argparse.ArgumentParser(description='EXCEL to CSV parser')
    parser.add_argument('-src', '--source', metavar='<src-filename>',
                        help='Specifies the absolute path to the source file. Default path - "pdf/pdf_data.pdf"')
    parser.add_argument('-dst', '--destination', metavar='<dst-filename>',
                        help='Specifies the absolute path to the output file.')

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


def split_address(row):
    if re.search(r',+', row):
        split_pattern = r"^([A-Za-z.\s']*\d[A-Za-z.\d\s']*)?,?\s?([A-Z'\sa-z]*?)?,?\s?([A-Z]{2})?\s?([\d-]*)?$"
        match = re.search(split_pattern, row)

        if match:
            address = match.group(1) if match.group(1) else None
            city = match.group(2) if match.group(2) else None
            state = match.group(3) if match.group(3) else None
            zipcode = match.group(4) if match.group(4) else None
            return address, city, state, zipcode

    return row, None, None, None


def normalize_mobile_number(row):
    digits_only = re.sub(r'\D', '', row)

    if len(digits_only) == 11:
        formatted_number = re.sub(r'(1)(\d{3})(\d{3})(\d{4})', r'\1-\2-\3-\4', digits_only)
    else:
        formatted_number = re.sub(r'(\d{3})(\d{3})(\d{4})', r'\1-\2-\3', digits_only)

    return formatted_number


def normalize_date(row):
    try:
        date = datetime.strptime(row, "%d %B %Y").date()
        if date > datetime.now().date():
            date = row
    except ValueError:
        date = row

    return date


def validation_process(row):
    pattern_name = r"^([A-Za-z\s\.',-]*)$"
    pattern_tel = r'^[1]?\d{10}$'
    pattern_email = r'^([a-z0-9_-]+\.)*[a-z0-9_-]+@[a-z0-9_-]+(\.[a-z0-9_-]+)*\.[a-z]{2,6}$'

    name = bool(re.match(pattern_name, row['name'])) if row['name'] else True
    tel = bool(re.match(pattern_tel, re.sub(r'\(|\)|-|\.|\s', '', row['tel']))) if row['tel'] else True
    email = bool(re.match(pattern_email, row['email'])) if row['email'] else True
    dob = bool(row['date'] != normalize_date(row['date'])) if row['date'] else True

    return all([name, tel, email, dob])


def compile_additional_info(row):
    return f"nationality:{str(row['nationality'])}" if row['nationality'] is not None else None


def processing(df, source, destination):
    df['valid'] = df.apply(validation_process, axis=1)
    df['user_additional_info'] = df.apply(compile_additional_info, axis=1)
    df['user_fullname'] = df['name']
    df['name'] = os.path.basename(source)
    df[['address', 'city', 'state', 'zip']] = df['address'].apply(split_address).apply(pd.Series)
    df['tel'] = df['tel'].apply(normalize_mobile_number)
    df['dob'] = df['date'].apply(normalize_date)

    ready_df = pd.DataFrame({'name': df['name'],
                             'usermail': df['email'],
                             'address': df['address'],
                             'user_fullname': df['user_fullname'],
                             'city': df['city'],
                             'state': df['state'],
                             'zip': df['zip'],
                             'tel': df['tel'],
                             'dob': df['dob'],
                             'user_additional_info': df['user_additional_info'],
                             'valid': df['valid']})

    selected_columns = ['name',
                        'usermail',
                        'address',
                        'user_fullname',
                        'city',
                        'state',
                        'zip',
                        'tel',
                        'dob',
                        'user_additional_info']

    df_to_csv(ready_df, destination, selected_columns)


def main():
    source, destination, debug = get_args().source, get_args().destination, get_args().debug
    start = time.time()

    if not source:
        source = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                              'pdf', 'data_pdf.pdf')

    if destination:
        destination = os.path.join(os.path.dirname(destination),
                                   destination.split('.')[0])

    if not destination and source:
        destination = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   os.path.basename(source).split('.')[0] + '_result')

    if destination and source:
        try:
            # old version needs x8 time for parsing from pdf, but new needs java
            # with pdfplumber.open(source) as pdf:
            #     page = pdf.pages[0]
            #     table = page.extract_table()[1:]
            #
            # df = pd.DataFrame(table, columns=['field', 'data'])

            df_tabula = read_pdf(source, encoding="ISO-8859-1", guess=False, stream=True, pages='all')
            data_2d = np.squeeze(df_tabula)
            df = pd.DataFrame(data_2d, columns=['field', 'data'])

            grouped_dates = df.groupby('field')['data']
            transposed_df = pd.DataFrame()

            for field, data in grouped_dates:
                transposed_df[str(field)] = pd.Series(data.values)

            processing(transposed_df, source, destination)

            print(f'CSV is ready. Path: {destination + ".csv"}')

            if debug:
                print(f'Time spent to execution: {time.time() - start} sec.')

        except FileNotFoundError:
            print('Src file not found')


if __name__ == '__main__':
    main()
