import configparser
import json
import os
import urllib.request
import xlrd

BASE_PATH = os.path.join(os.path.dirname(__file__), '..')


def download_xls(url, filename):
    if url is None:
        raise ValueError('URL value cannot be None, must always be provided')
    urllib.request.urlretrieve(url, filename)


def get_header(worksheet):
    return worksheet.row_values(0)


def get_body_rows(worksheet, limit):
    nrows = min(worksheet.nrows, limit)
    for r in range(1, nrows):
        yield worksheet.row_values(r)


def run_script(section, limit=float('inf')):
    params_file = os.path.join(BASE_PATH, 'params.ini')
    config = configparser.ConfigParser()
    config.read(params_file)

    url = config.get(section, 'mic url')
    xls_file = os.path.join(BASE_PATH, 'download',
                            config.get(section, 'filename'))
    if os.path.isfile(xls_file):
        os.remove(xls_file)
    download_xls(url, xls_file)

    workbook = xlrd.open_workbook(xls_file)
    worksheet = workbook.sheet_by_name(config.get(section, 'sheet name'))

    output = []

    header = get_header(worksheet)
    for row in get_body_rows(worksheet, limit):
        output.append({k: v for k, v in zip(header, row)})

    json_file = os.path.join(
        BASE_PATH, 'output', config.get(section, 'output file'))
    if os.path.isfile(json_file):
        os.remove(json_file)
    with open(json_file, 'w') as fp:
        json.dump(output, fp)


if __name__ == '__main__':
    run_script('ISO20022')
