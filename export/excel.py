import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font


def _make_headers(sheet, headers):
    cell_range = headers.get('range')
    sheet.merge_cells(cell_range)
    coordinates = cell_range.split(':')[0]
    c, r = re.sub(r'\d', '', coordinates), re.sub(r'\D', '', coordinates)
    c, r = ord(c) + 1 - 65, int(r)
    cell = sheet.cell(r, c, headers.get('title'))
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.font = Font(bold=True)
    for index, sub in enumerate(headers.get('subs', [])):
        sub_c, sub_r = c + index, r + 1
        sub_cell = sheet.cell(sub_r, sub_c, sub.get('title'))
        sub_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        sub_cell.font = Font(bold=True)


def _get_index(sheet, headers):
    sub_num = len(headers.get('subs'))
    start = 1
    while start < 3000:
        base = chr(64 + (start // 26)) if start // 26 else ''
        x1 = chr(65 + start % 26)
        x2 = chr(65 + start % 26 + sub_num - 1)
        title_range = '{}{}1:{}{}1'.format(base, x1, base, x2)
        title = sheet[title_range][0][0].value
        if title == headers.get('title'):
            # insert data
            return False, start, '{}{}'.format(base, x1)
        if not title:
            headers['range'] = title_range
            _make_headers(sheet, headers)
            return True, start, '{}{}'.format(base, x1)
        start += sub_num
    return False, -1,


def _insert_data(sheet, column, new_column, data):
    if new_column:
        row = 3
        table_date = sheet['A3'].value
        if not table_date:
            for i, item in enumerate(data):
                sheet.cell(row + i, 1, item[0])
                sheet.cell(row + i, column + 1, item[1])
                sheet.cell(row + i, column + 1 + 1, item[2])
            return
        table_date = table_date.split('-')
        new_row = 0
        for index in range(len(data)):
            data_date = data[index][0].split('-')
            for t_v, d_v in zip(table_date, data_date):
                if int(d_v) > int(t_v):
                    new_row += 1
                    break
            else:
                break
        if new_row:
            [sheet.insert_rows(3) for _ in range(new_row)]
            for i, item in enumerate(data):
                sheet.cell(3 + i, 1, item[0])
                sheet.cell(3 + i, column + 1, item[1])
                sheet.cell(3 + i, column + 1 + 1, item[2])
        else:
            data_date = data[0][0].split('-')
            for index in range(3, 100):
                table_date = sheet['A{}'.format(index)].value.split('-')
                if not table_date:
                    break
                for t_v, d_v in zip(table_date, data_date):
                    if int(d_v) < int(t_v):
                        row += 1
                        break
                else:
                    break
            for i, item in enumerate(data):
                sheet.cell(row + i, column + 1, item[1])
                sheet.cell(row + i, column + 1 + 1, item[2])
    else:
        for i, item in enumerate(reversed(data)):
            if sheet['A3'].value == item[0]:
                continue
            table_date = sheet['A3'].value.split('-')
            data_date = item[0].split('-')
            is_new = False
            for t_v, d_v in zip(table_date, data_date):
                if int(d_v) > int(t_v):
                    is_new = True
                elif int(d_v) == int(t_v):
                    continue
                else:
                    break
            if is_new:
                sheet.insert_rows(3)
                sheet.cell(3, 1, item[0])
                sheet.cell(3, column + 1, item[1])
                sheet.cell(3, column + 1 + 1, item[2])


def make(dtype, headers, data):
    file_name = './fond.xlsx'
    if os.path.exists(file_name):
        wb = load_workbook(file_name)
    else:
        wb = Workbook()
    try:
        sheet = wb.get_sheet_by_name(dtype)
    except KeyError as e:
        sheet = wb.create_sheet(dtype)
        _make_headers(sheet, {'title': '名称', 'range': 'A1:A2'})
    new_column, column, column_range = _get_index(sheet, headers)
    _insert_data(sheet, column, new_column, data)
    wb.save(file_name)


if __name__ == '__main__':
    a = [['2020-11-13', '8.0420', '8.9320'], ['2020-11-06', '7.9150', '8.8050'], ['2020-10-30', '7.7691', '8.6591'], ['2020-10-23', '7.5531', '8.4431']]
    b = [['2020-11-27', '7.6666', '5.8888'], ['2020-11-20', '3.6666', '5.8888'], ['2020-11-13', '3.6620', '5.3590'], ['2020-11-06', '3.6570', '5.3540'], ['2020-10-30', '3.5060', '5.2030'], ['2020-10-23', '3.4860', '5.1830']]
    c = [['2020-11-27', '22.6666', '22.8888'], ['2020-11-20', '3.6666', '5.8888'], ['2020-11-13', '3.6620', '5.3590'], ['2020-11-06', '3.6570', '5.3540'],
         ['2020-10-30', '3.5060', '5.2030'], ['2020-10-23', '3.4860', '5.1830']]
    data = [
        [110011, a],
        [519069, b],
        [519068, c],
        [519099, a],
    ]
    for fid, d in data:
        headers = {
            'title': '易方达（{}）'.format(fid),
            'subs': [
                {'title': '单位净值'},
                {'title': '累计净值'},
            ]
        }
        make('公募', headers, d)
