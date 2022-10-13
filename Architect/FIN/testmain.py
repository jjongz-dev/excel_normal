from openpyxl import load_workbook, Workbook

from Architect.FIN.ItemStandard import ItemStandard



def excel_normalize(name, column_dimensions=None):
    excel = load_workbook(
        'C:\\Users\ckddn\Desktop\건축.xlsx',
        data_only=True)

    items = []

    start_titles = ['부위', '도형', '구분']
    title_row_index = ""
    title_column_index = ""
    part_index = ""
    name_index = ""
    standard_index = ""
    unit_index = ""
    formular_index = ""
    quantity_index = ""
    figure_index = ""

    if '가설산출서' in excel.sheetnames:
        worksheet = excel['가설산출서']

        row_index = 1
        for row in worksheet.iter_rows():
            for column_index, column in enumerate(row):
                if column.value in start_titles:
                    title_row_index = row_index
                    title_column_index = column_index
                    break
            row_index += 1

        # 열의 타이틀을 가지고 그 열의 번호를 따서 row[]에 들어감
        for index, column in enumerate(worksheet[title_row_index]):
            match column.value:
                case '부위':
                    part_index = index
                case '품명':
                    name_index = index
                case '규격':
                    standard_index = index
                case '단위':
                    unit_index = index
                case '산식':
                    formular_index = index
                case '물량':
                    quantity_index = index

        for row in worksheet.iter_rows(min_col=0, max_col=worksheet._current_row, min_row=(title_row_index+1)):
            if (row[part_index].value is not None
                    and row[name_index].value is None
                    and row[standard_index].value is None
                    and row[unit_index].value is None
                    and row[formular_index].value is None
                    and row[quantity_index].value is None
            ):
                try:
                    temp_split_roomname = row[part_index].value.split('개소')[0]
                    if '구분명' in temp_split_roomname:
                        temp_roomname = temp_split_roomname.split(':')[-1].strip('[] ')
                    continue
                except Exception as e:
                    print('예외가 발생했습니다.', e)
                    print('가설산출서 : split오류' + str(item))

            # 품명없음 삭제
            if (row[quantity_index].value is None
                    or row[quantity_index].value == "'"
                    or row[quantity_index].value == '0'
                    or row[quantity_index].value == 0
            ):
                continue

            item = ItemStandard(
                floor='1F',
                location=temp_roomname,
                roomname=temp_roomname,
                name=row[name_index].value,
                standard=row[standard_index].value,
                unit=row[unit_index].value,
                type='내부',
                formula=row[formular_index].value,
                sum=row[quantity_index].value,
            )
            items.append(item)

    if '토공산출서' in excel.sheetnames:
        worksheet = excel['토공산출서']

        row_index = 1
        for row in worksheet.iter_rows():
            for column_index, column in enumerate(row):
                if column.value in start_titles:
                    title_row_index = row_index
                    title_column_index = column_index
                    break
            row_index += 1

        for index, column in enumerate(worksheet[title_row_index]):
            match column.value:
                case '도형':
                    figure_index = index
                case '품명':
                    name_index = index
                case '규격':
                    standard_index = index
                case '단위':
                    unit_index = index
                case '산식':
                    formular_index = index
                case '물량':
                    quantity_index = index

            print(figure_index)

        for row in worksheet.iter_rows(min_col=0, max_col=worksheet._current_row, min_row=(title_row_index+1)):
            if (row[figure_index].value is not None
                    and row[name_index].value is None
                    and row[standard_index].value is None
                    and row[unit_index].value is None
                    and row[formular_index].value is None
                    and row[quantity_index].value is None
            ):
                try:
                    temp_split_roomname = row[figure_index].value.split('개소')[0]
                    if '구분명' in temp_split_roomname:
                        temp_roomname = temp_split_roomname.split(':')[-1].strip('[] ')
                    continue
                except Exception as e:
                    print('예외가 발생했습니다.', e)
                    print('토공산출서 : split오류' + str(item))

            # 품명없음 삭제
            if (row[quantity_index].value is None
                    or row[quantity_index].value == "'"
                    or row[quantity_index].value == '0'
                    or row[quantity_index].value == 0
            ):
                continue

            item = ItemStandard(
                floor='1F',
                location=temp_roomname,
                roomname=temp_roomname,
                name=row[name_index].value,
                standard=row[standard_index].value,
                unit=row[unit_index].value,
                type='외부',
                formula=row[formular_index].value,
                sum=row[quantity_index].value,
            )
            items.append(item)


    for item in items:
        print(item)

# 저장할 엑셀
    new_workbook = Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = '건축(데이터변경X)'
    head_title = ['층', '호', '실', '대공종', '중공종', '코드', '품명', '규격', '단위', '부위', '타입', '산식', '수량', 'Remark']
    new_sheet.append(head_title)
    new_sheet.column_dimensions["G"].width = 30
    new_sheet.column_dimensions["H"].width = 30
    new_sheet.column_dimensions["L"].width = 30
    for item in items:
        new_sheet.append(item.to_excel())

    new_workbook.save("C:\\Users\ckddn\Desktop\건축완성.xlsx")



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    excel_normalize('PyCharm')


