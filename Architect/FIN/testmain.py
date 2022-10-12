from openpyxl import load_workbook, Workbook

from Architect.FIN.ItemStandard import ItemStandard


def excel_normalize(name, column_dimensions=None):
    excel = load_workbook(
        'C:\\Users\ckddn\Desktop\건축.xlsx',
        data_only=True)

    part_index = ""
    name_index = ""
    standard_index = ""
    unit_index = ""
    formular_index = ""
    quantity_index = ""

    items = []
    if '가설산출서' in excel.sheetnames:
        worksheet = excel['가설산출서']
        # row = worksheet.rows(4)
        # print(row)
        for index, column in enumerate(worksheet[4]):
            if column.value == '부위':
                part_index = index
            elif column.value == '품명':
                name_index = index
            elif column.value == '규격':
                standard_index = index
            elif column.value == '단위':
                unit_index = index
            elif column.value == '산식':
                formular_index = index
            elif column.value == '물량':
                quantity_index = index

        for row in worksheet.iter_rows(min_col=0, max_col=8, min_row=5):
            if (row[0].value is not None
                    and row[1].value is None
                    and row[2].value is None
                    and row[3].value is None
                    and row[4].value is None
                    and row[5].value is None
                    and row[6].value is None
                    and row[7].value is None
            ):
                temp_roomname = row[part_index].value.split('개소')[0]
                if '구분명' in temp_roomname:
                    temproomname = temp_roomname.split(':')[-1].replace("[", "").replace("]", "").replace(" ", "")
                continue

            # 품명없음 삭제
            if (row[quantity_index].value is None
                    or row[quantity_index].value == "'"
                    or row[quantity_index].value == '0'
                    or row[quantity_index].value == 0
                    or row[quantity_index].value == '물량'
            ):
                continue
            item = ItemStandard(
                floor='1F',
                location=temproomname,
                roomname=temproomname,
                name=row[name_index].value,
                standard=row[standard_index].value,
                unit=row[unit_index].value,
                type='내부',
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



