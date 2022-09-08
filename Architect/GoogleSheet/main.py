from openpyxl import load_workbook, Workbook

from Architect.GoogleSheet.ItemStandard import ItemStandard



def excel_normalize(name, column_dimensions=None):
    excel = load_workbook(
        'C:\\Users\ckddn\Desktop\창호.xlsx',
        data_only=True)

    items = []
    worksheet = excel['초안']
    for row in worksheet.iter_rows(min_col=0, max_col=17, min_row=3):
        # 이름변경
        if ( row[0].value is not None
                and row[1].value is not None
                and row[2].value is not None
        ):
            temp_names = f"{row[0].value} | {row[1].value}*{row[2].value} | M2"

        # 유리변경
        if ( row[3].value is None
                and row[4].value is None
                and row[5].value is None
        ):
            temp_standard = ''

        if ( row[3].value is not None
                and row[4].value is not None
                and row[5].value is not None
        ):
            temp_standard =f"{row[5].value}유리 | {row[4].value}mm | M2"


        if ( row[3].value is not None
                and row[4].value is not None
                and row[5].value is not None
                and row[6].value is not None
        ):
            temp_midstandard = f"{row[6].value.split('+')[0]}+{row[6].value.split('+')[1]}A+{row[6].value.split('+')[2]}"
            temp_standard =f"{row[5].value}유리 | {row[4].value}mm,({temp_midstandard}) | M2"


        if ( row[3].value is not None
                and row[4].value is not None
                and row[5].value is not None
                and row[6].value is not None
                and row[7].value is not None
        ):
            temp_midstandard = f"{row[6].value.split('+')[0]}+{row[6].value.split('+')[1]}Ar+{row[6].value.split('+')[2]}"
            temp_standard =f"{row[5].value}유리 | {row[4].value}mm,({temp_midstandard}) | M2"

        # 도어변경
        if ( row[8].value is None
                and row[9].value is None
                and row[10].value is None
                and row[11].value is None
        ):
            temp_glassdoor = ''

        if ( row[8].value is not None
                and row[9].value is not None
                and row[10].value is not None
                and row[11].value is not None
        ):
            if ( row[9].value == '강화'
            ):
                width_glassdoor = int(row[10].value*1000)
                height_glassdoor = int(row[11].value * 1000)
                temp_glassdoor = f"{row[9].value}유리도어 | 투명,{row[8].value}mm,{width_glassdoor}*{height_glassdoor},손보호용 | EA"
            else:
                width_glassdoor = int(row[10].value * 1000)
                height_glassdoor = int(row[11].value * 1000)
                temp_glassdoor = f"{row[9].value}유리도어 | 투명,{row[8].value}mm,{width_glassdoor}*{height_glassdoor} | EA"

        if ( row[16].value is not None
        ):
            temp_remark = row[16].value

        print(temp_glassdoor)
        item = ItemStandard(
            windows_name = temp_names,
            glass_standard = temp_standard,
            glass_door = temp_glassdoor,
            fire_entrance = '',
            system_door = '',
            insect_screen = '',
            houseHold = '',
            remark = temp_remark,
            )
        items.append(item)





    # 저장할 엑셀
    new_workbook = Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = '창호완성'
    head_title = ['창호', '유리', '도어', '소방진입창', '세이프도어', '방충망', '98세대', 'Remark']
    new_sheet.append(head_title)
    new_sheet.column_dimensions["A"].width = 20
    new_sheet.column_dimensions["B"].width = 30
    new_sheet.column_dimensions["C"].width = 50
    new_sheet.column_dimensions["D"].width = 12
    new_sheet.column_dimensions["E"].width = 12
    new_sheet.column_dimensions["F"].width = 12
    new_sheet.column_dimensions["G"].width = 12
    new_sheet.column_dimensions["H"].width = 12

    for item in items:
        new_sheet.append(item.to_excel())


    new_workbook.save("C:\\Users\ckddn\Desktop\창호완성.xlsx")



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    excel_normalize('PyCharm')


