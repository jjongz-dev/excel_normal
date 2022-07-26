

from openpyxl import load_workbook, Workbook

from FIN.ItemStandard import ItemStandard

from FIN.ItemStandard2 import ItemStandard2


def excel_normalize(name):
    excel = load_workbook(
        'C:\\Users\ckddn\Desktop\건축.xlsx',
        data_only=True)
    # names = excel.get_sheet_names()
    # print(names)
    worksheet = excel['산출근거집계표']
    items = []
    for row in worksheet.iter_rows(min_col=1, max_col=13, min_row=5):
        # #0값 삭제
        # if ( row[9].value == '0'
        #     and row[12].value == '0'):
        #     continue
        # 산출식 없음 삭제
        if ( row[2].value == '합        계'):
            continue
        # 구조이기 삭제
        if ( row[7].value == '구조이기'):
            continue

        # for cell in row:
        #     print(cell.value, end=', ')

        # print(','.join([
        #         row[0].value,
        #         row[1].value,
        #         row[2].value,
        #         row[3].value,
        #         row[4].value,
        #         row[5].value,
        #         row[6].value,
        #         row[7].value,
        #         row[8].value,
        #         row[9].value,
        #         row[10].value,
        #         row[11].value,]
        #     ),
        #     end=''
        # )
        # for cell in row:
        #     print(cell.col_idx, cell.value)
        item = ItemStandard(
            floor = row[7].value,
            location = '',
            roomname=row[8].value,



            name = row[2].value,
            standard= row[3].value,
            unit=row[4].value,
            part = '',
            type=row[5].value,
            formula = row[9].value,
            sum=row[12].value,
            )
        # print(item.to_excel())
        items.append(item)


    worksheet2 = excel['동별집계표']
    items2 = []
    constructionWork = ""
    # name = ""
    for row in worksheet2.iter_rows(min_col=1, max_col=5, min_row=4):
        #중공종
        if (row[1].value.__contains__('내  역  삭  제')
        ):
            continue

        if (row[1].value is not None
               and row[4].value is None
        ):
            constructionWork = row[1].value.replace(" ","")


        item2 = ItemStandard2(
            constructionWork = constructionWork,
            name = row[1].value,
            standard = row[2].value,
            unit = row[3].value,
            sum = row[4].value,
            )
        items2.append(item2)

    for item in items:
        #TYPE변경
        type1 = item.type.split('-')
        if (len(type1) == 2):
            item.type = item.type.split('-')[0]

        if (item.type == '토공'):
            item.type = '외부'

        if (item.type == '가설'):
            item.type = '내부'

        #품명변경
        if (item.name.startswith("★")):
            item.name = item.name.replace("★","")

        #창호
        if (item.type == '창호'):
            item.part = item.floor;
            item.floor = '';
            item.roomname =''

        #외부 입면 호실정리      일단 하자
        if (item.type == '외부'
                and item.floor.__contains__('정면')):
            item.location = '정면';
            item.roomname = '정면';
            item.floor = ''

        if (item.type == '외부'
            and item.floor.__contains__('배면')):
            item.location = '배면';
            item.roomname = '배면';
            item.floor = ''

        if (item.type == '외부'
            and item.floor.__contains__('좌측면')):
            item.location = '좌측면';
            item.roomname = '좌측면';
            item.floor = ''

        if (item.type == '외부'
            and item.floor.__contains__('우측면')):
            item.location = '우측면'
            item.roomname = '우측면'
            item.floor = '';

        #토공사정리
        if (item.floor == '토공사'):
            item.location = '기초하부';
            item.floor = 'FT'

        #공통가설 정리 ★★★★★★★★★★★★★
        aaa = item.formula.split('>')
        if (item.floor == '공통가설'
                and (len(aaa) == 2)):
            item.floor = aaa[0].replace('<', '');
            item.location = '공통가설'

        if (item.floor == '공통가설'
            and (len(aaa) == 3)):
            item.floor = aaa[0].replace('<','');
            item.location = '공통가설'

        if (item.floor == '공통가설'):
            item.location = '공통가설';
            item.floor = '1F'

        # 민원처리
        if (item.floor == '민원처리'):
            item.location = '민원처리';
            item.floor = '1F'

        # 철거
        if (item.floor == '철거'):
            item.location = '철거';
            item.floor = '1F'


        #골조가설정리 ★★★★★★★★★★★★★
        aaa = item.formula.split('>')
        if (item.floor == '골조가설'
                and (len(aaa) == 2)):
            item.floor = aaa[0].replace('<', '');
            item.location = '골조가설'

        if (item.floor == '골조가설'
            and (len(aaa) == 3)):
            item.floor = aaa[0].replace('<','');
            item.location = '골조가설'

        if (item.floor == '골조가설'):
            item.location = '골조가설';
            item.floor = '1F'






    # 저장할 엑셀
    new_workbook = Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = '건축(데이터변경X)'
    head_title = ['층', '호', '실', '대공종', '중공종', '코드', '품명', '규격', '단위', '부위', '타입', '산식', '수량', 'Remark']
    new_sheet.append(head_title)
    for item in items:
        new_sheet.append(item.to_excel())

    sheet = new_workbook.create_sheet(title='집계표')
    sheet.append(['중공종', '품명', '규격', '단위', '수량(할증전)'])
    for item2 in items2:
        sheet.append(item2.to_excel())


    new_workbook.save("C:\\Users\ckddn\Desktop\건축완성.xlsx")

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    excel_normalize('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
