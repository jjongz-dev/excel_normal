from Architect.FIN.ItemStandard import ItemStandard

import re


def launch(item: ItemStandard,levels,floorsupportlevels):
    # 산식 층정리
    if (item.floor == '공통가설' or item.floor == '골조가설') and item.formula.__contains__('>') and item.formula.__contains__('<'):
        aaa = item.formula.split('>')[0].split('<')[1]
        aaa2 = aaa.split('F')[0] + 'F'
        aaa3 = aaa.split('층')[0] + '층'
        aaa4 = aaa.split('붕')[0] + '붕'
        aaa5 = aaa.split('초')[0] + '초'
        if re.match('PH\\d{1,2}F', aaa) or re.match('B\\d{1,2}F', aaa) or re.match('\\d{1,2}F', aaa) or re.match('RF', aaa) or re.match('PHRF', aaa):
            item.location = item.floor
            item.roomname = item.floor
            item.floor = aaa2
        if re.match('지상\\d{1,2}층', aaa) or re.match('지하\\d{1,2}층', aaa) or re.match('옥탑\\d{1,2}층', aaa) or re.match('\\d{1}층', aaa):
            item.location = item.floor
            item.roomname = item.floor
            item.floor = aaa3
        if re.match('지붕', aaa) or re.match('옥탑지붕', aaa):
            item.location = item.floor
            item.roomname = item.floor
            item.floor = aaa4
        if re.match('기초', aaa):
            item.location = item.floor
            item.roomname = item.floor
            item.floor = aaa5

    # 층정리
    if item.type == '외부' and (re.match('지상\\d{1,2}층 \\w+', item.floor) or re.match('지하\\d{1,2}층 \\w+', item.floor)):
        split_floor = item.floor.split(' ')
        item.floor = split_floor[0]
        item.location = split_floor[1]
        item.roomname = split_floor[1]

    if re.match('\\d{1,2}[.] 옥탑\\d{1,2}층', item.floor):
        item.floor = 'PH' + re.sub(r'[^0-9]', '', item.floor.split(' ')[1]) + 'F'

    if re.match('\\d{1,2}[.] 지상\\d{1,2}층', item.floor):
        item.floor = re.sub(r'[^0-9]', '', item.floor.split(' ')[1]) + 'F'

    if re.match('\\d{1,2}[.] 지하\\d{1,2}층', item.floor):
        item.floor = 'B' + re.sub(r'[^0-9]', '', item.floor.split(' ')[1]) + 'F'

    if re.match('옥탑\\d{1,2}층', item.floor):
        item.floor = 'PH' + re.sub(r'[^0-9]', '', item.floor) + 'F'

    if re.match('지상\\d{1,2}층', item.floor):
        item.floor = re.sub(r'[^0-9]', '', item.floor) + 'F'

    if re.match('지하\\d{1,2}층', item.floor):
        item.floor = 'B' + re.sub(r'[^0-9]', '', item.floor) + 'F'

    if re.match('\\d{1,2}층', item.floor):
        item.floor = re.sub(r'[^0-9]', '', item.floor) + 'F'

    if item.floor.__contains__('지붕'):
        item.floor = item.floor.replace('지붕', 'RF')

    if item.floor.__contains__('기초'):
        item.floor = item.floor.replace('기초', 'FT')


    # 최대 레벨값 뽑기
    if item.name.__contains__('먹메김') and re.match('\\d{1,2}F', item.floor):
        ccc = int(item.floor.replace('F', ''))
        levels.append(ccc)

    # 동바리 <F2바닥> 으로 산출시 정리하기
    if item.name.__contains__('동바리') and item.formula.__contains__('바닥'):
        if re.match('B\\d{1,2}F', item.floor):
            item.floor = 'B' + str(int(item.floor.replace('F', '').replace('B', '')) + 1) + 'F'

        if re.match('\\d{1,2}F', item.floor):
            bbb = int(item.floor.replace('F', '')) - 1

            if bbb == '0' or bbb == 0:
                item.floor = 'B1F'
            else:
                item.floor = str(bbb) + 'F'
                floorsupportlevels.append(bbb)


