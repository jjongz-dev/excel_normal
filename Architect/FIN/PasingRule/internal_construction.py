from Architect.FIN.ItemStandard import ItemStandard

import re

def launch(item: ItemStandard):
    # 호,실 구분하기
    if '호' in item.roomname:
        temp_roomname = item.roomname.split('호')[0]
        if re.match('\\d', temp_roomname):
            item.location = temp_roomname + '호'
            item.roomname = item.roomname.split('호')[-1]