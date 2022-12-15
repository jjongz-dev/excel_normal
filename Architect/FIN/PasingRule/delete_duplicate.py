from Architect.FIN.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 구조물량이기 삭제
    if '구조이기' in item.roomname:
        item.name = item.name + '★삭제(구조이기물량)'
        item.formula = '0'
        item.sum = '0'

    if '타일면보양' in item.name:
        item.name = item.name + '★부자재삭제'
        item.formula = '0'
        item.sum = '0'

    if '석재면보양' in item.name:
        item.name = item.name + '★부자재삭제'
        item.formula = '0'
        item.sum = '0'

    if '유리끼우기및닦기' in item.name:
        item.name = item.name + '★부자재삭제'
        item.formula = '0'
        item.sum = '0'

    if '크러쉬버튼설치(소방관진입창 유리파괴장치)' in item.name:
        item.name = item.name + '★HB기준삭제'
        item.formula = '0'
        item.sum = '0'

    if '★크러쉬버튼설치(소방관진입창 유리파괴장치)' in item.name:
        item.name = item.name + '★HB기준삭제'
        item.formula = '0'
        item.sum = '0'