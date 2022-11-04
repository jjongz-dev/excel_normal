from Architect.FIN.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 낙화물방지망
    if item.name.__contains__('낙하물방지'):
        item.name = item.name + '★삭제'
        item.formula = '0'
        item.sum = '0'

    # 소방관집입창스티커부착 + 크러쉬버튼설치(소방관집입창 유리파괴장치)
    if item.name.__contains__('소방관진입창스티커부착') or item.name.__contains__('크러쉬버튼설치(소방관진입창 유리파괴장치)'):
        item.name = item.name + '★삭제'
        item.formula = '0'
        item.sum = '0'

    # 지반다짐
    if item.name.__contains__('지반다짐'):
        item.name = item.name + '★삭제'
        item.formula = '0'
        item.sum = '0'