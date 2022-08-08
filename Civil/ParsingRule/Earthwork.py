from Civil.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 터파기
    if (item.name in ['굴착 및 상차']):
        item.name = '터파기'
        item.standard = item.standard.replace(" ","") + '★크람쉘분리추가'

    # 상차
    if (item.name in ['운반 및 사토장 정지']):
        item.name = '잔토처리'
        item.standard = item.standard.replace(" ","")

