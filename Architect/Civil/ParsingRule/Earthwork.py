from Architect.Civil.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 터파기
    if item.name in ['굴착 및 상차']:
        item.name = '터파기'
        item.standard = item.standard.replace(' ','')


    # 상차
    if item.name in ['운반 및 사토장 정지']:
        item.name = '상차'
        item.standard = '백호'