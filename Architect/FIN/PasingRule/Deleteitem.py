from Architect.FIN.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 낙화물방지망
    if item.name.__contains__('낙하물방지'):
        item.name = item.name + '★삭제'
        item.formula = '0'
        item.sum = '0'

    if item.name.__contains__('소방관'):
        item.name = item.name + '★삭제'
        item.formula = '0'
        item.sum = '0'

    if item.name.__contains__('완강기'):
        item.name = item.name + '★삭제+기계전기전달'
        item.formula = '0'
        item.sum = '0'

    if item.name.__contains__('지반다짐'):
        item.name = item.name + '★삭제'
        item.formula = '0'
        item.sum = '0'

    if item.name.__contains__('타일면보양'):
        item.name = item.name + '★삭제'
        item.formula = '0'
        item.sum = '0'

    if item.name.__contains__('석재면보양'):
        item.name = item.name + '★삭제'
        item.formula = '0'
        item.sum = '0'

