from FIN.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 낙화물방지망
    if item.name.__contains__('낙하물방지'):
        item.name = item.name + '★삭제'
        item.formula = '0'
        item.sum = '0'


