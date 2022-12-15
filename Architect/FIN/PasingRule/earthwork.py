from Architect.FIN.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 구조물량이기 삭제
    if '우편물반송함' in item.name:
        item.name = '우편물수취함'
        item.formula = '<우편물반송함>' + item.formula

    if '폐건전지수거함' in item.name:
        item.name = '우편물수취함'
        item.formula = '<폐건전지수거함>' + item.formula
