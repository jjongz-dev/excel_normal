from Civil.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 천공(S.G.R)
    if item.name in ['천   공']:
        item.name = '천공(S.G.R)'
        item.standard = item.standard.replace(" ","")
        item.unit = 'M'

    # 주입량(S.G.R)
    if item.name in ['주입량']:
        item.name = '주입량(S.G.R)'
        item.standard = '자재포함'
        item.unit = 'M3'

    # 기계기구 설치 및 해체(S.G.R)
    if item.name in ['기계기구 설치']:
        item.name = '기계기구 설치 및 해체(S.G.R)'
        item.unit = '회'

    # 플랜트 설치 및 해체(S.G.R)
    if item.name in ['플랜트 조립 및 해체']:
        item.name = '플랜트 설치 및 해체(S.G.R)'
        item.unit = '회'

    # 시멘트 삭제 (주입량(S.G.R)부자재)
    if item.name in ['시멘트량']:
        item.name = 'SGR' + item.name + '★삭제아이템'
        item.formula = '0'
        item.sum = '0'