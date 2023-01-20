from Architect.Civil.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 천공(LW)
    if item.name in ['천   공']:
        item.name = '천공(LW)'
        item.standard = item.standard.replace(" ","")
        item.unit = 'M'

    # 주입량(LW)
    if item.name in ['L.W 주입량']:
        item.name = '주입량(LW)'
        item.standard = ''
        item.unit = 'M3'

    # SEAL제 주입량
    if item.name in ['SEAL제 주입량']:
        item.name = 'SEAL주입량(LW)'
        item.standard = ''
        item.unit = 'M3'

    # 멘젯튜브 설치
    if item.name in ['멘젯튜브 설치']:
        item.name = '맨젯튜브 설치(LW)'
        item.standard = ''
        item.unit = 'M'

    # 기계기구 설치 및 해체(LW)
    if item.name in ['기계기구 설치']:
        item.name = '기계기구 설치 및 해체(LW)'
        item.unit = '회'

    # 플랜트 설치 및 해체(LW)
    if item.name in ['플랜트 조립 및 해체']:
        item.name = '플랜트 설치 및 해체(LW)'
        item.unit = '회'

    # 시멘트 삭제 (주입량(S.G.R)부자재)
    if item.name in ['시멘트량']:
        item.name = '시멘트'
        item.standard = '40KG/포,LW용'