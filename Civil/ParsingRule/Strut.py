from Civil.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 버팀보(MAIN STRUT) 연결 and 버팀보(CORNER STRUT) 연결
    if item.name in ['STRUT 설치 및 철거(H-300×300×10×15)']:
        item.name = item.name + '★STRUT공 확인 STRUT -> MAIN and CORNER, 본 -> M'

    # JACK 설치 및 해체
    if item.name in ['스크류잭 설치 및 철거']:
        item.name = 'JACK 설치 및 해체'
        item.standard = '선행하중	'
        item.unit = 'EA'

    # 보강재 (보걸이 및 BRACING)
    if item.name in ['H-형강 설치 및 철거']:
        item.name = '보강재 (보걸이 및 BRACING)'
        item.standard = item.standard + '★규격체크'
        item.unit = 'M'

    # H-PILE 연결
    if item.name in ['# H-PILE 연결']:
        item.name = 'STRUT' + item.name
        item.standard = item.standard + '★규격체크'
        item.unit = '개소'