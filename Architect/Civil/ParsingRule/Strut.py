from Architect.Civil.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 버팀보(MAIN STRUT) 연결 and 버팀보(CORNER STRUT) 연결
    if item.name in ['STRUT 설치 및 철거(H-300×300×10×15)']:
        temp_standard = item.name.split('(')[1].split(')')[0]
        if item.standard == '3M이하':
            item.name = '버팀보(MAIN STRUT) 설치 및 해체'
            item.standard = temp_standard
            item.unit = 'M'
            item.formula = '★산출서 확인 후 값 변경'
            item.sum = '★산출서 확인 후 값 변경'
        elif item.standard == '3 ~ 5M':
            item.name = '버팀보(CORNER STRUT) 설치 및 해체'
            item.standard = temp_standard
            item.unit = 'M'
            item.formula = '★산출서 확인 후 값 변경'
            item.sum = '★산출서 확인 후 값 변경'
        else:
            item.name = item.name + '★삭제'
            item.formula = '0'
            item.sum = '0'


    # JACK 설치 및 해체
    if item.name in ['스크류잭 설치 및 철거'] or item.name in ['선행하중잭 설치 및 철거'] :
        item.name = 'JACK 설치 및 해체'
        item.standard = '선행하중'
        item.unit = 'EA'

    # 보강재 (보걸이 및 BRACING)
    if item.name in ['H-형강 설치 및 철거']:
        item.name = '보강재 (보걸이 및 BRACING)'
        item.unit = 'M'
        if item.standard is not None and '(' in item.standard and ')' in item.standard:
            item.standard = item.standard.split('(')[1].split(')')[0]

    # L-형강 설치 및 철거
    if item.name in ['L-형강 설치 및 철거']:
        item.name = '보강재 (L형강 BRACING)'
        item.unit = 'M'
        if item.standard is not None and '(' in item.standard and ')' in item.standard:
            item.standard = item.standard.split('(')[1].split(')')[0]


    # H-PILE 연결
    if item.name in ['H-PILE 연결STRUT 구간']:
        item.name = '버팀보(MAIN STRUT) 연결'
        item.standard = item.standard.strip('() ')

    if item.name in ['H-PILE 연결H-BEAM 구간']:
        item.name = '버팀보(CORNER STRUT) 연결'
        item.standard = item.standard.strip('() ')
