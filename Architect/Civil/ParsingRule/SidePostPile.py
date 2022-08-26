from Architect.Civil.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # SIDE-PILE천공
    if item.name.__contains__('SIDE-PILE 천공'):
        aaa = item.name.split('(')[1].split(')')[0]
        item.name = 'SIDE-PILE천공' + '(' + item.standard.replace(" ", "") + ')'
        item.standard = aaa
        item.unit = 'M'

    # SIDE-PILE박기
    if item.name in ['SIDE-PILE 박기']:
        item.name = 'SIDE-PILE박기'
        item.standard = item.standard.split('(')[1].split(')')[0]
        item.unit = '개소'

    # SIDE-PILE사장
    if item.name in ['SIDE-PILE 사장']:
        item.name = item.name + '★삭제아이템'
        item.formula = '0'
        item.sum = '0'

    # POST-PILE천공
    if item.name.__contains__('POST PILE 천공'):
        aaa = item.name.split('(')[1].split(')')[0]
        item.name = 'POST-PILE천공' + '(' + item.standard.replace(" ", "") + ')'
        item.standard = aaa
        item.unit = 'M'

    # POST-PILE박기
    if item.name in ['POST PILE 박기']:
        item.name = 'POST-PILE박기'
        item.standard = item.standard.split('(')[1].split(')')[0]
        item.unit = '개소'

    # POST-PILE인발
    if item.name in ['POST PILE 인발']:
        item.name = 'POST-PILE인발'
        item.standard = item.standard.split('(')[1].split(')')[0] + '★인발길이확인 후 개소 -> M'


    # POST-PILE사장
    if item.name in ['POST PILE 절단 및 사장']:
        item.name = 'POST-PILE절단'
        item.standard = item.standard.split('(')[1].split(')')[0]
        item.unit = '개소'

    # SIDE-PILE연결
    if item.name in ['H-PILE 연결SIDE-PILE']:
        item.name = 'SIDE-PILE연결'
        item.standard = item.standard.split('(')[1].split(')')[0]
        item.unit = '개소'

    # POST-PILE연결
    if item.name in ['H-PILE 연결POST PILE']:
        item.name = 'POST-PILE연결'
        item.standard = item.standard.split('(')[1].split(')')[0]
        item.unit = '개소'

    # 띠장(WALE)설치 및 해체
    if item.name in ['띠장(WALE)설치 및 철거']:
        item.name = '띠장(WALE)설치 및 해체'
        item.standard = item.standard.split('(')[1].split(')')[0]
        item.unit = 'M'

    # 브라켓
    if item.name in ['BRACKET 설치'] and item.standard in ['STRUT 구간']:
        item.name = 'BRACKET설치(SIDE-PILE+WALE)'
        item.standard = item.standard + '★규격 확인'
        item.unit = 'EA'

    if item.name in ['BRACKET 설치'] and item.standard in ['POST PILE 구간']:
        item.name = 'PIECE BRACKET설치(STRUT+POST-PILE)'
        item.standard = ''
        item.unit = 'EA'

    # 띠장(WALE)연결
    if item.name in ['띠장(WALE) 연결']:
        item.name = '띠장(WALE)연결'
        item.standard = item.standard.replace(' ','')
        item.unit = '개소'

    # 스티프너
    if item.name in ['스티프너 설치 및 철거']:
        item.name = '스티프너 설치 및 해체'
        item.standard = ''
        item.unit = 'EA'