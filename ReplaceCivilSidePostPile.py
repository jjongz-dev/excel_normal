from ExcelStandard import ExcelStandard

def launch(내역:  ExcelStandard):
    # SIDE-PILE천공
    if 'SIDE-PILE 천공' in 내역.품명:
        임시규격 = 내역.품명.split('(')[1].split(')')[0]
        내역.품명 = 'SIDE-PILE천공'
        내역.규격 = 임시규격 + ',' + 내역.규격.replace(" ", "")

    # SIDE-PILE박기
    if 내역.품명 in ['SIDE-PILE 박기']:
        내역.품명 = 'SIDE-PILE박기'
        내역.단위 = '개소'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    # SIDE-PILE사장
    if 내역.품명 in ['SIDE-PILE 사장']:
        내역.품명 = 내역.품명 + '★삭제아이템'
        내역.산식 = '0'
        내역.수량 = '0'

    # POST-PILE천공
    if 'POST PILE 천공' in 내역.품명:
        임시규격 = 내역.품명.split('(')[1].split(')')[0]
        내역.품명 = 'POST-PILE천공'
        내역.규격 = 임시규격 + ',' + 내역.규격.replace(" ", "")

    # POST-PILE박기
    if 내역.품명 in ['POST PILE 박기']:
        내역.품명 = 'POST-PILE박기'
        내역.단위 = '개소'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    # POST-PILE인발
    if 내역.품명 in ['POST PILE 인발']:
        내역.품명 = 'POST-PILE인발'
        내역.단위 = '개소'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    # POST-PILE절단
    try:
        if 내역.품명 in ['POST PILE 절단 및 사장']:
            내역.품명 = 'POST-PILE절단'
            내역.단위 = '개소'
            if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
                내역.규격 = 내역.규격.split('(')[1].split(')')[0]
    except Exception as e:
        print('예외가 발생했습니다.', e)
        print('POST-PILE사장 error!!!' + str(item))

    # SIDE-PILE연결
    if 내역.품명 in ['H-PILE 연결SIDE-PILE']:
        내역.품명 = 'SIDE-PILE연결'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    # POST-PILE연결
    if 내역.품명 in ['H-PILE 연결POST PILE']:
        내역.품명 = 'POST-PILE연결'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]


    # 띠장(WALE)설치 및 해체
    if 내역.품명 in ['띠장(WALE)설치 및 철거']:
        내역.품명 = '띠장(WALE)설치 및 해체'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    # 브라켓
    if 내역.품명 in ['BRACKET 설치'] and 내역.규격 in ['STRUT 구간']:
        내역.품명 = 'BRACKET설치(SIDE-PILE+WALE)'
        내역.규격 = ''

    if 내역.품명 in ['BRACKET 설치'] and 내역.규격 in ['POST PILE 구간']:
        내역.품명 = 'PIECE BRACKET설치(STRUT+POST-PILE)'
        내역.규격 = ''

    # 띠장(WALE)연결
    if 내역.품명 in ['띠장(WALE) 연결']:
        내역.품명 = '띠장(WALE)연결'
        내역.규격 = 내역.규격.replace(' ','')

    # 스티프너
    if 내역.품명 in ['스티프너 설치 및 철거']:
        내역.품명 = '스티프너 설치 및 해체'
        내역.규격 = ''