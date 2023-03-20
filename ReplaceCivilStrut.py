from ExcelStandard import ExcelStandard

def launch(내역:  ExcelStandard):
    # 버팀보(MAIN STRUT) 연결 and 버팀보(CORNER STRUT) 연결
    if 내역.품명 in ['STRUT 설치 및 철거(H-300×300×10×15)']:
        임시규격 = 내역.품명.split('(')[1].split(')')[0]
        if 내역.규격 == '3M이하':
            내역.품명 = '버팀보(MAIN STRUT) 설치 및 해체'
            내역.규격 = 임시규격
            내역.단위 = 'M'
            내역.산식 = '★산출서 확인 후 값 변경'
            내역.수량 = '★산출서 확인 후 값 변경'
        elif 내역.규격 == '3 ~ 5M':
            내역.품명 = '버팀보(CORNER STRUT) 설치 및 해체'
            내역.규격 = 임시규격
            내역.단위 = 'M'
            내역.산식 = '★산출서 확인 후 값 변경'
            내역.수량 = '★산출서 확인 후 값 변경'
        else:
            내역.품명 = 내역.품명 + '★삭제'
            내역.산식 = '0'
            내역.수량 = '0'


    # JACK 설치 및 해체
    if 내역.품명 in ['스크류잭 설치 및 철거'] or 내역.품명 in ['선행하중잭 설치 및 철거'] :
        내역.품명 = 'JACK 설치 및 해체'
        내역.규격 = '선행하중'
        내역.단위 = 'EA'

    # 보강재 (보걸이 및 BRACING)
    if 내역.품명 in ['H-형강 설치 및 철거']:
        내역.품명 = '보강재 (보걸이 및 BRACING)'
        내역.단위 = 'M'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    # L-형강 설치 및 철거
    if 내역.품명 in ['L-형강 설치 및 철거']:
        내역.품명 = '보강재 (L형강 BRACING)'
        내역.단위 = 'M'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]


    # H-PILE 연결
    if 내역.품명 in ['H-PILE 연결STRUT 구간']:
        내역.품명 = '버팀보(MAIN STRUT) 연결'
        내역.규격 = 내역.규격.strip('() ')

    if 내역.품명 in ['H-PILE 연결H-BEAM 구간']:
        내역.품명 = '버팀보(CORNER STRUT) 연결'
        내역.규격 = 내역.규격.strip('() ')
