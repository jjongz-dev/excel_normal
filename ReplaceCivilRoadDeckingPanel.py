from ExcelStandard import ExcelStandard

def launch(내역:  ExcelStandard):
    # 복공판
    if '복공판' in 내역.품명:
        내역.품명 = 내역.품명 + '★품규확인'

    # 주형보 설치 및 철거
    if 내역.품명 in ['주형보 설치 및 철거']:
        내역.품명 = '주형보 설치 및 해체'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    # 주형보받침보 설치 및 철거
    if 내역.품명 in ['주형보받침보 설치 및 철거']:
        내역.품명 = '주형지지보 설치 및 해체'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    # PIECE BRACKET 설치 및 철거
    if 내역.품명 in ['PIECE BRACKET 설치 및 철거']:
        내역.품명 = '주형보 PIECE BRACKET설치'
        내역.규격 = ''

    # L-형강 설치 및 철거
    if 내역.품명 in ['L-형강 설치 및 철거']:
        내역.품명 = '주형보강재 (L형강 BRACING)'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

