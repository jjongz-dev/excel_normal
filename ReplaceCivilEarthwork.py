from ExcelStandard import ExcelStandard

def launch(내역:  ExcelStandard):
    # 터파기
    if 내역.품명 in ['굴착 및 상차']:
        내역.품명 = '터파기'
        내역.규격 = 내역.규격.replace(' ','')


    # 상차
    if 내역.품명 in ['운반 및 사토장 정지']:
        내역.품명 = '상차'
        내역.규격 = '백호'