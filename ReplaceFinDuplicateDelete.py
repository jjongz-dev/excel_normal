from ExcelStandard import ExcelStandard

def launch(내역:  ExcelStandard):

    if '구조이기' in 내역.호:
        내역.품명 = 내역.품명 + '★삭제(구조이기물량)'
        내역.산식 = '0'
        내역.수량 = '0'

    if '타일면보양' in 내역.품명:
        내역.품명 = 내역.품명 + '★부자재삭제'
        내역.산식 = '0'
        내역.수량 = '0'

    if '석재면보양' in 내역.품명:
        내역.품명 = 내역.품명 + '★부자재삭제'
        내역.산식 = '0'
        내역.수량 = '0'

    if '유리끼우기및닦기' in 내역.품명:
        내역.품명 = 내역.품명 + '★부자재삭제'
        내역.산식 = '0'
        내역.수량 = '0'

    if '크러쉬버튼설치(소방관진입창 유리파괴장치)' in 내역.품명:
        내역.품명 = 내역.품명 + '★HB기준삭제'
        내역.산식 = '0'
        내역.수량 = '0'


    if '★크러쉬버튼설치(소방관진입창 유리파괴장치)' in 내역.품명:
        내역.품명 = 내역.품명 + '★HB기준삭제'
        내역.산식 = '0'
        내역.수량 = '0'