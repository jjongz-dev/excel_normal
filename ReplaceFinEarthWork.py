from ExcelStandard import ExcelStandard

def launch(내역:  ExcelStandard):

    if '우편물반송함' in 내역.품명:
        내역.품명 = '우편물수취함'
        내역.산식 = '<우편물반송함>'+내역.산식

    if '폐건전지수거함' in 내역.품명:
        내역.품명 = '우편물수취함'
        내역.산식 = '<폐건전지수거함>'+내역.산식
