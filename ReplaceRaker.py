from ExcelStandard import ExcelStandard

def launch(내역:  ExcelStandard):

    if 'RAKER구간 굴착' in 내역.품명:
        내역.품명 = 'RAKER 터파기'
        내역.규격 = 내역.규격.replace(" ","")
        내역.단위 = 'M3'

    if "CON'C 타설" in 내역.품명:
        내역.품명 = 'RAKER타설'
        내역.규격 = '무근,인력'
        내역.단위 = 'M3'
