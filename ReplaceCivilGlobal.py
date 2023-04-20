from ExcelStandard import ExcelStandard

def launch(내역:  ExcelStandard):

    if '굴착 및 직상차' in 내역.품명 and 내역.중공종 =='토공':
        내역.규격 = 내역.품명.split('(')[-1].replace(')', '').strip()
        내역.품명 = '터파기'

    if '운반 및 사토장 정지' in 내역.품명 and 내역.중공종 =='토공':
        내역.품명 = '상차'
        내역.규격 = '백호'


#SIDE-POST-PILE
    if 'SIDE-PILE 천공' in 내역.품명 and 내역.중공종 == '가시설공':
        규격추가 = 내역.품명.split('(')[-1].replace(')', '').strip()
        내역.규격 = f'{내역.규격},{규격추가}'
        내역.품명 = 'SIDE-PILE천공'

    if 'POST PILE 천공 (' in 내역.품명 and 내역.중공종 == '가시설공':
        규격추가 = 내역.품명.split('(')[-1].replace(')', '').strip()
        내역.규격 = f'{내역.규격},{규격추가}'
        내역.품명 = 내역.품명.split('(')[0].replace('POST PILE', 'POST-PILE').replace(' ', '')

    if 'POST PILE' in 내역.품명 and 내역.중공종 == '가시설공':
        내역.품명 = 내역.품명.replace('POST PILE', 'POST-PILE').replace(' ', '')

    if 'SIDE-PILE 천공' in 내역.품명:
        임시규격 = 내역.품명.split('(')[1].split(')')[0]
        내역.품명 = 'SIDE-PILE천공'
        내역.규격 = 임시규격 + ',' + 내역.규격.replace(" ", "")

    if 'SIDE-PILE 박기' in 내역.품명:
        내역.품명 = 'SIDE-PILE박기'
        내역.단위 = '개소'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if 'SIDE-PILE 사장' in 내역.품명:
        내역.품명 = 내역.품명 + '★삭제아이템'
        내역.산식 = '0'
        내역.수량 = '0'

    if 'POSTPILE천공' in 내역.품명.replace(' ', ''):
        임시규격 = 내역.품명.split('(')[1].split(')')[0]
        내역.품명 = 'POST-PILE천공'
        내역.규격 = 임시규격 + ',' + 내역.규격.replace(" ", "")

    if 'POSTPILE박기' in 내역.품명.replace(' ', ''):
        내역.품명 = 'POST-PILE박기'
        내역.단위 = '개소'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if 'POSTPILE인발' in 내역.품명.replace(' ', ''):
        내역.품명 = 'POST-PILE인발'
        내역.단위 = '개소'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if 'POSTPILE절단및사장' in 내역.품명.replace(' ', ''):
        내역.품명 = 'POST-PILE절단'
        내역.단위 = '개소'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if 'H-PILE연결SIDE-PILE' in 내역.품명.replace(' ', ''):
        내역.품명 = 'SIDE-PILE연결'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if 'H-PILE연결POSTPILE' in 내역.품명.replace(' ', ''):
        내역.품명 = 'POST-PILE연결'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if '띠장(WALE)설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = '띠장(WALE)설치 및 해체'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if '띠장(WALE)설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = '띠장(WALE)설치 및 해체'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if 'BRACKET설치' in 내역.품명.replace(' ', '') and 'STRUT구간' in 내역.규격.replace(' ', ''):
        내역.품명 = 'BRACKET설치(SIDE-PILE+WALE)'
        내역.규격 = ''

    if 'BRACKET설치' in 내역.품명.replace(' ', '') and 'POSTPILE구간' in 내역.규격.replace(' ', ''):
        내역.품명 = 'PIECE BRACKET설치(STRUT+POST-PILE)'
        내역.규격 = ''

    if '띠장(WALE)연결' in 내역.품명.replace(' ', ''):
        내역.품명 = '띠장(WALE)연결'
        내역.규격 = 내역.규격.replace(' ','')

    if '스티프너설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = '스티프너 설치 및 해체'
        내역.규격 = ''


#CIP
    if 'C.I.P천공' in 내역.품명.replace(' ', ''):
        내역.규격 = 내역.품명.split('(')[-1].replace(')', '').strip()
        내역.품명= 'CIP 천공'
        내역.규격 = 내역.규격.replace(" ", "")

    if '케이싱설치' in 내역.품명.replace(' ', ''):
        내역.품명= '케이싱 설치 및 해체'
        내역.규격 = 내역.규격 + ',자재손료 포함'

    if '철근가공및조립(C.I.P)' in 내역.품명.replace(' ', ''):
        내역.품명= '철근콘크리트용봉강'
        내역.규격 = 내역.규격 + ',SD400,지정장소도,CIP용'

    if '철근가공및조립(CAP BEAM)' in 내역.품명.replace(' ', ''):
        내역.품명= '철근콘크리트용봉강'
        내역.규격 = 내역.규격 + ',SD400,지정장소도,CAP BEAM용'

    if "CON'C타설C.I.P" in 내역.품명.replace('(', '').replace(')', '').replace(' ', ''):
        내역.품명= '레미콘'
        if '210' in 내역.규격.replace(' ', ''):
            내역.규격 = '25-21-12,CIP용'
        elif '240' in 내역.규격.replace(' ', ''):
            내역.규격 = '25-24-12,CIP용'
        else:
            내역.규격 = 내역.규격 + ',CIP용 ★파이선 규격추가'

    if "CON'C타설CAPBEAM" in 내역.품명.replace('(', '').replace(')', '').replace(' ', ''):
        내역.품명= '레미콘'
        if '210' in 내역.규격:
            내역.규격 = '25-21-12,CAP BEAM용'
        elif '240' in 내역.규격:
            내역.규격 = '25-24-12,CAP BEAM용'
        else:
            내역.규격 = 내역.규격 + ',CAP BEAM용 ★파이선 규격추가'

    if '거푸집설치' in 내역.품명.replace(' ', ''):
        내역.품명= '거푸집 설치및해체'
        내역.규격 = ''
        내역.단위 = 'M2'

    if '발생토처리' in 내역.품명.replace(' ', ''):
        내역.품명= '슬라임 처리'
        내역.규격 = ''
        내역.단위 = 'M3'

#STRUT공
    if 'STRUT설치및철거(H-300×300×10×15)' in 내역.품명.replace(' ', ''):
        임시규격 = 내역.품명.split('(')[1].split(')')[0]
        if 내역.규격.replace(' ', '') == '3M이하':
            내역.품명 = '버팀보(MAIN STRUT) 설치 및 해체'
            내역.규격 = 임시규격
            내역.단위 = 'M'
            내역.산식 = '★산출서 확인 후 값 변경'
            내역.수량 = '★산출서 확인 후 값 변경'
        elif 내역.규격.replace(' ', '') == '3~5M':
            내역.품명 = '버팀보(CORNER STRUT) 설치 및 해체'
            내역.규격 = 임시규격
            내역.단위 = 'M'
            내역.산식 = '★산출서 확인 후 값 변경'
            내역.수량 = '★산출서 확인 후 값 변경'
        else:
            내역.품명 = 내역.품명 + '★삭제'
            내역.산식 = '0'
            내역.수량 = '0'

    if '스크류잭설치및철거' in 내역.품명.replace(' ', '') or '선행하중잭설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = 'JACK 설치 및 해체'
        내역.규격 = '선행하중'
        내역.단위 = 'EA'

    if 'H-형강설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = '보강재 (보걸이 및 BRACING)'
        내역.단위 = 'M'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if 'L-형강설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = '보강재 (L형강 BRACING)'
        내역.단위 = 'M'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if 'H-PILE연결STRUT' in 내역.품명.replace(' ', '').replace('(','').replace(')',''):
        내역.품명 = '버팀보(MAIN STRUT) 연결'
        내역.규격 = 내역.규격.strip('() ')

    if 'H-PILE연결H-BEAM' in 내역.품명.replace(' ', '').replace('(','').replace(')',''):
        내역.품명 = '버팀보(CORNER STRUT) 연결'
        내역.규격 = 내역.규격.strip('() ')


#S.G.R공 집계표
    if '천공' in 내역.품명.replace(' ', '') and 'S.G.R' in 내역.중공종:
        내역.규격 = 내역.품명.split('(')[-1].replace(')', '').strip()
        내역.품명 = '천공(S.G.R)'
        내역.단위 = 'M'

    if '주입량' in 내역.품명.replace(' ', '') and 'S.G.R' in 내역.중공종:
        내역.품명 = '주입량(S.G.R)'
        내역.규격 = ''
        내역.단위 = 'M3'

    if '기계기구설치' in 내역.품명.replace(' ', '') and 'S.G.R' in 내역.중공종:
        내역.품명 = '기계기구 설치 및 해체(S.G.R)'
        내역.단위 = '회'

    if '플랜트조립및해체' in 내역.품명.replace(' ', '') and 'S.G.R' in 내역.중공종:
        내역.품명 = '플랜트 설치 및 해체(S.G.R)'
        내역.단위 = '회'

    if '시멘트량' in 내역.품명.replace(' ', '') and 'S.G.R' in 내역.중공종:
        내역.품명 = '시멘트'
        내역.규격 = '40KG/포,S.G.R용'


# 복공
    if '복공판' in 내역.품명:
        내역.품명 = 내역.품명 + '★품규확인'

    if '주형보설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = '주형보 설치 및 해체'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if '주형보받침보설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = '주형지지보 설치 및 해체'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]

    if 'PIECEBRACKET설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = '주형보 PIECE BRACKET설치'
        내역.규격 = ''

    if 'L-형강설치및철거' in 내역.품명.replace(' ', ''):
        내역.품명 = '주형보강재 (L형강 BRACING)'
        if 내역.규격 is not None and '(' in 내역.규격 and ')' in 내역.규격:
            내역.규격 = 내역.규격.split('(')[1].split(')')[0]






