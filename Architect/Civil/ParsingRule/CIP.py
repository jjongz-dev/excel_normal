from Architect.Civil.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # CIP 천공
    if item.name in ['C.I.P 천공']:
        item.name = 'CIP 천공'
        item.standard = item.standard.replace(" ","") + '★관경체크'
        item.unit = 'M'

    # 케이싱 삭제 (SIDE-PILE+CIP천공 부자재)
    if item.name in ['케이싱 설치']:
        item.name = '케이싱 설치 및 해체 ★삭제아이템'
        item.formula = '0'
        item.sum = '0'

    # 철근콘크리트용봉강
    if item.name in ['철근 가공 및 조립(C.I.P)']:
        item.name = '철근콘크리트용봉강'
        item.standard = item.standard + ',SD400,지정장소도,CIP용'
        item.unit = 'TON'

    if item.name in ['철근 가공 및 조립(CAP BEAM)']:
        item.name = '철근콘크리트용봉강'
        item.standard = item.standard + ',SD400,지정장소도,CAP BEAM용'
        item.unit = 'TON'

    # 레미콘
    if item.name in ["CON'C 타설C.I.P"]:
        item.name = '레미콘'
        item.standard = item.standard + ',CIP용 ★강도체크'
        item.unit = 'M3'

    if item.name in ["CON'C 타설CAP BEAM"]:
        item.name = '레미콘'
        item.standard = item.standard + ',CAP BEAM용 ★강도체크'
        item.unit = 'M3'

    # 거푸집
    if item.name in ['거푸집 설치']:
        item.name = '거푸집 설치및해체'
        item.standard = ''
        item.unit = 'M2'

    # 발생토 처리
    if item.name in ['발생토 처리']:
        item.name = '슬라임 처리'
        item.standard = ''
        item.unit = 'M3'