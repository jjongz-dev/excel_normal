from Architect.Civil.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # CIP 천공
    if item.name in ['C.I.P 천공']:
        item.name = 'CIP 천공'
        item.standard = item.standard.replace(" ","")

    # 케이싱 삭제 (SIDE-PILE+CIP천공 부자재)
    if item.name in ['케이싱 설치']:
        item.name = '케이싱 설치 및 해체'
        item.standard = item.standard + ',자재손료 포함'

    # 철근콘크리트용봉강
    if item.name in ['철근 가공 및 조립(C.I.P)']:
        item.name = '철근콘크리트용봉강'
        item.standard = item.standard + ',SD400,지정장소도,CIP용'

    if item.name in ['철근 가공 및 조립(CAP BEAM)']:
        item.name = '철근콘크리트용봉강'
        item.standard = item.standard + ',SD400,지정장소도,CAP BEAM용'

    # 레미콘
    if item.name in ["CON'C 타설C.I.P"]:
        item.name = '레미콘'
        if '210' in item.standard:
            item.standard = '25-21-12,CIP용'
        elif '240' in item.standard:
            item.standard = '25-24-12,CIP용'
        else:
            item.standard = item.standard + ',CIP용 ★파이선 규격추가'

    if item.name in ["CON'C 타설CAP BEAM"]:
        item.name = '레미콘'
        if '210' in item.standard:
            item.standard = '25-21-12,CAP BEAM용'
        elif '240' in item.standard:
            item.standard = '25-24-12,CAP BEAM용'
        else:
            item.standard = item.standard + ',CAP BEAM용 ★파이선 규격추가'

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