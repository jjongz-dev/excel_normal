from Architect.Civil.ItemStandard import ItemStandard



def launch(item: ItemStandard):
    # 복공판
    if '복공판' in item.name:
        item.name = item.name + '★품규확인'

    # 주형보 설치 및 철거
    if item.name in ['주형보 설치 및 철거']:
        item.name = '주형보 설치 및 해체'
        if item.standard is not None and '(' in item.standard and ')' in item.standard:
            item.standard = item.standard.split('(')[1].split(')')[0]

    # 주형보받침보 설치 및 철거
    if item.name in ['주형보받침보 설치 및 철거']:
        item.name = '주형지지보 설치 및 해체'
        if item.standard is not None and '(' in item.standard and ')' in item.standard:
            item.standard = item.standard.split('(')[1].split(')')[0]

    # PIECE BRACKET 설치 및 철거
    if item.name in ['PIECE BRACKET 설치 및 철거']:
        item.name = '주형보 PIECE BRACKET설치'
        item.standard = ''

    # L-형강 설치 및 철거
    if item.name in ['L-형강 설치 및 철거']:
        item.name = '주형보강재 (L형강 BRACING)'
        if item.standard is not None and '(' in item.standard and ')' in item.standard:
            item.standard = item.standard.split('(')[1].split(')')[0]

