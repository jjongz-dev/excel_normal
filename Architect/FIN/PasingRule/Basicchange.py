from Architect.FIN.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # TYPE변경
    type1 = item.type.split('-')
    if len(type1) == 2:
        item.type = item.type.split('-')[0]

    if item.type == '토공':
        item.type = '외부'

    if item.type == '가설':
        item.type = '내부'

    # 품명변경
    if item.name.startswith("★"):
        item.name = item.name.replace("★", "")

    # 창호
    if item.type == '창호':
        item.part = item.floor
        item.floor = ''
        item.roomname = ''

   #외부 입면 호실정리
    if item.type == '외부' and item.floor.__contains__('정면'):
        item.location = '정면'
        item.roomname = '정면'
        item.floor = ''

    if item.type == '외부' and item.floor.__contains__('배면'):
        item.location = '배면'
        item.roomname = '배면'
        item.floor = ''

    if item.type == '외부' and item.floor.__contains__('좌측면'):
        item.location = '좌측면'
        item.roomname = '좌측면'
        item.floor = ''

    if item.type == '외부' and item.floor.__contains__('우측면'):
        item.location = '우측면'
        item.roomname = '우측면'
        item.floor = ''

    if item.type == '외부' and item.floor.__contains__('남측'):
        item.location = '남측면'
        item.roomname = '남측면'
        item.floor = ''

    if item.type == '외부' and item.floor.__contains__('북측'):
        item.location = '북측면'
        item.roomname = '북측면'
        item.floor = ''

    if item.type == '외부' and item.floor.__contains__('서측'):
        item.location = '서측면'
        item.roomname = '서측면'
        item.floor = ''

    if item.type == '외부' and item.floor.__contains__('동측면'):
        item.location = '동측면'
        item.roomname = '동측면'
        item.floor = ''

    #토공사정리, 기초단열재
    if item.floor == '토공사' or item.floor == '기초단열재':
        item.location = '기초하부'
        item.roomname = '기초하부'
        item.floor = 'FT'

    # 기본 층정리
    if (item.name in ['가설컨테이너반입', '가설컨테이너반출', '가설수도', '가설전기', '가설울타리설치', '가설울타리해체', '가설출입구설치', '가설출입구해체',
                      '건축폐기물처리', '건축허가표지판', '경계측량및현황측량', '민원처리', '이동식가설화장실반입', '이동식가설화장실반출', '준공청소', '지내력시험',
                      '규준틀설치']):
        item.location = '공통가설'
        item.roomname = '공통가설'
        item.floor = '1F'

    # 조경공사
    if item.floor.__contains__('조경'):
        item.location = '조경'
        item.roomname = '조경'
        item.floor = '1F'

    # 부대토목공사
    if item.floor.__contains__('부대토목'):
        item.location = '부대토목'
        item.roomname = '부대토목'
        item.floor = '1F'

    # 철거
    if item.floor.__contains__('철거'):
        item.location = '철거'
        item.roomname = '철거'
        item.floor = '1F'

    # 기타공사
    if item.floor.__contains__('기타공사'):
        item.location = '기타공사'
        item.roomname = '기타공사'
        item.floor = '1F'

    # 포장공사
    if item.floor.__contains__('포장공사'):
        item.location = '포장공사'
        item.roomname = '포장공사'
        item.floor = '1F'

    # 우오수공사
    if item.floor.__contains__('우오수공사'):
        item.location = '우오수공사'
        item.roomname = '우오수공사'
        item.floor = '1F'

    # 정화조설치공사
    if item.floor.__contains__('정화조'):
        item.location = '정화조'
        item.roomname = '정화조'
        item.floor = 'FT'