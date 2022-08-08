from FIN.ItemStandard import ItemStandard


def launch(item: ItemStandard):
    # 기본 층정리
    if (item.name in ['가설컨테이너반입', '가설컨테이너반출', '가설수도', '가설전기', '가설울타리설치', '가설울타리해체', '가설출입구설치', '가설출입구해체',
                      '건축폐기물처리', '건축허가표지판', '경계측량및현황측량', '민원처리', '이동식가설화장실반입', '이동식가설화장실반출', '준공청소', '지내력시험', '규준틀설치']):
        item.location = '공통가설'
        item.roomname = '공통가설'
        item.floor = '1F'
