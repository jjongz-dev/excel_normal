from dataclasses import dataclass


@dataclass()
class ExcelGroup:
    중공종: str
    품명: str
    규격: str
    단위: str
    수량: float

    def to_excelGroup(self):

        중공종 = self.중공종
        품명 = self.품명
        규격 = self.규격
        단위 = self.단위
        수량 = float(self.수량)

        if type(품명) is str:
            품명 = 품명.replace('\n', '').replace('\r', '')

        if type(규격) is str:
            규격 = 규격.replace('\n', '').replace('\r', '')

        if type(단위) is str:
            단위 = 단위.replace('\n', '').replace('\r', '').replace('m', 'M').replace('m2', 'M2').replace('m3', 'M3').replace('ton', 'TON').replace('㎡', 'M2').replace('㎥', 'M3').replace('ea', 'EA')

        return [중공종, 품명, 규격, 단위, 수량]
