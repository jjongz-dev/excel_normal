from dataclasses import dataclass


@dataclass()
class ExcelStandard:
    층: str
    호: str
    실: str
    대공종: str
    중공종: str
    코드: str
    품명: str
    규격: str
    단위: str
    부위: str
    타입: str
    산식: str
    수량: float
    Remark: str
    개소: int

    def to_excel(self):

        호 = self.호
        실 = self.실
        산식 = self.산식
        품명 = self.품명
        규격 = self.규격
        단위 = self.단위

        if self.수량 =='★산출서 확인 후 값 변경':
            수량 = self.수량
        else:
            수량 = float(self.수량)

        if type(호) is str:
            호 = 호.replace(' ', '').replace('  ', '')

        if type(실) is str:
            실 = 실.replace(' ', '').replace('  ', '')

        if type(산식) is str:
            산식 = 산식.replace('\n', '').replace('\r', '')

        if type(품명) is str:
            품명 = 품명.replace('\n', '').replace('\r', '')

        if type(규격) is str:
            규격 = 규격.replace('\n', '').replace('\r', '')

        if type(단위) is str:
            단위 = 단위.replace('\n', '').replace('\r', '').replace('m', 'M').replace('m2', 'M2').replace('m3', 'M3').replace('ton', 'TON').replace('㎡', 'M2').replace('㎥', 'M3').replace('ea', 'EA')

        return [self.층, 호, 실, self.대공종, self.중공종, self.코드, 품명, 규격, 단위, self.부위, self.타입, 산식, 수량, self.Remark, self.개소]
