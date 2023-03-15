from dataclasses import dataclass


# head_title = ['층', '호', '실', '대공종', '중공종', '코드', '품명', '규격', '단위', '부위', '타입', '산식', '수량', 'Remark']

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
        return [self.층, self.호, self.실, self.대공종, self.중공종, self.코드,self.품명, self.규격,self.단위,self.부위,self.타입,self.산식,self.수량,self.Remark
            ,self.개소]