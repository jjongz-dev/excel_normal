from dataclasses import dataclass



@dataclass()
class ItemStandard:






    name: str
    standard: str
    unit: str


    formula: str
    sum: float


    def to_excel(self):
        return ['', '', '', '토목', '', '', self.name, self.standard, self.unit, '', '외부', self.formula, self.sum, '']