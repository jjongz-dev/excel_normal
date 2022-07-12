from dataclasses import dataclass



@dataclass()
class ItemStandard:
    floor: str
    name: str
    standard: str
    part: str
    formula: str
    roomname: str
    type: str
    unit: str
    sum: float


    def to_excel(self):
        return [self.floor, '', self.roomname, '건축', '', '', self.name, self.standard, self.unit, self.part, self.type, self.formula, self.sum, '']