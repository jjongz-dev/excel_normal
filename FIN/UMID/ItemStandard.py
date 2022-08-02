from dataclasses import dataclass



@dataclass()
class ItemStandard:
    floor: str
    location: str
    roomname: str



    name: str
    standard: str
    unit: str
    part: str
    type: str
    formula: str
    sum: float


    def to_excel(self):
        return [self.floor, self.location, self.roomname, '건축', '', '', self.name, self.standard, self.unit, '', self.type, self.formula, self.sum, '']