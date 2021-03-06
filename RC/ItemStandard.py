from dataclasses import dataclass



@dataclass()
class ItemStandard:
    floor: str
    location: str
    name: str
    standard: str
    part: str
    formula: str
    sum: float

    def to_excel(self):
        return [self.floor, self.location, '', '건축', '철근콘크리트공사', '', self.name, self.standard, '', self.part, '구조',self.formula, self.sum, '']
