from dataclasses import dataclass



@dataclass()
class ItemStandard2:
    constructionWork: str
    name: str
    standard: str
    unit: str
    sum: float


    def to_excel(self):
        return [self.constructionWork, self.name, self.standard, self.unit, self.sum]