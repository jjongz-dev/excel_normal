from dataclasses import dataclass


@dataclass()
class QuantityItemStandard2:
    name: str
    category: str
    standard: str
    unit: str
    sum: float

    def __post_init__(self):
        if self.category is None:
            self.category = ""

    def to_excel(self):
        return [self.category, self.name, self.standard, self.unit, self.sum]
