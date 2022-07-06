from dataclasses import dataclass


@dataclass()
class QuantityItemStandard2:
    name: str
    standard: str
    unit: str
    sum: float

    def to_excel(self):
        return ['', self.name, self.standard, self.unit, self.sum]
