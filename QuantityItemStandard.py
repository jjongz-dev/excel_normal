from dataclasses import dataclass


@dataclass()
class QuantityItemStandard:
    name: str
    standard: str
    unit: str
    formula: str
    unit_formula: str
    sum: float

    def to_excel(self):
        return ['', '', '', '전기', '', '', self.name, self.standard, self.unit, '', '전기',
                '({})*({})'.format(self.formula, self.unit_formula), self.sum, '']
