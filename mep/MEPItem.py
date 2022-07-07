from dataclasses import dataclass


@dataclass()
class MEPItem:
    name: str
    standard: str
    unit: str
    formula: str
    unit_formula: str
    sum: float

    def to_excel(self):
        return ['', '', '', '기계', '', '', self.name, self.standard, self.unit, '', '기계',
                '({})*({})'.format(self.formula, self.unit_formula), self.sum, '']
