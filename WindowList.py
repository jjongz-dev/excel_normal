from dataclasses import dataclass



@dataclass()
class WindowList:
    구분: str
    창호명: str
    가로: float
    세로: float
    면적: float
    공제면적: float
    BASE길이: str
    면적공식: str
    도어윈도우: str
    비고: str
    합계: float

    def to_excel(self):
        return [self.구분, self.창호명, self.가로, self.세로, self.면적, self.공제면적, self.BASE길이, self.면적공식, self.도어윈도우, self.비고, self.합계]