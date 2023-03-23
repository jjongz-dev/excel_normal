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

        가로 = float(self.가로)
        세로 = float(self.세로)
        면적 = float(self.면적)
        공제면적 = float(self.공제면적)
        합계 = float(self.합계)

        return [self.구분, self.창호명, 가로, 세로, 면적, 공제면적, self.BASE길이, self.면적공식, self.도어윈도우, self.비고, 합계]