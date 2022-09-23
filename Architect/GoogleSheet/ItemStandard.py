from dataclasses import dataclass



@dataclass()
class ItemStandard:
    windows_name: str
    glass_standard: str
    fire_entrance: str
    glass_door: str
    insect_screen: str
    remark: str


    def to_excel(self):
        return [self.windows_name, self.glass_standard, self.fire_entrance, self.glass_door, self.insect_screen, self.remark]