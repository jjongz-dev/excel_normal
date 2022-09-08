from dataclasses import dataclass



@dataclass()
class ItemStandard:
    windows_name: str
    glass_standard: str
    glass_door: str
    fire_entrance: str
    system_door: str
    insect_screen: str
    houseHold: str
    remark: str


    def to_excel(self):
        return [self.windows_name, self.glass_standard, self.glass_door, self.fire_entrance, self.system_door, self.insect_screen, self.houseHold, self.remark]