from dataclasses import dataclass, field
from tkinter import filedialog
import openpyxl


@dataclass
class BuyItem:
    OwnerName: str
    ProductName: str
    Region: str
    Grade: str
    quantity: int


@dataclass
class DisplayHeader:
    column00: list = field(default_factory=lambda: ['', 5, 'groove'])
    column01: list = field(default_factory=lambda: ['品名', 20, 'groove'])
    column02: list = field(default_factory=lambda: ['産地', 6, 'groove'])
    column03: list = field(default_factory=lambda: ['規格', 6, 'groove'])
    column04: list = field(default_factory=lambda: ['数量', 4, 'groove'])


class DisplayPart:
    def __init__(self):
        self.ProductListWithOwner = []

    def make_display_data():
        filename = filedialog.askopenfilename()
        wb = openpyxl.load_workbook(filename)
        ws = wb.worksheets[0]
        test_array = []
        for i, row in enumerate(ws.rows):
            buy = []
            if i > 0:
                for v in row:
                    buy.append(v.value)
                t_items = BuyItem(buy[0], buy[1], buy[2], buy[3], buy[4])
                test_array.append(t_items)
        display_part_obj = DisplayPart()
        t_owner_list = [value.OwnerName for value in test_array]
        t_all_owner_list = sorted(set(t_owner_list), key=t_owner_list.index)
        del t_owner_list
        t_display_product_with_owner = {}
        for t_i, t_owner in enumerate(t_all_owner_list):
            t_product = [(obj.ProductName, obj.Region, obj.Grade, obj.quantity)
                         for obj in test_array if obj.OwnerName == t_owner]
            t_display_product_with_owner[t_owner] = t_product
            display_part_obj.ProductListWithOwner.append(t_display_product_with_owner)


if __name__ == '__main__':
    main()
