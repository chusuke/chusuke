import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import openpyxl
from dataclasses import dataclass, field
from datetime import datetime, date, timedelta
import pprint


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
        self.today = datetime.today()
        self.tomorrow = self.today + timedelta(days=1)
        self.dft = self.today + timedelta(days=2)
        self.yesterday = self.today - timedelta(days=1)


def main():
    root = tk.Tk()

    style = ttk.Style()
    style.configure("TButton", font=('', 8))
    style.configure("TRadiobutton", font=('', 8))
    style.configure("TCheckbutton", font=('', 8))
    # ノートブック
    nb = ttk.Notebook(root)
    nb.pack()

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
    subframe = ttk.LabelFrame(root, text="!!!!!!")
    subframe.pack(fill=tk.X)
    ttk.Button(subframe,text="QQQQ").pack()
    t_display_product_with_owner = {}
    framearray = []
    # ttk.Label
    for t_i, t_owner in enumerate(t_all_owner_list):
        t_product = []
        t_product = [(obj.ProductName, obj.Region, obj.Grade, obj.quantity)
                     for obj in test_array if obj.OwnerName == t_owner]
        t_display_product_with_owner[t_owner] = t_product
        display_part_obj.ProductListWithOwner.append(t_display_product_with_owner)
        framearray.append(ttk.Frame(root))
        nb.add(framearray[t_i], text=t_owner, padding=20)

    root.mainloop()


if __name__ == '__main__':
    main()
