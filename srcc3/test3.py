# -*- coding: utf-8-*-
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import openpyxl
from datetime import datetime, timedelta


class Stockinglist:
    def __init__(self):
        pass

    @staticmethod
    def getfilename():
        filename = filedialog.askopenfilename()
        return filename

    @staticmethod
    def readcommonallitems(filename):
        header_cells = None
        oder_items = []
        wb = openpyxl.load_workbook(filename)
        ws = wb.worksheets[0]

        for row in ws.rows:
            if row[0].row == 1:
                header_cells = row
            else:
                row_dic = {}
                for k, v in zip(header_cells, row):
                    row_dic[k.value] = v.value
                oder_items.append(row_dic)
        return oder_items


class GuiApplication:
    def __init__(self, root, style):
        self.allitemlist = []
        self.allownerlist = []
        self.mainframe = []
        self.root = root
        self.style = style
        self.style.configure('.', font=('', 8))  # Ttk Widget のフォントの大きさを変更
        self.t_nb = ttk.Notebook(self.root)
        self.t_nb.pack()
        self.today = datetime.today()
        self.tomorrow = self.today + timedelta(days=1)
        self.dft = self.today + timedelta(days=2)
        self.t_itemarray = [("", "groove", 5), ("品名", "groove", 20), ("産地", "groove", 6),
                            ("規格", "groove", 6), ("数量", "groove", 5),
                            ("", "groove", 5), ("品名", "groove", 20), ("産地", "groove", 6),
                            ("規格", "groove", 6), ("数量", "groove", 5)]

    def make_placementwidget(self):
        self.mainframe = [ttk.Frame(self.root) for _ in self.allownerlist]

        for i, owner in enumerate(self.allownerlist):
            self.t_nb.add(self.mainframe[i], text=owner, padding=5)
            for col, arglist in enumerate(self.t_itemarray):
                a, b, c = arglist
                ttk.Label(self.mainframe[i], text=a, relief=b, width=c).grid(row=0, column=col, sticky=tk.NSEW)
            parts_product = [[tk.BooleanVar(), buyitem['品名'], buyitem['産地'], buyitem['規格'], tk.StringVar()]
                             for buyitem in self.allitemlist if buyitem['出荷者名'] == owner]

            disp_parts = []
            for parts in parts_product:
                t_parts = [[s1, s2] for s1, s2 in zip(parts, self.t_itemarray)]
                disp_parts.append(t_parts)
            for pc, s2 in enumerate(disp_parts):
                for col, s4 in enumerate(s2):
                    t_row = pc + 1 if pc + 1 < 30 else pc - 28
                    t_col = col if (pc + 1) < 30 else col + 5
                    s5, (_, s7, s8) = s4
                    if col == 0:
                        t_chk = tk.Checkbutton(self.mainframe[i], text=str(pc + 1), relief=s7, variable=s5)
                        t_chk.grid(row=t_row, column=t_col)
                    if (col == 1) or (col == 2) or (col == 3):
                        t_label = ttk.Label(self.mainframe[i], text=s5, relief=s7, width=s8)
                        t_label.grid(row=t_row, column=t_col, sticky=tk.NSEW)
                    if col == 4:
                        ent = ttk.Entry(self.mainframe[i], width=s8, textvariable=s5)
                        ent.grid(row=t_row, column=t_col, sticky=tk.NSEW)

    def setowner(self, *lists):
        self.allownerlist = [ownr for ownr in lists]

    def setallitems(self, *lists):
        self.allitemlist = [items for items in lists]

def main():
    root = tk.Tk()
    style = ttk.Style()
    stock = Stockinglist()
    g = GuiApplication(root, style)
    filename = stock.getfilename()
    if filename:
        allitemlist = stock.readcommonallitems(filename)
        g.setallitems(*allitemlist)
        ownerlist = [h["出荷者名"] for h in g.allitemlist]
        allownerlist = sorted(set(ownerlist), key=ownerlist.index)
        g.setowner(*allownerlist)
        g.make_placementwidget()
        root.mainloop()

    del g
    del root
    del style
    del stock


if __name__ == "__main__":
    main()
