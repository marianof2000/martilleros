import xlrd
from tkinter import ttk
import tkinter as tk

book = xlrd.open_workbook('PADRON_FERNANDA_30-09-13.xls')
sh = book.sheet_by_index(0) # Indicamos la primera hoja

print ("The number of worksheets is", book.nsheets)
print ("Worksheet name(s):", book.sheet_names())
sh = book.sheet_by_index(0)
print (sh.name, sh.nrows, sh.ncols)
print ("Cell (2,0) is: ", sh.cell_value(rowx=2, colx=0))
for rx in range(sh.nrows):
    print (sh.row(rx))

class Application(ttk.Frame):
   
    def __init__(self, main_window):
        super().__init__(main_window)
        main_window.title("Combobox")
       
        self.combo = ttk.Combobox(self)
        self.combo.place(x=50, y=50)
        self.combo = ttk.Combobox(self, state="readonly")
        self.combo["values"] = ["Python", "C", "C++", "Java"]
       
        main_window.configure(width=300, height=200)
        self.place(width=300, height=200)
main_window = tk.Tk()
app = Application(main_window)
app.mainloop()
