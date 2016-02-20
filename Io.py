'''
Created on Jan 26, 2016
#Io.py
@author: Irwanto
'''
import Tkinter as tk
import tkFileDialog
import ttk, os

class source(object):
    def __init__(self, root, gridRow, gridColumn):
        self.root = root
        self.path = tk.StringVar()
        self.entry = tk.Entry(self.root, textvariable = self.path, width = 40,\
                              state='disabled')
        self.entry.configure(disabledbackground="white")
        self.entry.grid(row=gridRow, column=gridColumn, padx=5, pady=3, sticky='w')
        self.button = tk.Button(self.root, text='Open', command=self.openFile)
        self.button.grid(row=gridRow, column=gridColumn+1, padx=5, pady=3, sticky='w')
    def openFile(self):
        self.res = tkFileDialog.askopenfilename()
        self.path.set(self.res)
    def getFile(self):
        return self.path.get()
    
class output(source):
    def __init__(self, root, gridRow, gridColumn):
        self.root = root
        self.out = tk.Label(self.root, text='Output:')
        self.out.grid(row=gridRow-1, column=gridColumn, padx=5, sticky='w')
        super(output, self).__init__(root, gridRow, gridColumn)
    def openFile(self):
        os.system("start "+self.getFile())
        
class excelInput(source):
    def __init__(self, root, gridRow, gridColumn):
        self.root = root
        self.out = tk.Label(self.root, text='Input:')
        self.out.grid(row=gridRow-1, column=gridColumn, padx=5, sticky='w')
        
        #Sheet name
        self.sh = tk.Label(self.root, text='Sheet Name:')
        self.sh.grid(row=gridRow+1, column=gridColumn, padx=5, sticky='w')  
        #Sheet combo box                    
        self.sheetName = tk.StringVar()
        self.sheet = ttk.Combobox(self.root, textvariable=self.sheetName, state='readonly')
        self.sheet['values'] = ()
        self.sheet.grid(row=gridRow+2, column=gridColumn, padx=5, pady=5,\
                        sticky='w')
        #Product name
        self.cell = tk.Label(self.root, text='Seller SKU:')
        self.cell.grid(row=gridRow+1, column=gridColumn+1, padx=5, sticky='w')
        #sku combo box
        self.skuName = tk.StringVar()
        self.sku = ttk.Combobox(self.root, textvariable=self.skuName, state='readonly')
        self.sku['values'] = ()
        self.sku.grid(row=gridRow+2, column=gridColumn+1, padx=5, pady=5,\
                        sticky='w')
        
        #Stok
        self.cell1 = tk.Label(self.root, text='Stock:')
        self.cell1.grid(row=gridRow+3, column=gridColumn+1, padx=5, sticky='w')
        #sku combo box
        self.stockName = tk.StringVar()
        self.stock = ttk.Combobox(self.root, textvariable=self.stockName, state='readonly')
        self.stock['values'] = ()
        self.stock.grid(row=gridRow+4, column=gridColumn+1, padx=5, pady=5,\
                        sticky='w')
        
        super(excelInput, self).__init__(root, gridRow, gridColumn)
        
