'''
Created on Jan 26, 2016

@author: Irwanto
'''
import Tkinter as tk
import bs4, openpyxl, Io
from ttk import LabelFrame
from Proccess import calc

class window(object):
    def __init__(self):
        self.version = 'V 1.0.0'
        self.root = tk.Tk()
        self.root.title('Olshop Scrapper ' + self.version)

        self.frame1 = tk.Frame(self.root)
        self.frame1.grid(row=0, column=0, sticky='snew')
        
        #Label Frame
        self.west = LabelFrame(self.frame1, text=' Source ')
        self.west.grid(column=0, row=0, sticky='wn', padx=5, pady=5)
        self.east = LabelFrame(self.frame1, text=' Excel ')
        self.east.grid(column=1, row=0, sticky='wn', padx=5, pady=5)

        #Making Source form
        self.src = []
        for i in range(6):
            self.src.append(Io.source(self.west, i, 0))
            
        #Output form
        self.outputName = Io.output(self.east, 1, 0)
        self.outputName.button.configure(state='disabled')
        self.excel = Io.excelInput(self.east, 3, 0)
        self.excel.sheet.bind('<<ComboboxSelected>>',self.check)
        self.work = calc(self.frame1, 1, 0, self.compare, self.load)
        self.work.proccess.configure(state='disabled')
        
        self.root.mainloop()
        
    def check(self, event): 
        self.work.update_status('Getting cell information...')
        self.workingSheet = self.wb.get_sheet_by_name(self.excel.sheet.get())
        self.excel.sku['values'] = [self.workingSheet.cell(row=1, column=i).value \
                                    for i in range(1, self.workingSheet.max_column, 1)\
                                    if self.workingSheet.cell(row=1, column=i).value != None]
        self.excel.sku.current(0) 
        self.excel.stock['values'] = [self.workingSheet.cell(row=1, column=i).value \
                                    for i in range(1, self.workingSheet.max_column, 1)\
                                    if self.workingSheet.cell(row=1, column=i).value != None]
        self.excel.stock.current(0)
        self.work.proccess.configure(state='active')
        self.work.status.set('Done.')  
        
    def load(self):
        self.work.update_status('Scrapping HTML...')
        self.scSoup = []
        #Get HTML into soup
        for url in self.src:
            if url.getFile() == '':
                continue
            scFile = open(url.getFile())
            self.scSoup.append(bs4.BeautifulSoup(scFile.read(), "html.parser"))
        
        #Opening Excel from excel input
        self.work.update_status('Getting excel sheet...')
        self.wb = openpyxl.load_workbook(self.excel.getFile(), data_only=True)
        self.excel.sheet['values'] = self.wb.get_sheet_names()
        self.excel.sheet.current(0)
        self.check(self)
        
    def compare(self):  
        self.columnSku = 0
        self.startColumnSku = self.workingSheet.max_column #start writing from last column + 1
        self.columnStock = 0
        self.startColumnStock = self.workingSheet.max_column+1 #start writing from last column + 1
        self.work.update_status('Getting column...')
        
        for name in self.excel.sku['values']:#getting sku ceolumn number
            self.columnSku += 1
            if name == self.excel.sku.get():
                break
        for name in self.excel.stock['values']:#getting stok ceolumn number
            self.columnStock += 1
            if name == self.excel.stock.get():
                break
            
        self.work.update_status('Please wait, proccessing...')
        self.workingSheet.cell(row=1, column=self.startColumnSku).value = 'SKU'
        self.workingSheet.cell(row=1, column=self.startColumnStock).value = 'Stock'
        
        for i in range(2, self.workingSheet.max_row+1, 1):#Getting each of cell in sku column
            y = 1
            for soup in self.scSoup:#check it with every soup available
                sku = soup.select('tr > td')
                stock = soup.select(' .stockOnSales')
                for index in range(3, len(sku), 11):
                    if self.workingSheet.cell(row=i, column=self.columnSku).value == sku[index].getText():
                        self.workingSheet.cell(row=i, column=self.startColumnSku).value = 'TRUE'
                        if (self.workingSheet.cell(row=i, column=self.columnStock).value) != (int(stock[y].getText())):
                            self.workingSheet.cell(row=i, column=self.startColumnStock).value = int(stock[y].getText())
                        else: self.workingSheet.cell(row=i, column=self.startColumnStock).value = '-'
                        break
                    y += 1
                break    
                        
        self.wb.save(self.excel.getFile())
        self.outputName.path.set(self.excel.getFile())
        self.outputName.button.configure(state='active')
        self.work.status.set('Complete.')
    
    '''                 
        #for soup in scSoup:
        price = self.scSoup[0].select(' .action.selling-price')
        name = self.scSoup[0].select(' .stockOnSales')
        sku = self.scSoup[0].select('tr > td') 
        for i in name:
            print i.getText()'''
         
main = window()

