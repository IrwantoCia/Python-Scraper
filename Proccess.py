'''
Created on Jan 31, 2016

@author: Irwanto
'''
import Tkinter as tk
from ttk import LabelFrame
import time

class calc():
    def __init__(self, root, gridRow, gridColumn, prc_cmd, load_cmd):
        self.root = root
        self.proccess = tk.Button(self.root, text='Proccess', command=prc_cmd)
        self.proccess.grid(row=gridRow, column=gridColumn, padx=5, pady=5,\
                           sticky='ws')
        self.load = tk.Button(self.root, text='Load', command=load_cmd)
        self.load.grid(row=gridRow, column=gridColumn, padx=75, pady=5,\
                           sticky='ws')
        
        status = LabelFrame(self.root, text=' Status ', labelanchor='ne')
        status.grid(row=gridRow, column=gridColumn+1, padx=5, pady=5, sticky='en')
        self.status = tk.StringVar()
        self.text = tk.Label(status, textvariable=self.status)
        self.text.grid(row=0, column=0, padx=5, sticky='w')                          
        self.status.set('Waiting')
        
    def update_status(self, text):
        self.status.set(text)
        self.root.update_idletasks()
        time.sleep(1)
        
