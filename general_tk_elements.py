
from tkinter import *
from tkinter import ttk
import my_styles


class ProgressGeneral(ttk.Progressbar):
    def __init__(self, *args, **kwargs):

        self.rowno = kwargs['rowno']
        self.orient = kwargs['orient']
        self.length = kwargs['length']
        self.mode = kwargs['mode']
        self.style = kwargs['style']

        print(args)

        super().__init__(self)
