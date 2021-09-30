from tkinter import *
from tkinter import ttk, filedialog, Text
from ttkthemes import ThemedStyle
# import main_menu

class MyStyle:
    def __init__(self):
        # self.all_style=ttk.Style()
        # self.all_style.configure('.',font=('Helvetica', 12))
        # self.all_style.theme_use('alt')
        # self.all_style.configure('Allinone.TCheckbutton',  font='helvetica 14', background='gray', padding=10)

        self.style = ThemedStyle()
        self.style.theme_use('adapta')  # white style adapta
        self.style.configure('.', background='lavender')
        self.style.configure('TNotebook.Tab', font=('URW Gothic L','10','bold'))

        # Payment
        self.style.configure('PaymentBrowse.TButton',  font='helvetica 12', padding=3, wraplength=150, justify=CENTER)
        self.style.configure('PaymentPrepareButton.TButton',  font='helvetica 14', padding=5)      
        self.style.configure('Payment.TRadiobutton', font='times 14 bold italic', padding=5, wraplength=350)

        # All in one
        self.style.configure('AllinoneEntry.TEntry',  font='helvetica 12', padding=10)
        self.style.configure('AllinoneBrowse.TButton',  font='helvetica 12', wraplength=180, justify=CENTER, padding=5)
        self.style.configure('Allinone.TCheckbutton',  font='helvetica 14', padding=10)        
        self.style.configure('AllinoneCopyButton.TButton',  font='helvetica 14', padding=5)        
        
        # XL to PDF
        self.style.configure('XltopdfBrowse.TButton',  font='helvetica 12', padding=5, wraplength=150, justify=CENTER)
        self.style.configure('Xltopdf.TRadiobutton', font='helvetica 14', padding=5, wraplength=350)
        self.style.configure('XltopdfConvertButton.TButton',  font='helvetica 14', padding=5)  

        # Invoice format to material table format
        self.style.configure('Totable.TCheckbutton',  font='helvetica 14', padding=10)
        self.style.configure('TotableBrowse.TButton',  font='helvetica 12', padding=5, wraplength=150, justify=CENTER)
        self.style.configure('TotableMainButton.TButton',  font='helvetica 14', padding=5)  

        # Progress bar
        self.style.configure("grey.Horizontal.TProgressbar", background='grey', foreground='blue')  


# if __name__=='__main__':
#     testx=main_menu.MainMenu()