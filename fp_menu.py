from tkinter import *
from tkinter import ttk, filedialog, Text
from PIL import ImageTk, Image
from pathlib import Path
from image_extracter_with_openpyxl_class import FP_Image_Extractor
from fiche_produit_extracter_class import FicheProduit_to_Excel
class MainMenu:
    def __init__(self, *args, **kwargs):
        self.root=Tk()
        self.root.title("Extract Images and all info of Fiche Produit")
        self.root.width=100
        self.root.height=100

        self.notebook=ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0)
        self.home_frame=ttk.Frame(self.notebook)
        self.settings_frame=ttk.Frame(self.notebook)
        self.notebook.add(self.home_frame, text='Home')
        self.notebook.add(self.settings_frame, text='Settings')

        # Elements of Home tab
        self.header_frame=Frame(self.home_frame, width=900, height=590)

        # BYTK logo
        self.logo_address=Path.cwd() / 'app_files/images/bytk_logo.png'
        print(self.logo_address)
        self.bytk_logo_pil=Image.open(self.logo_address)
        self.bytk_logo_pil=self.bytk_logo_pil.resize((180, 80), Image.ANTIALIAS)
        self.bytk_logo=ImageTk.PhotoImage(self.bytk_logo_pil)

        self.canvas_logo=Canvas(self.header_frame, width=200, height=100) #, bg='red')
        #self.canvas_logo.image=self.bytk_logo
        self.canvas_logo.create_image(10,10,anchor=NW, image=self.bytk_logo)
        self.canvas_logo.pack(side=LEFT)
        # self.canvas_logo.grid(row=0, column=0, columnspan=4, sticky=W)

        #Image fp_to_xl
        self.fp_to_xl_address=Path.cwd() / 'app_files/images/fp_to_xl.png'
        print(self.fp_to_xl_address)
        self.fp_to_xl_pil=Image.open(self.fp_to_xl_address)
        self.fp_to_xl_pil=self.fp_to_xl_pil.resize((360, 100))
        self.fp_to_xl=ImageTk.PhotoImage(self.fp_to_xl_pil)

        self.canvas_fp_to_xl=Canvas(self.home_frame, width=200, height=55)
        self.canvas_fp_to_xl.create_image(15,15,anchor=NW, image=self.fp_to_xl)
        self.canvas_fp_to_xl.grid(row=8, column=0, rowspan=3, columnspan=4, stick='nwes', padx=5, pady=5)

        # Styles
        self.font_usual=ttk.Style()
        self.font_usual.configure('.',font=('Helvetica', 12))
        # Banner Label
        self.label_banner=ttk.Label(self.header_frame, wrap=400, text='Fiche Produit Transfer', font=('Helvetica', 30))
        self.label_banner.pack(side=LEFT)

        self.label_source=ttk.Label(self.home_frame, text="Source Folder")
        self.label_dest=ttk.Label(self.home_frame, text="Destination Folder")

        self.entry_source=ttk.Entry(self.home_frame, width=70)
        self.entry_source.insert(INSERT, Path.cwd() / "cra_list_test")
        self.entry_source['state']='disabled'

        self.entry_dest=ttk.Entry(self.home_frame, width=70)
        self.entry_dest.insert(INSERT, Path.cwd() / "result")
        self.entry_dest['state']='disabled'

        self.browse_source=ttk.Button(self.home_frame,width=5, text="...", command=lambda: self.select_folder("source"))
        self.browse_dest=ttk.Button(self.home_frame,width=5, text="...", command=lambda: self.select_folder("destination"))

        self.extract_button=ttk.Button(self.home_frame, width=20, text="Extract", command=self.extract_click)

        # Putting everything in place
        # Elements of Home tab
        self.header_frame.grid(row=0, column=0, rowspan=3, columnspan=6, stick='nwes')

        self.label_source.grid(row=5, column=0, stick='w', padx=5, pady=5)
        self.entry_source.grid(row=5, column=1, columnspan=4, padx=5, pady=5)
        self.browse_source.grid(row=5, column=5, padx=10)

        self.label_dest.grid(row=6, column=0, stick='w', padx=5, pady=5)
        self.entry_dest.grid(row=6, column=1, columnspan=4, padx=5, pady=5)
        self.browse_dest.grid(row=6, column=5, padx=10)

        self.extract_button.grid(row=8, column=4, columnspan=2, sticky=E, padx=10, pady=40, ipadx=10, ipady=10)

        # Elements of Settings tab
        self.frame_is_fp=ttk.Frame(self.settings_frame)
        self.frame_is_prd_name=ttk.Frame(self.settings_frame)
        self.label_is_fp=ttk.Label(self.frame_is_fp, text='What must your Fiche Produit must contain at least? Please specify!', wraplength=260, justify='center')
        self.label_is_prd_name=ttk.Label(self.frame_is_prd_name, text='What is the label name of Product Name in you Fiche Produit table? Please specify!', wraplength=260, justify='center')
        self.entry_fp1=ttk.Entry(self.settings_frame, width=40)
        self.entry_fp2=ttk.Entry(self.settings_frame, width=40)
        self.entry_fp3=ttk.Entry(self.settings_frame, width=40)
        self.entry_fp4=ttk.Entry(self.settings_frame, width=40)
        self.entry_prd_name1=ttk.Entry(self.settings_frame, width=40)
        self.entry_prd_name2=ttk.Entry(self.settings_frame, width=40)
        self.entry_prd_name3=ttk.Entry(self.settings_frame, width=40)

        self.entry_fp1.insert(INSERT, "FICHE PRODUIT")
        self.entry_fp2.insert(INSERT, "Nom du produit")
        self.entry_fp3.insert(INSERT, "Name of Product")
        self.entry_fp4.insert(INSERT, "")
        self.entry_prd_name1.insert(INSERT, "Nom du produit")
        self.entry_prd_name2.insert(INSERT, "Name of Product")
        self.entry_prd_name3.insert(INSERT, "")

        self.frame_is_fp.grid(row=0, column=0)
        self.label_is_fp.grid(row=0, column=0, padx=20, pady=20)
        self.entry_fp1.grid(row=3, column=0, pady=5)
        self.entry_fp2.grid(row=4, column=0, pady=5)
        self.entry_fp3.grid(row=5, column=0, pady=5)
        self.entry_fp4.grid(row=6, column=0, pady=5)

        self.frame_is_prd_name.grid(row=0, column=1)
        self.label_is_prd_name.grid(row=0, column=0, padx=20, pady=20)
        self.entry_prd_name1.grid(row=3, column=1, pady=5)
        self.entry_prd_name2.grid(row=4, column=1, pady=5)
        self.entry_prd_name3.grid(row=5, column=1, pady=5)


        self.root.mainloop()

    def select_folder(self, *args):
        folder_path=filedialog.askdirectory()

        if "source" in args:
            self.entry_source['state']='!disabled'
            self.entry_source.delete(0, END)
            self.entry_source.insert(0, folder_path)
            self.entry_source['state']='disabled'
        elif "destination" in args:
            self.entry_dest['state']='!disabled'
            self.entry_dest.delete(0, END)
            self.entry_dest.insert(0, folder_path)
            self.entry_dest['state']='disabled'

    def extract_click(self, *args):
        self.source_path=self.entry_source.get()
        self.dest_path=self.entry_dest.get()

        fp_to_excel_table=FicheProduit_to_Excel(xpath=self.source_path, result_path=self.dest_path)
        fp_to_excel_table.loop_trough_excels()

if __name__=='__main__':
    global test
    test=MainMenu()

