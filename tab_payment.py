from tkinter import *
from tkinter import ttk
from pathlib import Path
import time 

from payment import Payment

class TabPayment:
    def __init__(self, *args, **kwargs):

        self.tab=kwargs['tab']

        # Fetch user's selections from a file
        self.situation_path_from_file=''
        self.ginvoice_path_from_file=''
        self.mt_folder_path_from_file=''
        try:
            with open(Path('users_selections') / 'situation_file.txt', 'r') as f:
                self.situation_path_from_file = f.readline()
            with open(Path('users_selections') / 'ginvoice_file.txt', 'r') as f:
                self.ginvoice_path_from_file = f.readline()
            with open(Path('users_selections') / 'mt_folder.txt', 'r') as f:
                self.mt_folder_path_from_file = f.readline()
        except:
            pass
        # File and Address Selection selection
        self.payment_excel_value=StringVar()
        self.mt_folder_value=StringVar()
        self.ginvoice_excel_value = StringVar()

        # Destination selection
        # Payment excel selection
        self.payment_excel = ttk.Entry(self.tab, width=50, font='Helvetica 14', textvariable=self.payment_excel_value)
        self.payment_excel.insert(INSERT, Path(self.situation_path_from_file) )
        self.payment_excel.grid(row=1, column=1, columnspan=3, sticky=W, padx=(50,10), pady=20, ipadx=20, ipady=10 )
        self.browse_payment_excel = ttk.Button(
            self.tab, width=20 , text="Where is your excel situation file?", style='PaymentBrowse.TButton')
        self.browse_payment_excel.grid(row=1, column=5, columnspan=2, padx=5, ipadx=10, ipady=5 )

        # General Invoice excel selection
        self.ginvoice_excel = ttk.Entry(self.tab, width=50, font='Helvetica 14', textvariable=self.ginvoice_excel_value)
        self.ginvoice_excel.insert(INSERT, Path(self.ginvoice_path_from_file) )
        self.ginvoice_excel.grid(row=2, column=1, columnspan=3, sticky=W, padx=(50,10), pady=5, ipadx=20, ipady=5 )
        self.browse_ginvoice_excel = ttk.Button(
            self.tab, width=20 , text="General Invoice excel file?", style='PaymentBrowse.TButton')
        self.browse_ginvoice_excel.grid(row=2, column=5, columnspan=2, padx=5, ipadx=10, ipady=5 )

        # Result Folder selection
        self.mt_folder = ttk.Entry(self.tab, width=50, font='Helvetica 14', textvariable=self.mt_folder_value)
        self.mt_folder.insert(INSERT, Path(self.mt_folder_path_from_file))
        self.mt_folder.grid(row=3, column=1, columnspan=3, sticky=W, padx=(50,10), pady=10, ipadx=20, ipady=10 )
        self.browse_mt_folder = ttk.Button(
            self.tab, width=20 , text="Where is MT folder?", style='PaymentBrowse.TButton')
        self.browse_mt_folder.grid(row=3, column=5, columnspan=2, padx=5, ipadx=10, ipady=5 )

        # # Result Folder selection
        # self.result_folder = ttk.Entry(self.tab, width=35, font='Helvetica 14')
        # self.result_folder.insert(INSERT, Path.cwd())
        # self.result_folder.grid(row=2, column=1, sticky=W, padx=60, pady=10, ipadx=20, ipady=10 )
        # self.browse_result_folder = ttk.Button(
        #     self.tab, width=20 , text="Where do you want to save result?", style='PaymentBrowse.TButton')
        # self.browse_result_folder.grid(row=2, column=2, columnspan=2, padx=5, ipadx=10, ipady=5 )

        # Checkboxes
        self.facture=IntVar(value=1)
        self.routage=IntVar(value=1)
        self.decl=IntVar(value=1)
        self.tds=IntVar(value=1)
        self.coo = IntVar(value=1)

        # self.decl_page = IntVar(value=1)
        # self.tds_page = IntVar(value=0)

        self.facture_cbx = ttk.Checkbutton(self.tab, text='Facture', variable=self.facture,
                                           onvalue=1, offvalue=0, width=50, style='Allinone.TCheckbutton')
        self.routage_cbx = ttk.Checkbutton(
            self.tab, text='Routage', variable=self.routage, onvalue=1, offvalue=0, width=50, style='Allinone.TCheckbutton')
        self.decl_cbx = ttk.Checkbutton(self.tab, text='Declaration', variable=self.decl,
                                        onvalue=1, offvalue=0, width=12, style='Allinone.TCheckbutton', command=lambda:self.disable_enable_whole_and_page_rbtns(which='decl'))
        self.tds_cbx = ttk.Checkbutton(self.tab, text='TDS', variable=self.tds,
                                       onvalue=1, offvalue=0, width=12, style='Allinone.TCheckbutton', command=lambda:self.disable_enable_whole_and_page_rbtns(which='tds'))
        self.coo_cbx = ttk.Checkbutton(self.tab, text='Certificate of Origin', variable=self.coo,
                                       onvalue=1, offvalue=0, width=50, style='Allinone.TCheckbutton')

        # self.cbx_list = [self.facture_cbx, self.routage_cbx, self.decl_cbx, self.tds_cbx, self.coo_cbx]
        # for i, cbx in enumerate(self.cbx_list, start=5):
        #     if not cbx==self.decl_cbx or not cbx==self.tds_cbx:
        #         cbx.grid(row=i, column=1,sticky=W,
        #                 padx=(50,5), pady=5, ipadx=10, ipady=5)
        #     else:
        #         cbx.grid(row=i, column=1, columnspan=3, sticky=E,
        #                 padx=(50,5), pady=5, ipadx=10, ipady=5)
        #     # print(cbx['style'])
        
        self.facture_cbx.grid(row=5, column=1,  columnspan=3,sticky=W, padx=(50,5), pady=5, ipadx=10, ipady=5)
        self.routage_cbx.grid(row=6, column=1,  columnspan=3,sticky=W, padx=(50,5), pady=5, ipadx=10, ipady=5)
        self.coo_cbx.grid(row=9,    column=1,   columnspan=3,sticky=W, padx=(50,5), pady=5, ipadx=10, ipady=5)
        self.decl_cbx.grid(row=7,   column=1,               sticky=W, padx=(50,0), pady=1, ipadx=10, ipady=5)
        self.tds_cbx.grid(row=8,    column=1,               sticky=W, padx=(50,0), pady=1, ipadx=10, ipady=5)

        # Declaration RadioButtons
        self.decl_page_rbtn_value = IntVar(value=1)
        self.whole_decl_rbtn = ttk.Radiobutton(
                                self.tab, text='Whole', 
                                variable=self.decl_page_rbtn_value, value=0, width=10, style='Payment.TRadiobutton')
        self.decl_page_rbtn = ttk.Radiobutton(
                                self.tab, text='Specific page',
                                variable=self.decl_page_rbtn_value, value=1, width=15, style='Payment.TRadiobutton')
        self.whole_decl_rbtn.grid(row=int(self.decl_cbx.grid_info()['row']), column=int(self.decl_cbx.grid_info()['column'])+1,     sticky=W,
                               padx=1, pady=5, ipadx=1, ipady=5)
        self.decl_page_rbtn.grid(row=int(self.decl_cbx.grid_info()['row']), column=int(self.whole_decl_rbtn.grid_info()['column'])+1, sticky=W,
                               padx=1, pady=5, ipadx=1, ipady=5)
        
        # TDS RadioButtons
        self.tds_page_rbtn_value = IntVar(value=1)
        self.whole_tds_rbtn = ttk.Radiobutton(
                                self.tab, text='Whole', 
                                variable=self.tds_page_rbtn_value, value=1, width=10, style='Payment.TRadiobutton')
        self.tds_page_rbtn = ttk.Radiobutton(
                                self.tab, text='Specific page',
                                variable=self.tds_page_rbtn_value, value=0, width=15, style='Payment.TRadiobutton')
        self.whole_tds_rbtn.grid(row=int(self.tds_cbx.grid_info()['row']), column=int(self.tds_cbx.grid_info()['column'])+1,     sticky=W,
                               padx=1, pady=5, ipadx=1, ipady=5)
        self.tds_page_rbtn.grid(row=int(self.tds_cbx.grid_info()['row']), column=int(self.whole_tds_rbtn.grid_info()['column'])+1, sticky=W,
                               padx=1, pady=5, ipadx=1, ipady=5)

        # Prepare Button
        self.prepare_button=ttk.Button(self.tab, text='Prepare', style='PaymentPrepareButton.TButton')
        self.prepare_button.grid(row=int(self.coo_cbx.grid_info()['row'])+3, column=5, columnspan=2, sticky=E, padx=10, pady=10, ipadx=50, ipady=20)
        self.prepare_button.config(command=self.prepare_situation_documents)

    def disable_enable_whole_and_page_rbtns(self, **kwargs):
        self.which_one = kwargs['which']
        if self.which_one == 'decl':
            if self.decl.get():
                self.whole_decl_rbtn['state']='!disabled'
                self.decl_page_rbtn['state']='!disabled'
            else:
                self.whole_decl_rbtn['state']='disabled'
                self.decl_page_rbtn['state']='disabled'
        elif self.which_one == 'tds':
            if self.tds.get():
                self.whole_tds_rbtn['state']='!disabled'
                self.tds_page_rbtn['state']='!disabled'
            else:
                self.whole_tds_rbtn['state']='disabled'
                self.tds_page_rbtn['state']='disabled'

    def prepare_situation_documents(self):
        start=time.perf_counter()

        # Save user's selections to a file
        with open(Path('users_selections') / 'situation_file.txt', 'w') as f:
            f.write(self.payment_excel_value.get())
        with open(Path('users_selections') / 'ginvoice_file.txt', 'w') as f:
            f.write(self.ginvoice_excel.get())
        with open(Path('users_selections') / 'mt_folder.txt', 'w') as f:
            f.write(self.mt_folder_value.get())
            
        situation_file=Path(self.payment_excel_value.get())
        source_parent=Path(self.mt_folder_value.get())
        # general_invoice_file=source_parent / 'GENERAL INVOICE rev19.xlsm'
        general_invoice_file=Path(self.ginvoice_excel.get())
        destination=situation_file.parent / 'Payment Documentations'

        if not destination.exists():
            destination.mkdir(parents=True)

        situation_robot=Payment(source_parent=source_parent, destination=destination, situation_file=situation_file, general_invoice_file=general_invoice_file)
        # Assign user's wishes
        user_choice = lambda x : False if x==0 else True
        # self.facture=IntVar(value=1)
        # self.routage=IntVar(value=1)
        # self.decl=IntVar(value=1)
        # self.tds=IntVar(value=1)
        # self.coo = IntVar(value=1)
        situation_robot.want_facture = user_choice(self.facture.get())
        situation_robot.want_routage = user_choice(self.routage.get())
        situation_robot.want_decl = user_choice(self.decl.get())
        situation_robot.want_tds = user_choice(self.tds.get())
        situation_robot.want_coo = user_choice(self.coo.get())

        # Handle if user wishes page by page or not
        if situation_robot.want_decl:
            situation_robot.want_decl_page = user_choice(self.decl_page_rbtn_value.get())
            situation_robot.want_whole_decl = not user_choice(self.decl_page_rbtn_value.get())

        if situation_robot.want_tds:
            situation_robot.want_tds_page = not user_choice(self.tds_page_rbtn_value.get())
            situation_robot.want_whole_tds = user_choice(self.tds_page_rbtn_value.get())

        # Start processing
        situation_robot.assign_source_files_to_dataframe()
        situation_robot.assign_situation_file_to_dictionary()
        situation_robot.assign_general_invoice_to_dataframe()
        situation_robot.start_searching()
        situation_robot.copy_files()
        situation_robot.merge_all_in_one()
        situation_robot.write_not_founds_to_excel()
        situation_robot = None

        finish=time.perf_counter()
        print('Time spent for preparation of situation documents: ', round(finish-start,0), ' seconds')
        
