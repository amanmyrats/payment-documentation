from tkinter import *
from tkinter import ttk, filedialog, Text
from PIL import ImageTk, Image
from pathlib import Path

import my_styles
import tab_allinone, tab_payment, tab_xltopdf, tab_xl_to_table


class MainMenu(my_styles.MyStyle):
    def __init__(self, *args, **kwargs):
        
        self.root=Tk()
        self.root.title("Payment documentations")
        self.root.width=100
        self.root.height=100

        my_styles.MyStyle.__init__(self)

        self.notebook=ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0)

        self.payment_frame=ttk.Frame(self.notebook)
        self.allinone_frame=ttk.Frame(self.notebook)
        self.converttopdf_frame=ttk.Frame(self.notebook)
        self.invoicetotable_frame=ttk.Frame(self.notebook)

        self.notebook.add(self.payment_frame, text='Payment Documentation')
        self.notebook.add(self.allinone_frame, text='Organizing Documents (AllInOne)')
        self.notebook.add(self.converttopdf_frame, text='Convert Excel Invoice to PDF')
        self.notebook.add(self.invoicetotable_frame, text='Invoice to Material Table')

        # TAB - Payment
        self.tab_payment=tab_payment.TabPayment(tab=self.payment_frame)
        self.tab_payment.browse_payment_excel.config(command=lambda:self.select_file('payment_excel_source'))
        self.tab_payment.browse_ginvoice_excel.config(command=lambda:self.select_file('ginvoice_excel_source'))
        self.tab_payment.browse_mt_folder.config(command=lambda:self.select_folder('mt_folder_dest'))

        # TAB - All in One
        self.tab_allinone=tab_allinone.TabAllinone(tab=self.allinone_frame)
        self.tab_allinone.browse_mt_folder.config(command=lambda:self.select_folder('mt_folder'))
        self.tab_allinone.browse_ginv_excel.config(command=lambda:self.select_file('ginv_excel'))

        # TAB - Excel to PDF
        self.tab_xltopdf=tab_xltopdf.TabXltopdf(tab=self.converttopdf_frame)
        self.tab_xltopdf.browse_excel_invoices.config(command=lambda:self.select_folder('excel_invoices_source'))
        self.tab_xltopdf.browse_result_folder.config(command=lambda:self.select_folder('excel_pdf_dest'))

        # TAB - Facture to Material Table
        self.tab_xl_to_table=tab_xl_to_table.TabXltoTable(tab=self.invoicetotable_frame)
        self.tab_xl_to_table.browse_all_factures_path.config(command=lambda:self.select_folder('all_factures_path'))
        self.tab_xl_to_table.browse_material_table_path.config(command=lambda:self.select_file('material_table'))


        self.root.mainloop()


    def select_folder(self, *args):
        self.folder_path=filedialog.askdirectory()
        
        # Payment tab
        if "payment_excel_source" in args:
            self.tab_payment.payment_excel['state']='!disabled'
            self.tab_payment.payment_excel.delete(0, END)
            self.tab_payment.payment_excel.insert(0, self.folder_path)
            self.tab_payment.payment_excel['state']='disabled'
        elif "mt_folder_dest" in args:
            self.tab_payment.mt_folder['state']='!disabled'
            self.tab_payment.mt_folder.delete(0, END)
            self.tab_payment.mt_folder.insert(0, self.folder_path)
            self.tab_payment.mt_folder['state']='disabled'
        
        # Tab all in one
        elif 'mt_folder' in args:
            self.tab_allinone.mt_folder['state']='!disabled'
            self.tab_allinone.mt_folder.delete(0, END)
            self.tab_allinone.mt_folder.insert(0, self.folder_path)
            self.tab_allinone.mt_folder['state']='disabled'
        
        # Excel to PDF Tab
        elif 'excel_invoices_source' in args:
            self.tab_xltopdf.excel_invoices['state']='!disabled'
            self.tab_xltopdf.excel_invoices.delete(0, END)
            self.tab_xltopdf.excel_invoices.insert(0, self.folder_path)
            self.tab_xltopdf.excel_invoices['state']='disabled'
            
            # Suggested destination path for converted PDFs
            self.tab_xltopdf.result_folder['state']='!disabled'
            self.tab_xltopdf.result_folder.delete(0, END)
            self.tab_xltopdf.result_folder.insert(0, str(self.folder_path)+ '/1-FACTURE PDF')
            self.tab_xltopdf.result_folder['state']='disabled'

            # Test
            # print(self.tab_xltopdf.rbtn_value.get())

        elif 'excel_pdf_dest' in args:
            self.tab_xltopdf.result_folder['state']='!disabled'
            self.tab_xltopdf.result_folder.delete(0, END)
            self.tab_xltopdf.result_folder.insert(0, self.folder_path)
            self.tab_xltopdf.result_folder['state']='disabled'
        
        # Facture to Material Table Tab
        elif 'all_factures_path' in args:
            self.tab_xl_to_table.all_factures_path['state']='!disabled'
            self.tab_xl_to_table.all_factures_path.delete(0, END)
            self.tab_xl_to_table.all_factures_path.insert(0, self.folder_path)
            self.tab_xl_to_table.all_factures_path['state']='disabled'


    def select_file(self, *args):
        self.file_path=filedialog.askopenfilename()

        # Payment tab
        if "payment_excel_source" in args:
            self.tab_payment.payment_excel['state']='!disabled'
            self.tab_payment.payment_excel.delete(0, END)
            self.tab_payment.payment_excel.insert(0, self.file_path)
            self.tab_payment.payment_excel['state']='disabled'
        elif 'ginvoice_excel_source' in args:
            self.tab_payment.ginvoice_excel['state']='!disabled'
            self.tab_payment.ginvoice_excel.delete(0, END)
            self.tab_payment.ginvoice_excel.insert(0, self.file_path)
            self.tab_payment.ginvoice_excel['state']='disabled'

        # All in one tab
        if "ginv_excel" in args:
            self.tab_allinone.ginv_excel['state']='!disabled'
            self.tab_allinone.ginv_excel.delete(0, END)
            self.tab_allinone.ginv_excel.insert(0, self.file_path)
            self.tab_allinone.ginv_excel['state']='disabled'

        # Facture to Material Table Tab
        if "material_table" in args:
            self.tab_xl_to_table.material_table_path['state']='!disabled'
            self.tab_xl_to_table.material_table_path.delete(0, END)
            self.tab_xl_to_table.material_table_path.insert(0, self.file_path)
            self.tab_xl_to_table.material_table_path['state']='disabled'
        
  

if __name__=='__main__':
    test=MainMenu()
