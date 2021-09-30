from tkinter import *
from tkinter import ttk
import pymsgbox
from pathlib import Path
from Finder import FileFinder
import time
from pathlib import Path
import percentage_holder


class TabAllinone:
    def __init__(self, *args, **kwargs):
        # Fetch user's previous selections
        self.users_mt_folder_path = ''
        self.users_ginv_path = ''
        try:
            with open(Path('users_selections') / 'mt_folder.txt', 'r') as f:
                self.users_mt_folder_path = f.readline()
            with open(Path('users_selections') / 'ginv.txt', 'r') as f:
                self.users_ginv_path = f.readline()
        except:
            pass

        self.tab=kwargs['tab']

        # MT Folder selection
        self.mt_folder_value=StringVar()
        self.mt_folder = ttk.Entry(self.tab, width=50, font='Helvetica 14', textvariable=self.mt_folder_value, style='AllinoneEntry.TEntry')
        self.mt_folder.insert(INSERT, Path(self.users_mt_folder_path))
        # self.mt_folder.grid(row=1, column=1, sticky=W, padx=60, pady=40, ipadx=20, ipady=10 )
        self.mt_folder.grid(row=1, column=1, sticky=W, padx=50, pady=15)

        self.browse_mt_folder = ttk.Button(
            self.tab, width=25 , text="Select MT folder, which contains files.", style='AllinoneBrowse.TButton')
        self.browse_mt_folder.grid(row=1, column=2, columnspan=2, padx=5, pady=5, ipadx=10, ipady=5 )


        # General Invoice excel file selection
        self.ginv_excel_value=StringVar()
        self.ginv_excel = ttk.Entry(self.tab, width=50, font='Helvetica 14', textvariable=self.ginv_excel_value, style='AllinoneEntry.TEntry')
        self.ginv_excel.insert(INSERT, Path(self.users_ginv_path))
        # self.ginv_excel.grid(row=2, column=1, sticky=W, ipadx=20, ipady=10 )
        self.ginv_excel.grid(row=2, column=1, sticky=W, padx=50, pady=15)

        self.browse_ginv_excel = ttk.Button(
            self.tab, width=25 , text="Select the GeneralInvoice.xlsx file.", style='AllinoneBrowse.TButton')
        self.browse_ginv_excel.grid(row=2, column=2, columnspan=2, padx=5, pady=5, ipadx=10, ipady=5 )


        # Checkboxes
        self.facture=StringVar(value='facture')
        self.routage=StringVar(value='routage')
        self.decl=StringVar(value='decl')
        self.tds=StringVar(value='tds')
        self.coo = StringVar(value='coo')

        self.facture_cbx = ttk.Checkbutton(self.tab, text='Facture', variable=self.facture,
                                           onvalue='facture', offvalue=0, width=50, style='Allinone.TCheckbutton')
        self.routage_cbx = ttk.Checkbutton(self.tab, text='Routage', variable=self.routage, onvalue='routage', offvalue=0, width=50, style='Allinone.TCheckbutton')
        self.decl_cbx = ttk.Checkbutton(self.tab, text='Declaration', variable=self.decl, onvalue='decl', offvalue=0, width=50, style='Allinone.TCheckbutton')
        self.tds_cbx = ttk.Checkbutton(self.tab, text='TDS', variable=self.tds, onvalue='tds', offvalue=0, width=50, style='Allinone.TCheckbutton')
        self.coo_cbx = ttk.Checkbutton(self.tab, text='Certificate of Origin', variable=self.coo, onvalue='coo', offvalue=0, width=50, style='Allinone.TCheckbutton')

        self.cbx_list = [self.facture_cbx, self.routage_cbx,
                         self.decl_cbx, self.tds_cbx, self.coo_cbx]
        for i, cbx in enumerate(self.cbx_list, start=3):
            cbx.grid(row=i, column=1, columnspan=2, sticky=E,
                     padx=50, pady=5, ipadx=20, ipady=5)
            print(cbx['style'])
        
        # Progres Bar
        
        # self.progress=general_tk_elements.ProgressGeneral(self.tab,rowno=int(self.coo_cbx.grid_info()['row'])+1, orient = HORIZONTAL, length = 400, mode = 'determinate', style='grey.Horizontal.TProgressbar')
        # self.progress.grid(row=int(self.coo_cbx.grid_info()['row'])+1, column=1, sticky=W, padx=25, pady=5, ipadx=20, ipady=20) 
        self.progress = ttk.Progressbar(self.tab, orient = HORIZONTAL, length = 400, mode = 'determinate', style='grey.Horizontal.TProgressbar')
        self.progress.grid(row=int(self.coo_cbx.grid_info()['row'])+1, column=1, sticky=W, padx=25, pady=5, ipadx=20, ipady=20) 

        # Copy Button
        self.copy_button=ttk.Button(self.tab, text='Copy to AllInOne', style='AllinoneCopyButton.TButton')
        self.copy_button.grid(row=int(self.coo_cbx.grid_info()['row'])+1, column=2, columnspan=2, sticky=E, padx=10, pady=10, ipadx=50, ipady=20)
        self.copy_button.config(command=self.copy_files)

    def copy_files(self):
        # Save user's selections to a file
        with open(Path('users_selections') / 'mt_folder.txt', 'w') as f:
            f.write(self.mt_folder_value.get())
        with open(Path('users_selections') / 'ginv.txt', 'w') as f:
            f.write(self.ginv_excel_value.get())

        self.start = time.perf_counter()
        self.finish = time.perf_counter()
        self.users_selection=[self.facture.get(),self.routage.get(),self.decl.get(), self.tds.get(), self.coo.get()]
        self.files_type_to_copy=[]
        self.total_for_progress={}
        # self.files_type_to_copy=['facture', 'routage', 'decl', 'tds', 'coo']

        
        for slctn in self.users_selection:
            if slctn!='0':
                self.files_type_to_copy.append(slctn)

        # Debug
        print(self.files_type_to_copy)

        


        # Progress Bar
        # self.progress=Progress(self.tab)

        self.stime={}
        self.ftime={}

        # Loop through user's seledtions one by one
        # and loop through inside each file type
        percentage_holder.resetpercentage()
        self.tab.update_idletasks()
        for file_type in self.files_type_to_copy:
            print('File type: {}'.format(file_type))
            self.stime.update({file_type: time.perf_counter()})

            # Check if user selected right folder and right excel file
            self.file_finder = FileFinder(file_type, mt_folder=Path(self.mt_folder_value.get()), wb=Path(self.ginv_excel_value.get()))
            if self.file_finder.is_wb_error:
                pymsgbox.alert('You did not choose excel file')
                self.finish=time.perf_counter()
                break

            if not self.file_finder.all_folders_in_place:
                if self.file_finder.partially_found:
                    pymsgbox.alert(text=self.file_finder.partially_found_alert_message)
                    if not file_type in self.file_finder.found_paths:
                        continue
                else:
                    self.alert_text=self.file_finder.none_found_alert_message
                    self.alert_text+='\n'
                    self.alert_text+='Program will quit, please choose proper folder and try again later.'
                    pymsgbox.alert(text=self.alert_text)
                    break


            # self.total_facture=self.file_finder.total_factures
            print('Searching for {}'.format(file_type))
            self.file_finder.find_files()
            print('Copying {}'.format(file_type))
            # self.file_finder.copy_files()
            self.file_finder.copy_files(mainmenu=self.tab, progressbar=self.progress, total_selection=len(self.files_type_to_copy)-len(self.file_finder.not_found_paths))

            # Update general invoice sheet with 'NOT FOUND' info
            self.file_finder.update_ginv_sheet()


        # Check users selection, if there is no requested folders inside selection folder, then warn
        if self.file_finder.partially_found:
            self.alert_text='These files are not processed, because there is no folder related to them'
            self.alert_text+='\n'
            self.alert_text+=str(list(map(lambda x: x, self.file_finder.not_found_paths)))

            # Check if user requested the file type that was not found
            # for f in self.file_finder.not_found_paths:
            #     if f not in self.files_type_to_copy:
            #         self.file_finder.not_found_paths.remove(f)
            
            self.alert_text+='\n'
            self.alert_text+='\n'
            self.alert_text+='But luckily this files has been processed successfully: '
            self.alert_text+=str(list(map(lambda x: x, self.file_finder.found_paths)))

            pymsgbox.alert(text=self.alert_text)

        self.ftime.update({file_type: time.perf_counter()})
        self.finish=time.perf_counter()
        
        print('Total time spent: {}'.format(round(self.finish-self.start,2)))
        for xtime in self.stime:
            try:
                print('Total time spent for {}: {}'.format(xtime, round(self.ftime[xtime]-self.stime[xtime],2)))
            except:
                pass
