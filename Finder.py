from pathlib import Path
from openpyxl import Workbook, load_workbook
import os
import re
import threading
import time
import pandas as pd
import shutil
from multiprocessing import Pool, Process, Manager
import concurrent.futures
import psutil
from tkinter import *
from tkinter import ttk
from progressbar import * 
from FilesClass import AllFiles

from helper import *
from last_modife import LastModifeFinder
import percentage_holder

# ginv = {}
all_files = {}


class FileFinder(AllFiles):
    def __init__(self, file_type, **kwargs):
        self.mt_folder=kwargs['mt_folder']
        self.wb=kwargs['wb']
        self.general_invoice_full_name = kwargs['wb']
        self.file_type=file_type

        super().__init__(self, mt_folder=self.mt_folder, wb=self.wb)
        # super().__init__(self)

    def check_user_selections(self):
        return self.is_wb_error

    def find_files(self):
        print(f'Finding process of {self.file_type} started.')
        self.list_of_files_to_search=[]

        # Convert paths to list
        self.list_path=[str(x) for x in self.overall_path['source'][self.file_type].rglob('*.pdf')]

        # Convert lists to Dataframe.
        self.df_of_files_to_work=pd.DataFrame(self.list_path, columns=[self.file_type])

        # Testing purpose
        self.df_of_files_to_work['pure_no']='Empty'
        self.df_of_files_to_work['modife']='Empty'

        # Pooling
        self.total_physical_core=psutil.cpu_count(logical=False)
        # print(f'Total physical cores: {self.total_physical_core-1}')
            # p=Pool(self.total_physical_core-1)
            # self.result_of_pool_as_list=p.map(self.pool_thread, self.iter_list_for_pooling)
            # p.close()
            # p.join()
            
            # Above pooling gave me a list
            # Here I convert that result list to a dictionary (my standard dictionary format)
            # self.ginv_of_pooling={}
            # for dcts in self.result_of_pool_as_list:
            #     self.ginv_of_pooling.update(dcts)

        # Normal threading
        self.ginv_of_pooling={}
        for root in self.ginv:
            dct = {}
            dct = self.pool_thread(root)
            self.ginv_of_pooling.update(dct)
            # print(root)


        # self.ginv_of_pooling={}
            # for root in self.iter_list_for_pooling:
            #     dct = {}
            #     dct = self.pool_thread(root)
            #     self.ginv_of_pooling.update(dct)


    def pool_thread(self, root):
        # print('searching for {}'.format(root))
        # I am not sure but it looks like 1st if is not necessary
        # Then loop through each file, according to requested file type
        # print('inside pool thread')
        if self.ginv[root][self.file_type]!='':
            # print('inside pool thread if')
            for _ in self.ginv[root][self.file_type]:
                self.ginv[root][self.file_type]['file']='Not Found'
                self.no_to_find=self.ginv[root][self.file_type]['no']

                # Handly empties and '0's
                if self.no_to_find =='0' or self.no_to_find=='' or len(self.no_to_find)<4:
                    print('Empty {} : {}'.format(self.file_type, root))
                    # print('before not found dict assignment')
                    self.not_found_files_dict[self.file_type].append(root)
                    # print('after not found dict assignment')
                    
                    try:
                        self.ginv[root][self.file_type]['file'] = list()
                        return {root : self.ginv[root]}
                    except:
                        return {root : self.ginv[root]}

                # Switch patterns according to file types
                if self.file_type=='facture' or self.file_type=='coo':
                    self.pattern1=r'.*?(\D|\b)' + str(self.no_to_find) + r'\D.*pdf$'
                    self.pattern2=r'.*?[^-]' + str(self.no_to_find) + r'\D.*pdf$'
                else:
                    self.pattern1=r'.*?' + str(self.no_to_find) + r'\D.*pdf$'

                # Here is main matching process is done by using str.match method
                self.match_result=self.df_of_files_to_work[self.df_of_files_to_work[self.file_type].str.match(self.pattern1,case=False, na=False)]
                if self.file_type=='facture' or self.file_type=='coo':
                    self.match_result=self.match_result[self.match_result[self.file_type].str.match(self.pattern2,case=False, na=False)]
                
                
                # If requested file type is facture then do following
                if self.match_result.shape[0]>1:
                    
                    if self.file_type=='facture':
                        self.ysyrgayjy = LastModifeFinder(parent_path = self.overall_path['source']['facture'])
                        self.ysyrgayjy.find_last_modife(facture_no = root, file_name_list = self.match_result[self.file_type].tolist())

                        if not str(self.ysyrgayjy.last_modife).lower()=='none':
                            self.ginv[root][self.file_type]['file'] = list()
                            self.ginv[root][self.file_type]['file'].append(self.ysyrgayjy.last_modife)
                        else:
                            print('Last modife returned none in {}'.format(root))
                            print('Normally one of these factures should be selected: (I selected 1st one)')
                            print(self.match_result[self.file_type].tolist())
                            self.ginv[root][self.file_type]['file'] = list()
                            self.ginv[root][self.file_type]['file'].append(self.match_result[self.file_type].tolist()[0])

                    # If requested file type is not a facture then do following
                    # If there is more than one match, then just select first one
                    elif self.file_type=='coo':
                        if self.match_result.shape[0]>10:
                            print('Shape is bigger than 10.')
                            self.ginv[root][self.file_type]['file'] = list()
                            return {root : self.ginv[root]}
                        try:
                            self.ginv[root][self.file_type]['file'] = self.match_result[self.file_type].tolist()
                        except:
                            pass
                    else:
                        self.df_of_files_to_work.loc[self.match_result.index.tolist(),'pure_no']=root
                        self.df_of_files_to_work.loc[self.match_result.index.tolist(),'modife']=0
                        self.all_current_root = self.df_of_files_to_work[(self.df_of_files_to_work['pure_no'] == root)]
                        try:
                            self.ginv[root][self.file_type]['file'] = list()
                            self.ginv[root][self.file_type]['file'].append(self.all_current_root.iat[0, 0])
                        except:
                            pass
                        
                # If matched file number is equal to one then just assign that file to --- self.ginv[root][self.file_type]['file'] ---
                elif self.match_result.shape[0]==1:
                    self.df_of_files_to_work.loc[self.match_result.index.tolist(),'pure_no']=root
                    self.df_of_files_to_work.loc[self.match_result.index.tolist(),'modife']=0
                    self.all_current_root =self.df_of_files_to_work[(self.df_of_files_to_work['pure_no']==root)]
                    try:
                        self.ginv[root][self.file_type]['file'] = list()
                        self.ginv[root][self.file_type]['file'].append(self.all_current_root.iat[0,0])
                    except:
                        pass
                elif self.match_result.shape[0]==0:
                    # Search partially
                    if not self.file_type=='facture' or not self.file_type=='coo':
                        self.do_partial_search(root_for_partial = root)

                    if self.match_result.shape[0]==0:
                        print('No match found for {} : {}'.format(self.file_type, root))
                        self.not_found_files_dict[self.file_type].append(root)
                        self.ginv[root][self.file_type]['file'] = list()


                return {root : self.ginv[root]}
        
    def do_partial_search(self, *args, **kwargs):
        # self.df_for_partial_search = self.df
        self.root_for_partial = kwargs['root_for_partial']
        if self.file_type=='routage':
            self.partial_routage_list = self.routage_splitter(self.no_to_find)
            if len(self.partial_routage_list)==0:
                return 'none'
            
            if len(self.partial_routage_list)==1 and self.partial_routage_list[0]==self.no_to_find:
                return 'none'

            self.ginv[self.root_for_partial][self.file_type]['file'] = list()
            for rtg in self.partial_routage_list:
                # Here is main matching process is done by using str.match method
                self.match_result=self.df_of_files_to_work[self.df_of_files_to_work[self.file_type].str.match(r'.*?' + str(rtg) + r'\D.*pdf$',case=False, na=False)]
                if self.match_result.shape[0]>0:
                    
                    # self.df_of_files_to_work.loc[self.match_result.index.tolist(),'pure_no']=self.root_for_partial
                    # self.df_of_files_to_work.loc[self.match_result.index.tolist(),'modife']=0
                    # self.all_current_root =self.df_of_files_to_work[(self.df_of_files_to_work['pure_no']==self.root_for_partial)]
                    try:
                        self.ginv[self.root_for_partial][self.file_type]['file'].append(self.match_result.iat[0,0])
                        # print('Partially found {}'.format(self.all_current_root.iat[0,0]))
                    except:
                        print('Error inside routage partial search.')
                        pass
                else:
                    self.not_found_files_dict[self.file_type].append(self.root_for_partial)

        elif self.file_type == 'decl':
            self.partial_decl_list = self.decl_splitter(self.no_to_find)
            # print('\nThis is declaration list to find:')
            # print(self.partial_decl_list, '\n')
            if len(self.partial_decl_list)==0:
                return 'none'
            
            if len(self.partial_decl_list)==1 and self.partial_decl_list[0]==self.no_to_find:
                return 'none'

            self.ginv[self.root_for_partial][self.file_type]['file'] = list()
            for dcl in self.partial_decl_list:
                # Here is main matching process is done by using str.match method
                self.match_result=self.df_of_files_to_work[self.df_of_files_to_work[self.file_type].str.match(r'.*?' + str(dcl) + r'\D.*pdf$',case=False, na=False)]
                if self.match_result.shape[0]>0:
                    
                    # self.df_of_files_to_work.loc[self.match_result.index.tolist(),'pure_no']=self.root_for_partial
                    # self.df_of_files_to_work.loc[self.match_result.index.tolist(),'modife']=0
                    # self.all_current_root =self.df_of_files_to_work[(self.df_of_files_to_work['pure_no']==self.root_for_partial)]
                    try:
                        self.ginv[self.root_for_partial][self.file_type]['file'].append(self.match_result.iat[0,0])
                        # print('Partially found {}'.format(self.all_current_root.iat[0,0]))
                    except:
                        print('Error inside declaration partial search.')
                        pass
                else:
                    self.not_found_files_dict[self.file_type].append(self.root_for_partial)

        elif self.file_type == 'tds':
            self.partial_tds_list = self.tds_splitter(self.no_to_find)
            print('\nThis is tds list to find:')
            print(self.partial_tds_list, '\n')
            if len(self.partial_tds_list)==0:
                return 'none'
            
            if len(self.partial_tds_list)==1 and self.partial_tds_list[0]==self.no_to_find:
                return 'none'

            self.ginv[self.root_for_partial][self.file_type]['file'] = list()
            for tds in self.partial_tds_list:
                # Here is main matching process is done by using str.match method
                self.match_result=self.df_of_files_to_work[self.df_of_files_to_work[self.file_type].str.match(r'.*?' + str(tds) + r'\D.*pdf$',case=False, na=False)]
                if self.match_result.shape[0]>0:
                    
                    # self.df_of_files_to_work.loc[self.match_result.index.tolist(),'pure_no']=self.root_for_partial
                    # self.df_of_files_to_work.loc[self.match_result.index.tolist(),'modife']=0
                    # self.all_current_root =self.df_of_files_to_work[(self.df_of_files_to_work['pure_no']==self.root_for_partial)]
                    try:
                        self.ginv[self.root_for_partial][self.file_type]['file'].append(self.match_result.iat[0,0])
                        # print('Partially found {}'.format(self.all_current_root.iat[0,0]))
                    except:
                        print('Error inside tds partial search.')
                        pass
                else:
                    self.not_found_files_dict[self.file_type].append(self.root_for_partial)

    def copy_files(self, **kwargs):
        print(f'Copying {self.file_type} started.')
        self.iter_list_for_copying=[]
        # manager=Manager()
        self.processes=[]
        
        # Debug
        # print('ginv of pooling')
        # print(self.ginv_of_pooling)

        self.pooling_items_total_count=len(self.ginv_of_pooling)
        self.old_percentage = 0
        for i, root in enumerate(self.ginv_of_pooling):

            self.percentage=(1/self.pooling_items_total_count)*100/kwargs['total_selection']
            percentage_holder.updatepercentage(newpercentage=self.percentage)
            if int(percentage_holder.getpercentage()) % 2 == 0 and not self.old_percentage == int(percentage_holder.getpercentage()):
                self.old_percentage = int(percentage_holder.getpercentage())
                kwargs['progressbar']['value']=percentage_holder.getpercentage()
                # print('Total factures: ', self.total_factures)
                # print(percentage_holder.getpercentage())
                kwargs['mainmenu'].update_idletasks()
                # print('after idletasks()')

            if self.ginv_of_pooling[root][self.file_type]['file']=='Not Found':
                continue
            try:
                # print('list: ', self.ginv_of_pooling[root][self.file_type]['file'])
                for file in self.ginv_of_pooling[root][self.file_type]['file']:
                    # print('one by one: ', file)

                    self.destionation_path=Path(self.overall_path['destination']['parent'] / str(root) )
                    self.destionation_path.mkdir(parents=True, exist_ok=True)
                    self.file_full_name=Path(file)
                    self.only_file_name=self.prefix_dict[self.file_type] + Path(self.file_full_name).name

                    # Add prefix in front of destination pdf file
                    self.destination_full_name = self.destionation_path / self.only_file_name

                    if self.file_type=='facture':
                            for file in self.destination_full_name.parent.rglob('1-*.pdf'):
                                if not file==self.destination_full_name:
                                    file.unlink()

                    # Check if user wants to replace or not
                        # User wants to replace
                        # if self.want_replace:
                        #     self.destination_full_name.unlink()
                        # else: # Copy only new pdf files
                        #     if not self.destination_full_name.exists():
                        #         self.dest_file_type_counter=0
                        #         # Check if there is a same file type, if there is delete it, because it is possible that file name is changed.
                        #         for _ in self.destination_full_name.parent.rglob(self.prefix_dict[self.file_type] + '*.pdf'):
                        #             self.dest_file_type_counter+=1
                        #             print(f'Counter is: {self.dest_file_type_counter}')

                        #         # If there is a same file type of currently ongoing file, then delete it.
                        #         if self.dest_file_type_counter>0:
                        #             # Delete files and copy new one
                        #             for pdffile in self.destination_full_name.parent.rglob(self.prefix_dict[self.file_type] + '*.pdf'):
                        #                 pdffile.unlink()

                    # print([self.file_full_name, self.destination_full_name])
                    self.iter_list_for_copying.append([self.file_full_name, self.destination_full_name])
                    # shutil.copy(self.file_full_name, self.destination_full_name)
                
            except:
                self.test_delete_later=self.ginv_of_pooling[root][self.file_type]['no']
                print(f'Error in {root} - {self.only_file_name} - {self.test_delete_later}')
                pass
        
        # Debug
        # print('self.iter_list_for_copying')
        # print(self.iter_list_for_copying)
        # pbar.finish()

        with concurrent.futures.ThreadPoolExecutor(max_workers=self.total_physical_core-1) as executor:
            executor.map(self.pool_executor, self.iter_list_for_copying)

    def pool_executor(self, source_dest):
        if self.want_replace:
            source_dest[1].unlink()
            # print('replacing: ', source_dest[0], source_dest[1])
            shutil.copy(source_dest[0], source_dest[1])
        else:
            if not Path(source_dest[1]).exists():
                print(f'copying: {source_dest[0]} -  {source_dest[1]}')
                shutil.copy(source_dest[0], source_dest[1])
    
    def update_ginv_sheet(self):
        for root in self.not_found_files_dict[self.file_type]:
            self.ws.cell(row=self.ginv[root]['row'], column=self.not_found_columns[self.file_type], value='NOT FOUND')
        # print(type(self.general_invoice_full_name))
        # print(self.general_invoice_full_name)
        try:
            self.wb.save(self.general_invoice_full_name)
        except:
            print('I could not save {} \n Probably, it is open, close and restart application.'.format(self.general_invoice_full_name.name))

if __name__ == '__main__':
    # start = time.perf_counter()
        # files_list_to_copy=['facture', 'routage', 'decl', 'tds', 'coo']
            # stime={}
            # ftime={}
            # for file_type in files_list_to_copy:

            #     stime.update({file_type: time.perf_counter()})
            #     # if file_type=='routage':
                
            #     test = FileFinder(file_type)
            #     test.find_files()
            #     test.copy_files()
            #     ftime.update({file_type: time.perf_counter()})
            #     finish=time.perf_counter()
            

            # print('Total time spent: {}'.format(round(finish-start,2)))
            # for xtime in stime:
            #     try:
            #         print('Total time spent for {}: {}'.format(xtime, round(ftime[xtime]-stime[xtime],2)))
            #     except:
            #         pass
            # Save user's selections to a file
                # with open(Path('users_selections') / 'mt_folder.txt', 'w') as f:
                #     f.write(self.mt_folder_value.get())
                # with open(Path('users_selections') / 'ginv.txt', 'w') as f:
                #     f.write(self.ginv_excel_value.get())

                # self.start = time.perf_counter()
                # self.finish = time.perf_counter()
                # users_selection=[self.facture.get(),self.routage.get(),self.decl.get(), self.tds.get(), self.coo.get()]
        files_type_to_copy=[]
        total_for_progress={}
        users_selection=['facture', 'routage', 'decl', 'tds', 'coo']
        
        for slctn in users_selection:
            if slctn!='0':
                files_type_to_copy.append(slctn)

        # Debug
        print(files_type_to_copy)

        # Progress Bar
            # progress=Progress(tab)

            # stime={}
            # ftime={}

            # Loop through user's seledtions one by one
            # and loop through inside each file type
            # percentage_holder.resetpercentage()
            # tab.update_idletasks()
        temp_mt_folder = r'D:\BYTK_Facturation\7. MT'
        temp_ginv = r'D:\BYTK_Facturation\7. MT\GENERAL INVOICE rev19 - cri 037 rev A.xlsm'
        for file_type in files_type_to_copy:

            # stime.update({file_type: time.perf_counter()})

            # Check if user selected right folder and right excel file
            file_finder = FileFinder(file_type, mt_folder=Path(temp_mt_folder), wb=Path(temp_ginv))
            if file_finder.is_wb_error:
                # pymsgbox.alert('You did not choose excel file')
                # finish=time.perf_counter()
                break

            if not file_finder.all_folders_in_place:
                if file_finder.partially_found:
                    # pymsgbox.alert(text=file_finder.partially_found_alert_message)
                    if not file_type in file_finder.found_paths:
                        continue
                else:
                    # alert_text=file_finder.none_found_alert_message
                    # alert_text+='\n'
                    # alert_text+='Program will quit, please choose proper folder and try again later.'
                    # pymsgbox.alert(text=alert_text)
                    break

            # total_facture=file_finder.total_factures

            file_finder.find_files()
            file_finder.copy_files()
            file_finder.update_ginv_sheet()
    # Check users selection, if there is no requested folders inside selection folder, then warn
        # if file_finder.partially_found:
        #     alert_text='These files are not processed, because there is no folder related to them'
        #     alert_text+='\n'
        #     alert_text+=str(list(map(lambda x: x, file_finder.not_found_paths)))

            # Check if user requested the file type that was not found
            # for f in file_finder.not_found_paths:
            #     if f not in files_type_to_copy:
            #         file_finder.not_found_paths.remove(f)
            
            # alert_text+='\n'
            # alert_text+='\n'
            # alert_text+='But luckily this files has been processed successfully: '
            # alert_text+=str(list(map(lambda x: x, file_finder.found_paths)))

            # pymsgbox.alert(text=alert_text)

            # ftime.update({file_type: time.perf_counter()})
            # finish=time.perf_counter()
        

        # print('Total time spent: {}'.format(round(finish-start,2)))
        # for xtime in stime:
        #     try:
        #         print('Total time spent for {}: {}'.format(xtime, round(ftime[xtime]-stime[xtime],2)))
        #     except:
        #         pass

