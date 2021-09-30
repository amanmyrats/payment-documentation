from pathlib import Path
from openpyxl import Workbook, load_workbook
import os
import re
import threading
import time
import pandas as pd
import shutil
from multiprocessing import Pool
from FilesClass import AllFiles

from helper import *



# ginv = {}
all_files = {}


class FindFacture(AllFiles):
    # def __init__(self):
    #     self.assign_all_files_into_list()
    #     super().__init__(self)
    #     #print(self.ginv)
    #     #self.ginv=super().ginv

    def find_files(self):
        self.facture_list=[]
        
        for pdf in self.overall_path['source']['facture'].glob('*.pdf'):
            self.facture_list.append(str(pdf))

        self.df_facture=pd.DataFrame(self.facture_list, columns=['facture_no'])
        #print(df.head())

        self.iter_list_for_pooling=[]
        for root in self.ginv:
            self.iter_list_for_pooling.append(root)

        # Testing purpose
        self.df_facture['pure_no']='Empty'
        self.df_facture['modife']='Empty'

        p=Pool()
        self.result=p.map(self.pool_thread, self.iter_list_for_pooling)
        p.close()
        p.join()
        
        self.result_ginv={}
        for dcts in self.result:
            # self.temp_key=dcts.keys()
            # self.temp_value=dcts.values()
            self.result_ginv.update(dcts)
            # print(f'key: {self.temp_key} value: {self.temp_value}')


        # print(self.result_ginv)
        # print(self.result_ginv.keys())


    def pool_thread(self, root):
        # for root in kwargs['ginv']:
        if self.ginv[root]['facture']!='':
            for _ in self.ginv[root]['facture']:
                self.ginv[root]['facture']['file']='Not Found'
                self.no_to_find=self.ginv[root]['facture']['no']
                self.df_to_search=self.df_facture
                self.pattern1=r'(\D|\b)' + str(self.no_to_find) + r'\D.+pdf$'
                self.pattern2=r'[^-]' + str(self.no_to_find) + r'\D.+pdf$'
                # self.pattern1='\(\\D|\\b\)' + str(self.no_to_find) + '\\D.+pdf$'
                # self.pattern2='[^-]' + str(self.no_to_find) + '\\D.+pdf$'
                self.match_result=self.df_to_search[self.df_to_search['facture_no'].str.contains(self.pattern1)==True]
                self.match_result=self.match_result[self.match_result['facture_no'].str.contains(self.pattern2)==True]

                if self.match_result.shape[0]>1:
                    # Write all modife numbers into dataframe
                    for xindex in self.match_result.index.tolist():
                        self.df_facture.loc[xindex,'pure_no']=root

                        self.splitted=re.split(r'mod\D+', self.match_result.at[xindex,'facture_no'])
                        print(f'  --  {self.no_to_find}  --  {self.splitted}')
                        if len(self.splitted)>1:
                            try:
                                self.modife_number=self.splitted[1][0]
                            except:
                                pass

                            if self.modife_number.isdigit():
                                self.df_facture.loc[xindex,'modife']=self.modife_number
                        else:
                            self.df_facture.loc[xindex,'modife']=0
                        
                    # Loop dataframe of current facture and select only last modife, then delete other old modifes
                    self.latest_modife=0

                    self.all_current_root =self.df_facture[(self.df_facture['pure_no']==root) & (self.df_facture['modife']==self.latest_modife)]
                    try:
                        self.ginv[root]['facture']['file'] =self.all_current_root.iat[0,0]
                    except:
                        pass
                    # print('TEST', self.all_current_root.iat[0,0])
                    for xindex in self.match_result.index.tolist():
                        if int(self.df_facture.loc[xindex]['modife'])>int(self.latest_modife):
                            self.latest_modife=self.df_facture.loc[xindex]['modife']
                            self.all_current_root =self.df_facture[(self.df_facture['pure_no']==root) & (self.df_facture['modife']==self.latest_modife)]
                            try:
                                self.ginv[root]['facture']['file']=self.all_current_root.iat[0,0]
                            except:
                                pass

                else:
                    self.df_facture.loc[self.match_result.index.tolist(),'pure_no']=root
                    self.df_facture.loc[self.match_result.index.tolist(),'modife']=0
                    self.all_current_root =self.df_facture[(self.df_facture['pure_no']==root) & (self.df_facture['modife']==0)]
                    try:
                        self.ginv[root]['facture']['file']=self.all_current_root.iat[0,0]
                    except:
                        pass
                    # print('\n - root {} \n - no: {} \n - file_address: {}'.format(root,self.ginv[root]['facture']['no'], self.ginv[root]['facture']['file']))

                # Returning a dictionary for facture_path_and_name
                # print('\n - root {} \n - no: {} \n - file_address: {}'.format(root,self.ginv[root]['facture']['no'], self.ginv[root]['facture']['file']))
                #os.system('pause')
                # print(f'General invoice {self.ginv}')
                return {root : self.ginv[root]}
        

    def copy_files(self):
        print(f'Result general invoice {self.result_ginv}')
        for root in self.result_ginv:
            if root==131815:
                print('Possible issue')
            try:
                self.destionation_path=Path(self.overall_path['destination']['parent'] / str(root) )
                self.destionation_path.mkdir(parents=True, exist_ok=True)
                # file_full_name='GENERAL INVOICE rev15.xlsm'
                file_full_name=self.result_ginv[root]['facture']['file']
                # destination_full_name = self.overall_path['destination']['parent'] / 'GENERAL INVOICE rev15.xlsm'
                destination_full_name = self.destionation_path / Path(file_full_name).name
                shutil.copy(file_full_name, destination_full_name)
            except:
                print(f'Error in {root}')
                pass
        pass

    def assign_all_files_into_list(self):
        for file_type in overall_path['source']:
            all_files[file_type] = {}
            for xfile in overall_path['source'][file_type].rglob('*.pdf'):

                all_files[file_type].update({xfile.name: xfile})


if __name__ == '__main__':
    start = time.perf_counter()

    test = FindFacture()
    test.find_files()
    test.copy_files()

    finish=time.perf_counter()

    print('Total time spent: {}'.format(round(finish-start,2)))
