from pathlib import Path


class LastModifeFinder:
    """
        test = LastModifeFinder(parent_path = r'D:\\BYTK_Facturation\\7. MT\\1-FACTURE')
        test.find_last_modife(facture_no = '0018-BYTK', file_name_list=file_names_list)

        print('This is last modified: ', test.last_modife)
    """

    def __init__(self, *args, **kwargs):
        self.parent_path = kwargs.get('parent_path')
        self.words_to_search = ['mo', 'rev']
        

        try:
            kwargs['parent_path']
            if not kwargs['parent_path']=='' and Path(kwargs['parent_path']).exists():
                self.search_according_to_time=True
        except:
            self.search_according_to_time=False
    
    def find_last_modife(self, *args, **kwargs):
        """
        test.find_last_modife(facture_no = '0018-BYTK', file_name_list=file_names_list)
        print('This is last modified: ', test.last_modife)
    """
        # print('inside: find_last_modife')
        self.facture_no=kwargs['facture_no']
        self.file_name_list=kwargs['file_name_list']
        self.match_type=kwargs.get('match_type')

        self.last_modife = 'none'
        self.last_modife_found_name='none'
        self.last_modife_found_time='none'

        self.search_according_to_time=False
        self.is_time_error=False

        self.files_dict={}
        
        if len(self.file_name_list)==1:
            self.last_modife=self.file_name_list[0]
        elif len(self.file_name_list)==0:
            self.last_modife = 'none'
        else:
            # print('inside last modife: ', self.file_name_list)
            for file_name in self.file_name_list:
                # print(self.file_name_list)
                # print('inside: ', file_name)
                try:
                    self.files_dict[file_name]
                except:
                    self.files_dict[file_name]={}
                    
                self.files_dict[file_name]['path']=Path(self.parent_path) / file_name
                self.files_dict[file_name]['modife_number']=0
                self.files_dict[file_name]['modife_time']=0
                self.files_dict[file_name]['file_name']=file_name
                self.files_dict[file_name]['facture_no']=self.facture_no

            # print('# Call search functions here: ')
            # Call search functions here
            if self.search_according_to_time:
                # print('before calling if')
                self.last_modife_by_name()
                if not self.is_time_error:
                    self.last_modife_by_time()
                # print('before calling else')
            else:
                self.last_modife_by_name()

            # Analyze results, if there is not result by name then assign result of by time to the self.last_modife
            if self.last_modife_found_name=='none':
                self.last_modife=self.last_modife_found_time
            else:
                self.last_modife=self.last_modife_found_name

            # print('found last modife: ', self.last_modife)

    def last_modife_by_time(self):
        
        self.old_mod_time=0
        for file in self.files_dict:
            try:
                self.mod_time=self.files_dict[file]['path'].stat().st_mtime
                # print('mod_time of {}'.format(self.files_dict[file]['path'].name), self.mod_time)
                if self.mod_time > self.old_mod_time:
                    self.last_modife_found_time=self.files_dict[file]['path'].name
                    self.last_modife=self.last_modife_found_time
                self.old_mod_time=self.mod_time
            except:
                print('Error when matching according to time: ', file)
                self.is_time_error=True

    def last_modife_by_name(self):
        
        # print('inside last_modife_by_name')
        for file in self.files_dict:
            self.name_to_search=str(self.files_dict[file]['file_name'])
            self.name_to_find=str(self.files_dict[file]['facture_no'])
            self.start_counting_from=self.name_to_search.find(self.name_to_find) + len(self.name_to_find) + 1
            self.length_of_file_name=len(self.files_dict[file]['file_name'])
            for i in range(self.start_counting_from, self.length_of_file_name + 1):
                self.alone_char=str(self.files_dict[file]['file_name'])[i-1]
                if str(self.alone_char).isdigit():
                    self.files_dict[file]['modife_number']=self.alone_char
                    break
            
        self.old_mod_number=0
        for file in self.files_dict:
            self.new_mod_number=int(self.files_dict[file]['modife_number'])
            if self.new_mod_number > self.old_mod_number:
                self.last_modife_found_name=self.files_dict[file]['file_name']
                self.last_modife=self.files_dict[file]['file_name']
                # break

            self.old_mod_number=self.new_mod_number

            # print('start counting: ', self.start_counting_from)


if __name__ == '__main__':
    file_names = ['Chantier__0018-BYTK TK.xlsx',
                  'Chantier__0018-BYTK TK modife 2.xlsx', 'Chantier__0018-BYTK TK modife 1.xlsx']

    test = LastModifeFinder(parent_path=r'D:\BYTK_Facturation\7. MT\1-FACTURE')
    test.find_last_modife(facture_no='0018-BYTK', file_name_list=file_names)

    print('\nThis is last modified by time: ', test.last_modife_found_time)
    print('This is last modified by name: ', test.last_modife_found_name)

    print('\nThis is last modified: ', test.last_modife)
