import pandas as pd
import openpyxl
from openpyxl import load_workbook
from pathlib import Path
import psutil
import re


overall_path = {'source': {'facture': Path.cwd() / '1-FACTURE/1-FACTURE PDF',
                           'routage': Path.cwd() / '2-CMR',
                           'decl': Path.cwd() / '3-DECLARATION',
                           'tds': Path.cwd() / '4-TDS',
                           'coo': Path.cwd() / '5-CO'},
                'destination': {'parent': Path.cwd() / '00_All in One',
                                'sub': Path.cwd() / ''}
                }

# ginv = {'root_facture': {'facture': {'facture_no': '',
    #                                      'facure_file': ''},
    #                          'routage': {'routage_no': '',
    #                                      'routage_file': ''},
    #                          'decl': {'decl_no': '',
    #                                   'decl_file': ''},
    #                          'tds': {'tds_no': '',
    #                                  'tds_file': ''},
    #                          'coo': {'coo_no': '',
    #                                  'coo_file': ''}
    #                          }

    #         }


def next_available_name(full_path):
    full_path=Path(full_path)
    fname=full_path.name
    fstem=full_path.stem
    fsuffix=full_path.suffix
    i=0
    while full_path.exists():
        i+=1
        fname=fstem + ' - ' + str(i) + fsuffix
        full_path=full_path.parent / fname
    
    # return full_path.parent / fname
    return fname

def next_available_folder_name(full_path):
    initial_full_path=Path(full_path)
    full_path=Path(full_path)
    i=0
    while full_path.exists():
        i+=1
        new_folder_name=''
        new_folder_name=str(initial_full_path.parts[-1]) + ' - '  + str(i)
        full_path=full_path.parent / new_folder_name

    return full_path

def available_ram(type):
    total_ram_byte = psutil.virtual_memory()._asdict()['total']
    avail_ram_byte = psutil.virtual_memory()._asdict()['available']
    if type=='mb':
        return avail_ram_byte / 1024 / 1024
    elif type=='gb':
        return avail_ram_byte / 1024 / 1024 / 1024
    elif type=='percentage':
        return round(avail_ram_byte / total_ram_byte * 100,0)

def facture_miner(*args):
    # 020-TK-NCC 50%
    # NCCI-TK-2470 (20%)
    facture_pattern1 = r'[a-zA-Z]{2,3}-[a-zA-Z]{2}-\d{3,5}'
    facture_pattern2 = r'\d{3,5}-[a-zA-Z]{2}-[a-zA-Z]{2,3}'
    if not len(re.findall(facture_pattern1, str(args[0])))==0:
        return re.findall(facture_pattern1, str(args[0]))
    else:
        return re.findall(facture_pattern2, str(args[0]))

def routage_splitter(*args):
    # Check the task of slash, does it seperate two routage
    # or is it only seperator one routage.
    check_slash_pattern = r'[/-]\d{2}[a-zA-Z]{2,3}'
    slash_as_order_seperator_pattern = r'\d{2}[a-zA-z]{2,3}\d{2,3}(?:[/-]{0,1}\d{1,2}){0,1}'
    slash_as_routage_serepator_pattern = r'\d{2}[a-zA-z]{2,3}\d{2,3}'

    if len(re.findall(check_slash_pattern, str(args[0])))>0:
        # Then the slash is only routage seperator,
        # print('As seperator {}'.format(args[0]))
        # print('Seperator result {}'.format(re.findall(slash_as_routage_serepator_pattern, args[0])))
        return re.findall(slash_as_routage_serepator_pattern, str(args[0]))
    else:
        # Slash is order seperator
        # print('As order slash {}'.format(args[0]))
        # print('Order slash result {}'.format(re.findall(slash_as_order_seperator_pattern, args[0])))
        return re.findall(slash_as_order_seperator_pattern, str(args[0]))

def tds_splitter(*args):
    tds_pattern = r'\d+\.\d+\.\d+\.\d+'
    # test = str(args[0])
    # print(f'it is the type {type(args[0])} and type converted {type(test)} and name {args[0]}')
    return re.findall(tds_pattern, str(args[0]))

def decl_splitter(*args):
    decl_pattern = r'\d+[/-]\d+[/-]\d+'
    return re.findall(decl_pattern, str(args[0]))

class ExcelAnalyzer:
    """
        arguments: self.excel = kwargs['excel']

        then call function: analyze_for_ygt()

    """
    def __init__(self, *args, **kwargs):
        self.excel = kwargs['excel']
        self.what_is_it = 'unknown'

    def analyze_for_ygt(self):
        try:
            self.df_situation = pd.read_excel(self.excel, sheet_name='MATERIAL TABLE')
            self.what_is_it='ygt'
        except:
            # ERROR NOT YGT SITUATION
            self.what_is_it='no'

