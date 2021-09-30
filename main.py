from pathlib import Path

from ginv import AssignGInvoice

overall_path = {'source': {'facture': Path.cwd() / '1-FACTURE/1-FACTURE PDF',
                           'routage': Path.cwd() / '2-CMR',
                           'decl': Path.cwd() / '3-DECLARATION',
                           'tds': Path.cwd() / '4-TDS',
                           'coo': Path.cwd() / '5-CO'},
                'destination': {'parent': Path.cwd() / '00_All in One',
                                'sub': Path.cwd() / ''}
                }


if __name__=='__main__':
    test=AssignGInvoice()
