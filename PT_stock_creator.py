from os import listdir, getcwd, system
from os.path import isfile, join
from getpass import getpass
from re import findall
import requests
import warnings

from openpyxl import load_workbook
import pandas as pd

from src.classes import Protein, attrib_dict

# TODO
# 1) decorators - get_file, task_start/end
# 2) don't add stocks already with IDs
# 3) update excel from LG
# 4) select sheets for update from/to LG
# 5) if sequence - calculate mass, len, pI, A0.1%


def main():
    system('cls')
    global TOKEN
    global PB_ALL

    test_mode = select_mode()

    TOKEN = get_token(test_mode=test_mode)
    # TOKEN = 'token'
    PB_ALL = r'P:\_research group folders\PT Proteins\_PT_stock_creator'

    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

    while True:
        system('cls')
        menu_dict = {
            '1': ['Create new template file', 
                generate_template],
            '2': ['Add new protein records and stocks to Labguru',
                add_to_lg],
            '3': ['Update Labguru records and stocks from excel',
                update_to_lg],
            '4': ['Update excel records and stocks from Labguru',
                update_from_lg],
            '5': ['Create excel file for label printing',
                create_label_xlsx]
                }

        # Print menu
        print_menu(menu_dict)

        user_input = input('\nSelect task number: ')
        
        if user_input.lower() == "q":
            break
        elif user_input in menu_dict:
            system('cls')
            menu_dict[user_input][1]()
            system('pause')
        else:
            print('---< Wrong input. Try again >---'.center(80))
            system('pause')

# Helpers

def select_mode():
    while True:
        mode_selection = input('Test mode (Y/N): ').lower()
        if 'y' in mode_selection:
            return True
        elif 'n' in mode_selection:
            return False
        else:
            print('---< Wrong input. Try again >---'.center(80))

def print_menu(menu_dict):
        print('/ PT Stock creator \\'.center(80, '_'), '\n')
        for k, v in menu_dict.items():
            print(f'{k:>2}) {v[0]}')
        print('\n Q - Quit')
        print(''.center(80, '_'))

def task_start(file = None):
    print("/ Task in progress \\".center(80, '_'), '\n', sep='')
    if file is not None:
        print(f'File: {file}\n')

def task_end():
    "Prints a bar after ending a task"
    print("\n", "\\ Task completed /".center(80, '‾'), '\n', sep="")

def pball_connection(file):
    while True:
        if isfile(join(PB_ALL, file)):
            break        
        else:
            print('---< Template file not found: check connection to PB_all >---'.center(80))
            system('pause')

def print_task(func):
    def wrap(*args, **kwargs):
        task_start()
        result = func(*args, **kwargs)
        task_end()
        return result
    return wrap

def get_path_file(ext: str):
    mypath = getcwd()
    xlsx_list = scan_files(mypath, ext)
    file = choose_file(xlsx_list, ext)
    if file is None:
        return mypath, None

    return mypath, file

def scan_files(path, ext):
    """ Returns list of xlsx files in given path """
    xlsx_list = [file for file in listdir(path) if check_filename(path, file, ext)]
    return xlsx_list

def check_filename(path, file: str, ext: str):
    """  """
    is_file = isfile(join(path, file))
    is_xlsx = file.endswith(f'.{ext}')
    not_temp = not file.startswith('~$')
    conditions = [is_file, is_xlsx, not_temp]
    return all(conditions)

def choose_file(file_list: list, ext: str):
    print('/ Select file \\'.center(80, '_'), '\n', sep='')
    
    xlsx_dict = {str(i): file for i, file in enumerate(file_list, 1)}    
    for i, file in xlsx_dict.items():
        print(f'{i.rjust(2)})', file.replace(f'.{ext}', ''), sep=" ")
        
    print('\n', 'Q - Return to menu')
    print(''.center(80, '_'), '\n', sep='')
    
    while True:
        file_i = input('File number: ').lower()
        if file_i in xlsx_dict:
            break
        elif file_i == 'q':
            return None
        else:
            print('---< Wrong input. Try again >---'.center(80))

    system('cls')
    
    return xlsx_dict[file_i]

def scan_genebank(path):
    gb_files = [f for f in listdir(path) if isfile(join(path, f)) and f.rpartition(".")[-1] == "gb"]
    return gb_files

def update_workbook(mypath, file, prot_list):
    while True:
        wb = load_workbook(join(mypath, file))
        try:
            for prot in prot_list:
                prot.update_sheet_prot(wb)
                prot.update_sheet_stocks(wb)
            wb.save(join(mypath, file))
        except PermissionError:
            print(f'---< Unable to save changes in {file} >---'.center(80))
            print('Close the file and continue'.center(80))
            system('pause')
        else:
            print(f'\nFile updated:\t\t {file}')
            break
        finally:
            wb.close()

def select_sheets(wb):
    print('\nSheets:')
    for i, sheet in enumerate(wb.sheetnames):
        print(f'   {i+1:>2d}) {sheet}')
    while True:
        try:
            selection = input('\nEnter sheet numbers or "a" to select all: ').lower()
            selection = list(findall(r"[\d]+|a", selection))
            if 'a' in selection:
                print()
                return wb.sheetnames
            elif (sheets := [wb.sheetnames[int(x)-1] for x in selection]):
                print()
                return sheets
            else:
                print('---< Wrong input. Try again >---'.center(80))
        except IndexError:
            print('---< Wrong input. Try again >---'.center(80))
            
def protein_from_sheet_gen(wb):
    ws_selection = select_sheets(wb)
    for ws_name in ws_selection:
        ws = wb[ws_name]
        protein_data = get_protein_data(ws)
        yield Protein(protein_data, TOKEN)

def if_calc_params():
    while True:
        resp = input('Calculate protein parameters for proteins with added sequence? (Y/N): ').lower()
        if 'y' in resp:
            return True
        elif 'n' in resp:
            return False
        else:
            print('---< Wrong input. Try again >---'.center(80))

# API requests

def get_token(token=None, test_mode=False):
    if token is None:
        print('\nEnter Labguru credentials: name (n.surname) and password')
        while True:

            if test_mode is True:
                name = 'LG_User7@purebiologics.com'
                password = 'LG_user7'
            else:
                name = str(input('Name: ')).lower() + '@purebiologics.com'
                password = str(getpass('Password: '))
            
            body = {'login': name,
                    'password': password}
            
            message = '\nAcquiring authentication token'
            print(f'{message.ljust(73, ".")}', end='')
            
            session = requests.post('https://my.labguru.com/api/v1/sessions.json', body)
            token = session.json()['token']

            if token and token != '-1':
                print('SUCCESS')
                break
            else:
                print('..ERROR')
                print('Wrong login or password. Try again\n')
    return token

# TODO Move to Protein and refactor
def get_protein_data(ws) -> dict:

    protein_data = {'ws_name': ws.title}

    for attrib_name, value, *row in ws.iter_rows(min_row=0,
                            min_col=0,
                            # max_row=100,
                            max_col=2,
                            values_only=True):
        try:
            if attrib_name is not None:
                lg_name = attrib_dict.get(attrib_name, None)
                if lg_name is not None:
                    protein_data[lg_name] = value
        except KeyError as e:
            print(e)
    return protein_data


# 1) New template

@print_task
def generate_template():
    """ Create copy of template file stored on PB_all """
    template_file = 'templates\\PT stock creator_template.xlsx'
    pball_connection(template_file)
    
    name = input('Name of the template file: ') + '.xlsx'
    mypath = getcwd()
    while name in listdir(mypath):
        print('---< A duplicate file name exists >---'.center(80))
        name = input('Select other name: ') + '.xlsx'
        
    system(f'copy "{join(PB_ALL, template_file)}" "./{name}" >nul')
    print(f'\nTemplate file {name} created')  

# 2) Add records and stocks

def add_to_lg():
    """
    
    """
    
    mypath, file = get_path_file('xlsx')
    if file is None:
        return None
    
    task_start(file)
    
    prot_list = []
    prot_count = 0
    params_flag = if_calc_params() # TODO add to method
    
    wb = load_workbook(join(mypath, file), data_only=True)
    # ws_selection = select_sheets(wb)
    # for ws_name in ws_selection:
    #     ws = wb[ws_name]
    #     protein_data = get_protein_data(ws)
    #     protein = Protein(protein_data, TOKEN)
    for protein in protein_from_sheet_gen(wb):
        if params_flag:
            protein.calc_params()
        if protein.id is None:
            protein.create_new_record()
            prot_count += 1
            
        protein.create_stocks(file)
        prot_list.append(protein)
        
        update_workbook(mypath, file, prot_list)

    stock_count = sum([len(prot.added_stocks) for prot in prot_list])

    wb.close()

    print(f'Protein entries created: {prot_count}')
    print(f'Stocks added:\t\t {stock_count:}\n')
    
    task_end()

# 3) Update existing records in Labguru

def update_to_lg():

    mypath, file = get_path_file('xlsx')
    if file is None:
        return None
    
    task_start(file)
    
    wb = load_workbook(join(mypath, file), data_only=True)
    for protein in protein_from_sheet_gen(wb):
        if protein.id is not None:
            protein.update_lg_record()
        protein.update_lg_stocks(file) # TODO wb->ws ?
    wb.close()

    task_end()

# 4) Update excel file from Labguru records

def update_from_lg():

    mypath, file = get_path_file('xlsx')
    if file is None:
        return None
    
    task_start(file)
    
    prot_list = []
    
    wb = load_workbook(join(mypath, file), data_only=True)
    for protein in protein_from_sheet_gen(wb):
        if protein.id is not None:
            protein.update_excel(wb)
            
        protein.update_excel_stocks(file)
        prot_list.append(protein)
        
        update_workbook(mypath, file, prot_list)

    task_end()

# 5) Labels

def create_label_xlsx():
      
    mypath = getcwd()
    ext = 'xlsx'
    xlsx_list = scan_files(mypath, ext)
    file = choose_file(xlsx_list, ext)
    if file is None:
        return None
    
    task_start(file)
    
    wb = load_workbook(join(mypath, file), data_only=True)
    columns = {'Stock ID': int,
               'Stock name': str,
               'Concentration': float,
               'Stock volume': int,
               'Stock mass': int,
               'Box name': str,
               'Position': str}
    dfs = []
    label_file = file.replace('.xlsx', '_labels.xlsx')
    
    # TODO turn into generator
    
    for protein in protein_from_sheet_gen(wb):
        stock_df = protein.get_stock_df(file).astype(columns)
        stock_df['Concentration'] = stock_df['Concentration'].map(lambda x: f'{x:.3f}')
        dfs.append((ws, stock_df.loc[:, columns.keys()]))
    
    while True:
        try:
            with pd.ExcelWriter(label_file) as writer:
                    for ws, df in dfs:
                        df.to_excel(writer,
                                    sheet_name=ws,
                                    index=True,
                                    header=False)
            break
        except PermissionError:
            print(f'---< Unable to save changes in {file} >---'.center(80))
            print('Close the file and continue'.center(80))
            system('pause')

    wb.close()
    print(f'File saved as {label_file}')
    task_end()

if __name__ == '__main__':
    main()
