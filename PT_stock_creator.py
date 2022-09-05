from email import header
from os import listdir, getcwd, system
from os.path import isfile, join
from getpass import getpass
import requests
import warnings
# from collections import namedtuple

from openpyxl import load_workbook
import pandas as pd

# TODO
# 1) decorators - get_file, task_start/end

global attrib_dict

attrib_dict = {
    'POI name': 'name',
    'POI ID': 'stockable_id',
    'POI SysID': 'sys_id',
    'Link': 'url',
    'Format': 'format',
    'Tag': 'tag',
    'Length': 'len',
    'MW': 'mw',
    'pI': 'pI',
    'A0.1% (Ox)': 'extinction_ox',
    'A0.1% (Red)': 'extinction_red',
    'Concentration': 'concentration',
    'Total volume': 'vol_total',
    'Total mass': 'mass_total',
    'Description': 'description',
    'Stock ID': 'id',
    'Stock name': 'name',
    'Stock volume': 'volume',
    'Stock mass': 'weight',
    'Box ID': 'storage_id',
    'Box name': 'box',
    'Position': 'position',
    'Description': 'description',
    'Purification method': 'purification_method',
    'Storage buffer': 'storage buffer',
    'Source': 'species',
    'Sequence': 'sequence'
}


class Protein:
    
    def __init__(self, data: dict) -> None:
        self.ws_name = data.get("ws_name", "")
        self.name = data.get("name", "")
        self.prot_id = data.get("prot_id", "")
        self.description = data.get("description", "")
        self.owner_id = data.get("owner_id", "")
        self.alternative_name = data.get("alternative_name", "")
        self.gene = data.get("gene", "")
        self.species = data.get("species", "")
        self.mutations = data.get("mutations", "")
        self.chemical_modifications = data.get("chemical_modifications", "")
        self.tag = data.get("tag", "")
        self.purification_method = data.get("purification_method", "")
        self.mw = data.get("mw", "")
        extinction_ox = data.get("extinction_ox", "")
        extinction_red = data.get("extinction_red", "")
        self.extinction_coefficient_280nm = f'{extinction_ox} (Ox) / {extinction_red} (Red)'
        self.storage_temperature = data.get("storage_temperature", "")
        self.sequence = data.get("sequence", "")

    def __generate_prot_item(self) -> dict:
        
        item_list = ["name",
                     "description",
                     "owner_id",
                     "alternative_name",
                     "gene", "species",
                     "mutations",
                     "chemical_modifications", 
                     "tag",
                     "purification_method",
                     "mw",
                     "extinction_coefficient_280nm",
                     "storage_buffer",
                     "storage_temperature",
                     "sequence",
                     
            ]
        
        item = dict((k, v) for k, v in self.__dict__.items() if k in item_list)
        return item

    def row_gen(self, ws, min_row: int, min_col: int, n_row=None,  n_col=None, i=0):
        """
        generator which retuns i and row list
        """
        max_row = min_row+n_row if n_row is not None else n_row
        max_col = min_col+n_col if n_col is not None else n_col
        
        for row in ws.iter_rows(
                min_row=min_row, 
                min_col=min_col, 
                max_row=max_row,
                max_col=max_col):
            i += 1
            yield i, row

    def col_gen(self, ws, min_row: int, min_col: int, n_row=None, n_col=None, i=0):
        """
        generator which retuns i and row list
        """
        max_row = min_row+n_row if n_row is not None else n_row
        max_col = min_col+n_col if n_col is not None else n_col
        
        for row in ws.iter_cols(
                min_row=min_row, 
                min_col=min_col, 
                max_row=max_row,
                max_col=max_col):
            i += 1
            yield i, row

    def get_stock_header(self, ws) -> dict:
        min_row = 0
        min_col = 5
        n_row = 0
        
        header = {}
        
        for i, col in self.col_gen(
            ws=ws,      
            min_row=min_row, 
            n_row=n_row, 
            min_col=min_col,
            i=min_col-1):
            header[col[0].value] = i
        
        return header   
                
    def create_record(self):
        
        url = "https://my.labguru.com/api/v1/proteins"
        body = {"token": TOKEN,
                "item": self.__generate_prot_item()}

    # TODO Turn on
    
        # session = requests.post(url, json=body)
            
        # if session.status_code == 201:
        #     response = session.json()
        #     self.prot_id = response.get("id", "")
        #     self.uuid = response.get("uuid", "")
        #     self.sys_id = response.get("sys_id", "")
        #     self.url = response.get("url", "")
        #     self.class_name = response.get("class_name", "")
        #     print(f'{self.sys_id:>10s} | {self.name:<50s} - New protein entry added')
            
        # else:
        #     print(f'Error while handling {self.name} - Code {session.status_code}')

    def update_record(self):
        url = f"https://my.labguru.com/api/v1/proteins/{self.prot_id}"
        body = {"token": TOKEN,
                "item": self.__generate_prot_item()}
    
        session = requests.put(url, json=body)
            
        if session.status_code == 201:
            print(f'{self.sys_id:>10s} | {self.name:<50s} - Labguru protein record updated')
            
        else:
            print(f'Error while handling {self.name} - Code {session.status_code}')

    def get_stock_df(self, wb):
        min_row = 0
        n_row = 100
        min_col = 5
        n_col = 9
        
        ws = wb[self.ws_name]
        
        data = [[i.value for i in j] for j in ws.iter_rows(
        min_row=min_row, 
        min_col=min_col, 
        max_row=min_row+n_row-1,
        max_col=min_col+n_col-1)
            ]
        
        stock_df = pd.DataFrame(data=data[1:], columns=data[0]).dropna(how="all")
        return stock_df

    def __generate_stock_items(self, df: pd.DataFrame):
        """
        
        """
        for index, row in df.iterrows():
            if (box_id := row["Box ID"]) is not None:
                stock_id = row["Stock ID"]
                item = {"name": row["Stock name"],
                        "storage_id": int(box_id),
                        "storage_type": "System::Storage::Box",
                        "stockable_type": "Biocollections::Protein",
                        "stockable_id": self.prot_id,
                        "description": row["Description"],
                        # "barcode": "",
                        # "stored_by": "",
                        "concentration": row["Concentration"],
                        "concentration_prefix": "",
                        "concentration_unit_id": 9, # mg/mL
                        "concentration_exponent": "",
                        "concentration_remarks": "",
                        "volume": int(row["Stock volume"]),
                        "volume_prefix": "",
                        "volume_unit_id": 8, # uL
                        "volume_exponent": "",
                        "volume_remarks": "",
                        # "weight": "",
                        # "weight_prefix": "",
                        # "weight_unit_id": 1,
                        # "weight_exponent": "25",
                        # "weight_remarks": "weight remarks"
                    }
                
                yield index+1, stock_id, item
            else:
                continue

    def create_stocks(self, wb):
                
        self.added_stocks = []
        
        stock_df = self.get_stock_df(wb)
        stock_items = self.__generate_stock_items(stock_df)
        url = 'https://my.labguru.com/api/v1/stocks'
        
        for i, stock_id, item in stock_items:
            body = {'token': TOKEN,
                    'item': item}

            box_id = item["storage_id"]
            print(item)
       
            session = requests.post(url, json=body)
            if session.status_code == 201:
                response = session.json()
                
                stock_id = response['id']
                position = response['position']
                box_name = response['box']['name']
            
                new_stock = {
                    'index': i + 1,
                    'id': (stock_id, f'https://my.labguru.com/storage/stocks/{stock_id}'),
                    'box': (box_name, f'https://my.labguru.com/storage/boxes/{box_id}'),
                    'position': position, 
                }
                self.added_stocks.append(new_stock)
                
                print(f'\tStock {i} (ID: {new_stock["id"][0]:>6}) - Box: {new_stock["box"][0]}, Position: {new_stock["position"]}')
            else:
                    print(f'Error while handling {self.name}: Stock {i} - Code {session.status_code}')      

    def update_stocks(self, wb):
                                
            stock_df = self.get_stock_df(wb)
            stock_items = self.__generate_stock_items(stock_df)

            for i, stock_id, item in stock_items:
                url = f'https://my.labguru.com/api/v1/stocks/{stock_id}'
                body = {'token': TOKEN,
                        'item': item}
                session = requests.put(url, json=body)
                if session.status_code == 201:                  
                    print(f'\tStock {i} (ID: {stock_id:>6}) - stock attributes updated')
                else:
                    print(f'Error while handling {self.name}: Stock {i} - Code {session.status_code}')    

    def update_sheet_prot(self, wb):
        ws = wb[self.ws_name]
        prot_dict = dict((k, v) for k, v in self.__dict__.items())
        for i, cells in self.row_gen(
            ws=ws,
            min_row=0,
            min_col=0,
            n_col=2
            ):

            attrib, value = cells[0].value, cells[1].value

            if attrib is not None and value is None:
                print(attrib_dict[attrib])
                new_value = prot_dict.get(attrib_dict[attrib], "")
                print(attrib, new_value)
                try:
                    if 'https:' not in new_value:
                        ws.cell(row=i, column=2).value = new_value
                    else:
                        ws.cell(row=i, column=2).value = 'LINK'
                        ws.cell(row=i, column=2).hyperlink = new_value
                except TypeError as e:
                    pass
        
    def update_sheet_stocks(self, wb):
        ws = wb[self.ws_name]

        header = self.get_stock_header(ws)
        print(header)
        
        for stock in self.added_stocks:
            row = stock['index']
            columns = ['Stock ID', 'Box name', 'Position']
            for name in columns:
                column = header[name]
                attrib = attrib_dict[name]
                value = stock[attrib]
                print(f'Row: {row:>3}\tColumn: {column:>3}')
                if isinstance(value, tuple):
                    value, hyperlink = value
                    ws.cell(row=row, column=column).value = value
                    ws.cell(row=row, column=column).hyperlink = hyperlink
                else:
                    ws.cell(row=row, column=column).value = value


def main():
    system('cls')
    global TOKEN
    global PB_ALL
    global TEST_MODE 
    
    TEST_MODE = False
    # token = get_token()
    TOKEN = 'token'
    PB_ALL = r'P:\_research group folders\PT Proteins\_PT_stock_creator' # TODO

    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

    while True:
        system('cls')
        menu_dict = {
            '1': ['Create new template file', 
                generate_template],
            '2': ['Add new protein records and stocks to Labguru',
                add_to_lg],
            '3': ['Update existing Labguru records and stocks',
                update_to_lg],
            '4': ['Create excel file for label printing',
                create_label_xlsx]
                }

        # Print menu
        print_menu(menu_dict)

        user_input = input('\nTask number: ')
        
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
    file = choose_file(mypath, xlsx_list, ext)
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

def choose_file(path, file_list: list, ext: str):
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
            print(f'\nFile updated:          {file}')
            break
        except PermissionError:
            print(f'---< Unable to save changes in {file} >---'.center(80))
            print('Close the file and continue'.center(80))
            system('pause')
        finally:
            wb.close()


# API requests

def get_token(token=None):
    if token is None:
        print('\nEnter Labguru credentials: name (n.surname) and password')
        while True:

            if test_mode is True:
                name = 'LG_User1@purebiologics.com'
                password = 'LG_user1'
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
def get_protein_data(ws, attrib_dict: dict) -> dict:

    protein_data = {'ws_name': ws.title}

    for row in ws.iter_rows(min_row=0,
                            min_col=0,
                            max_col=2,
                            values_only=True):
        if row[0] is not None:
            protein_data[attrib_dict[row[0]]] = row[1]
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
    
    wb = load_workbook(join(mypath, file), data_only=True)
    for ws in wb.sheetnames:
        protein_data = get_protein_data(wb[ws], attrib_dict)
        protein = Protein(protein_data)
        if not protein.prot_id:
            protein.create_record()
            prot_count += 1
            
        protein.create_stocks(wb)
        prot_list.append(protein)
        
        update_workbook(mypath, file, prot_list)

    stock_count = sum([len(prot.added_stocks) for prot in prot_list])

    wb.close()

    print(f'Protein entries created: {prot_count}')
    print(f'Stocks added:            {stock_count:}\n')
    
    task_end()

# 3) Update existing records

def update_to_lg():

    mypath, file = get_path_file('xlsx')
    if file is None:
        return None
    
    task_start(file)
    
    wb = load_workbook(join(mypath, file), data_only=True)
    for ws in wb.sheetnames:
        protein_data = get_protein_data(wb[ws], attrib_dict)
        protein = Protein(protein_data)
        if not protein.prot_id:
            protein.update_record()
            prot_count += 1
            
        protein.update_stocks(wb) 
    wb.close()

# 4) Labels

def create_label_xlsx():
      
    mypath = getcwd()
    ext = 'xlsx'
    xlsx_list = scan_files(mypath, ext)
    file = choose_file(mypath, xlsx_list, ext)
    if file is None:
        return None
    
    task_start(file)
    
    wb = load_workbook(join(mypath, file), data_only=True)
    columns = ['Stock ID', 'Stock name', 'Concentration', 'Stock volume', 'Stock mass']
    dfs = []
    label_file = file.replace('.xlsx', '_labels.xlsx')
    
    for ws in wb.sheetnames:
        protein_data = get_protein_data(wb, ws, attrib_dict)
        protein = Protein(protein_data)
        stock_df = protein.get_stock_df(wb)
        dfs.append((ws, stock_df.loc[:, columns]))
    
    with pd.ExcelWriter(label_file) as writer:
        for ws, df in dfs:
            df.to_excel(writer,
                        sheet_name=ws,
                        index=False,
                        header=False)

    wb.close()
    print(f'File saved as {label_file}')
    task_end()

if __name__ == '__main__':
    main()
