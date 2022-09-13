from msilib import sequence
import requests

import pandas as pd
from Bio.Seq import Seq

from src.data import attrib_dict, aa_dict, prot_attribs


global TOKEN

class Protein:
    
    # TODO autoassign
    def __init__(self, data: dict, token: str, prot_attribs=prot_attribs) -> None:
        
        self.token = token
        
        for attrib in prot_attribs:
            setattr(self, attrib, data.get(attrib, None))

        self.sequence = str(Seq(data.get("sequence", "")))

        ext_ox = f'{self.extinction_ox} (Ox)' if self.extinction_ox else ''
        ext_red = f'{self.extinction_red} (Red)' if self.extinction_ox else ''
        self.extinction_coefficient_280nm = f'{ext_ox}{" / " if ext_ox else ""}{ext_red}'
            
        if (uniprot_id := data.get("web_page", "")):
            self.uniprot_id = uniprot_id
            self.web_page = f"https://www.uniprot.org/uniprot/{uniprot_id}"
        else:
            self.uniprot_id, self.web_page = None, None
        
        prot_url = 'https://my.labguru.com/biocollections/proteins/'
        self.url = f'{prot_url + str(self.id) if self.id else ""}'

        # for k, v in self.__dict__.items():
        #     print(f'{k:<20s}{v}')

    def __generate_prot_item(self) -> dict:
        
        #TODO Refactor - dict: self.attrib ?
        
        item = {"name": self.name,
                "description": self.description,
                "owner_id": self.owner_id,
                "alternative_name": self.alternative_name,
                "gene": self.gene, 
                "species": self.species,
                "mutations": self.mutations,
                "chemical_modifications": self.chemical_modifications, 
                "tag": self.tag,
                "purification_method": self.purification_method,
                "mw": f'{self.mw}{" Da" if self.mw else ""}',
                "extinction_coefficient_280nm": self.extinction_coefficient_280nm,
                "storage_buffer": self.storage_buffer,
                "storage_temperature": self.storage_temp,
                "sequence": self.sequence,
        }
        
        # item = dict((k, v) for k, v in self.__dict__.items() if k in item_list and v is not None)
        return {k: v for k, v in item.items() if v is not None}

    def row_gen(self, ws, min_row: int, min_col: int, n_row=None,  n_col=None, i=0):
        """
        Generator returning index: i and row list: row
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
        Generator yielding row index: int and row data: list
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
                
    def create_new_record(self):
        
        url = "https://my.labguru.com/api/v1/proteins"
        body = {"token": self.token,
                "item": self.__generate_prot_item()}
   
        session = requests.post(url, json=body)
            
        if session.status_code == 201:
            try:
                response = session.json()
                self.id = response.get("id", None)
                self.uuid = response.get("uuid", None)
                self.sys_id = response.get("sys_id", None)
                if (url := response.get("url", None)):
                    self.url = f'https://my.labguru.com/{url}'
                else:
                    self.url = ""
                self.class_name = response.get("class_name", None)
                print(f'{self.sys_id:>10s} | {self.name:<50s} - New protein entry added')
            except Exception as e:
                print(e)   
        else:
            print(f'Error while handling {self.name} - Code {session.status_code}')

    def update_lg_record(self):
        url = f"https://my.labguru.com/api/v1/proteins/{self.id}"
        body = {"token": self.token,
                "item": self.__generate_prot_item()}
    
        session = requests.put(url, json=body)
        if session.status_code == 200:
            response = session.json()
            self.sys_id = response['sys_id']
            self.name = response['name']
            print(f'{self.sys_id:>10s} | {self.name:<30s} - Labguru protein record updated')
        else:
            print(f'Error while handling {self.name} - Code {session.status_code}')

    def get_stock_df(self, file):

        stock_df = pd.read_excel(file, self.ws_name, engine='openpyxl', nrows=101, index_col=0, usecols='D:M')
        return stock_df.dropna(how="all")
    
    # def get_stock_df(self, wb):
    #     min_row = 0
    #     n_row = 101
    #     min_col = 5
    #     n_col = 9
        
    #     ws = wb[self.ws_name]
        
    #     data = [[i.value for i in j] for j in ws.iter_rows(
    #         min_row=min_row, 
    #         min_col=min_col, 
    #         max_row=min_row+n_row,
    #         max_col=min_col+n_col)
    #         ]
    #     print(data)
    #     stock_df = pd.DataFrame(data=data[1:],
    #                             columns=data[0],
    #                             index=range(1,n_row)
    #                             )
    #     print(stock_df)
    #     return stock_df.dropna(how="all")

    def __generate_stock_items(self, stock_df: pd.DataFrame):
        """
        
        """
        
        def validate_stock(row):
            # Check stock row for required values
            try:
                box_id = row["Box ID"] is not None
                stock_name = row["Stock name"] is not None
                volume = row["Stock volume"] is not None
                conditions = [box_id, stock_name, volume]
                return all(conditions)
            except Exception as e:
                print(e)
                return False
        
        for index, row in stock_df.iterrows():
            row = row.fillna("")
            if validate_stock(row):
                if (stock_id := row["Stock ID"]):
                    stock_id = int(stock_id)
                try:
                    item = {"name": row["Stock name"],
                            "storage_id": int(row["Box ID"]),
                            "storage_type": "System::Storage::Box",
                            "stockable_type": "Biocollections::Protein",
                            "stockable_id": self.id,
                            "description": row["Description"],
                            # "barcode": "",
                            # "stored_by": "",
                            "concentration": str(round(row["Concentration"], 3)),
                            "concentration_prefix": "",
                            "concentration_unit_id": 9, # mg/mL
                            "concentration_exponent": "",
                            "concentration_remarks": "",
                            "volume": str(round(row["Stock volume"], 0)),
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
                    yield index, stock_id, item
                    
                except TypeError:
                    print(f'<<< File was not closed - error loading data of Stock ID {stock_id} >>>'.center(80))
                    print('--> Close the file and try again to update stocks')
                    continue
            else:
                continue

    def create_stocks(self, file):
                
        self.added_stocks = []
        
        stock_df = self.get_stock_df(file)
        stock_items = self.__generate_stock_items(stock_df)
        url = 'https://my.labguru.com/api/v1/stocks'
        
        for i, stock_id, item in stock_items:
            if not stock_id:
                body = {'token': self.token,
                        'item': item}

                box_id = item["storage_id"]
        
                session = requests.post(url, json=body)
                if session.status_code == 201:
                    response = session.json()
                    
                    stock_id = response['id']
                    position = response['position']
                    box_name = response['box']['name']

                    new_stock = {
                        'index': i,
                        'id': (stock_id, f'https://my.labguru.com/storage/stocks/{stock_id}'),
                        'box': (box_name, f'https://my.labguru.com/storage/boxes/{box_id}'),
                        'position': position, 
                    }
                    self.added_stocks.append(new_stock)
                    
                    print(f'\tStock {i} (ID: {new_stock["id"][0]:>6}) - Box: {new_stock["box"][0]}, Position: {new_stock["position"]}')
                else:
                        print(f'Error while handling {self.name}: Stock {i} - Code {session.status_code}')      

    def update_lg_stocks(self, file):
                                
            stock_df = self.get_stock_df(file)
            stock_items = self.__generate_stock_items(stock_df)

            for i, stock_id, item in stock_items:
                url = f'https://my.labguru.com/api/v1/stocks/{stock_id:}'
                body = {'token': self.token,
                        'item': item}
                session = requests.put(url, json=body)
                if session.status_code == 200:                  
                    print(f'\tStock {i} (ID: {stock_id:>6}) - stock attributes updated')
                else:
                    print(f'Error while handling {self.name}: Stock {i} - Code {session.status_code}')    

    def update_excel_record(wb):
        pass
        
    def update_excel_stocks(wb):
        pass

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

            if (lg_attrib := attrib_dict.get(attrib)):
                link_attribs = {
                    'sys_id': (self.sys_id, self.url),
                    'web_page': (self.uniprot_id, self.web_page),
                    }
                try:
                    if lg_attrib not in link_attribs:
                        new_value = prot_dict.get(lg_attrib, "")
                        ws.cell(row=i, column=2).value = new_value
                    else:
                        value, url = link_attribs.get(lg_attrib)
                        ws.cell(row=i, column=2).value = value
                        ws.cell(row=i, column=2).hyperlink = url
                except TypeError as e:
                    print(e)
        
    def update_sheet_stocks(self, wb):
        ws = wb[self.ws_name]

        header = self.get_stock_header(ws)
        
        for stock in self.added_stocks:
            row = stock['index'] + 1
            columns = ['Stock ID', 'Box name', 'Position']
            for name in columns:
                column = header[name]
                attrib = attrib_dict.get(name, None)
                if attrib is not None:
                    value = stock[attrib]
                    if isinstance(value, tuple):
                        value, hyperlink = value
                        ws.cell(row=row, column=column).value = value
                        ws.cell(row=row, column=column).hyperlink = hyperlink
                    else:
                        ws.cell(row=row, column=column).value = value

    def calc_params(self):

        self.len = len(self.sequence)
        self.aa_distr()
        if self.mw is None:
            self.prot_mass()
        if self.extinction_coefficient_280nm is None:
            self.abs_coeff()
        if self.iso_point is None:
            self.pI()
        
    def aa_distr(self):
        self.aa_count = {aa: self.sequence.count(aa) for aa in aa_dict}
        self.aa_proc = {aa: self.aa_count[aa] / self.len for aa in aa_dict}
        
    # Calculate protein mass
    def prot_mass(self):
        self.mw = round(sum(aa_dict.get(aa, 0)[2] for aa in self.sequence) + 18.01528, 2)

    # Estimate protein absorption coefficients
    def abs_coeff(self):
        base = 5500 * self.aa_count["W"] + 1490 * self.aa_count["Y"]
        self.extinction_red = round(base / self.mw, 3)
        self.extinction_ox = round((base + 125 * (self.aa_count["C"] // 2)) / self.mw, 3)
        self.extinction_coefficient_280nm = f'{self.extinction_ox} (Ox) / { self.extinction_red} (Red)'

    def pI(self):
        
        # Calculates charge of aa at given pH based on pKa
        def charge(ph, pka):
            return 1 / (10 ** (ph - pka) + 1)
        
        # Set range of pI search and start pH
        pH_min, pH, pH_max = 2, 8, 14.0
        
        base_pka = {'K': 10.0,
                    'R': 12.0,
                    'H': 5.98}
        acid_pka = {'D': 4.05,
                    'E': 4.45,
                    'C': 9.0,
                    'Y': 10.0}
        n_term_dict = {'A': 7.59, 
                       'M': 7.0, 
                       'S': 6.93, 
                       'P': 8.36, 
                       'T': 6.82, 
                       'V': 7.44, 
                       'E': 7.7}
        c_term_dict = {'D': 4.55, 
                       'E': 4.75}

        n_pka = n_term_dict.get(self.sequence[0], 7.7)
        c_pka = c_term_dict.get(self.sequence[-1], 3.55)
        
        p_charge = sum([charge(pH, base_pka[aa]) * self.aa_count[aa] for aa in base_pka.keys()]) + charge(pH, n_pka)
        n_charge = sum([charge(acid_pka[aa], pH) * self.aa_count[aa] for aa in acid_pka.keys()]) + charge(c_pka, pH)
        net_charge = p_charge - n_charge

        while abs(net_charge) > 0.0001:
            if  net_charge > 0:
                pH_min = pH
            else:
                pH_max = pH
            
            pH = (pH_min + pH_max) / 2
            
            p_charge = sum([charge(pH, base_pka[aa]) * self.aa_count[aa] for aa in base_pka.keys()]) + charge(pH, n_pka)
            n_charge = sum([charge(acid_pka[aa], pH) * self.aa_count[aa] for aa in acid_pka.keys()]) + charge(c_pka, pH)
            net_charge = p_charge - n_charge

        self.iso_point = pH