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

    def generate_prot_item(self) -> dict:
        
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
                "item": self.generate_prot_item()}

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
                "item": self.generate_prot_item()}
    
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