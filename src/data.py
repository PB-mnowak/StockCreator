attrib_dict = {
    'POI name': 'name',
    'POI ID': 'id',
    'POI SysID': 'sys_id',
    'Link': 'url',
    'Format': 'format',
    'Tag': 'tag',
    'Length': 'len',
    'MW': 'mw',
    'pI': 'iso_point',
    'A0.1% (Ox)': 'extinction_ox',
    'A0.1% (Red)': 'extinction_red',
    'Concentration': 'concentration',
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
    'Sequence': 'sequence',
    'UniProt ID': 'web_page',
}

aa_dict = {
    'A': ('Ala', 'Alanine', 71.0788), 
    'R': ('Arg', 'Arginine', 156.1875), 
    'N': ('Asn', 'Asparagine', 114.1038), 
    'D': ('Asp', 'Aspartic Acid', 115.0886), 
    'C': ('Cys', 'Cysteine', 103.1388), 
    'Q': ('Gln', 'Glutamine', 128.1307), 
    'E': ('Glu', 'Glutamic Acid', 129.1155), 
    'G': ('Gly', 'Glycine', 57.0519), 
    'H': ('His', 'Histidine', 137.1411), 
    'I': ('Ile', 'Isoleucine', 113.1594), 
    'L': ('Leu', 'Leucine', 113.1594), 
    'K': ('Lys', 'Lysine', 128.1741), 
    'M': ('Met', 'Methionine', 131.1926), 
    'F': ('Phe', 'Phenylalanine', 147.1766), 
    'P': ('Pro', 'Proline', 97.1167), 
    'S': ('Ser', 'Serine', 87.0782), 
    'T': ('Thr', 'Threonine', 101.1051), 
    'W': ('Trp', 'Tryptophan', 186.2132), 
    'Y': ('Tyr', 'Tyrosine', 163.176), 
    'V': ('Val', 'Valine', 99.1326)
    }

prot_attribs = ["ws_name",
                "name",
                "id",
                "sys_id",
                "owner_id",
                "alternative_name",
                "gene",
                "species",
                "mutations",
                "chemical_modifications",
                "purification_method",
                "mw",
                "iso_point",
                "extinction_ox",
                "extinction_red",
                "storage_buffer",
                "storage_temperature",
                "sequence",
                "web_page",]

        # self.ws_name = data.get("ws_name", None)
        # self.name = data.get("name", None)
        # self.id = data.get("id", None)
        # self.sys_id = data.get("sys_id", None)
        # self.description = data.get("description", "")
        # self.owner_id = data.get("owner_id", None)
        # self.alternative_name = data.get("alternative_name", None)
        # self.gene = data.get("gene", None)
        # self.species = data.get("species", None)
        # self.mutations = data.get("mutations", None)
        # self.chemical_modifications = data.get("chemical_modifications", None)
        # self.tag = data.get("tag", None)
        # self.purification_method = data.get("purification_method", None)
        # self.mw = data.get("mw", None)
        # self.iso_point = data.get("iso_point", None)
        # self.extinction_ox = data.get("extinction_ox", None)
        # self.extinction_red = data.get("extinction_red", None)
        # # TODO refactor:
        # if self.extinction_ox is not None and self.extinction_red is not None:
        #     self.extinction_coefficient_280nm = f'{self.extinction_ox} (Ox) / {self.extinction_red} (Red)'
        # else:
        #     self.extinction_coefficient_280nm = None
        # self.storage_buffer = data.get("storage_buffer", None)
        # self.storage_temperature = data.get("storage_temperature", None)
        # sequence = Seq(data.get("sequence", ""))
        # self.sequence = str(sequence)
        # if (uniprot_id := data.get("web_page", "")):
        #     self.uniprot_id = uniprot_id
        #     self.web_page = f"https://www.uniprot.org/uniprot/{uniprot_id}"
        # else:
        #     self.uniprot_id, self.web_page = None, None
        # self.token = token
        # if self.id is not None:
        #     self.url = f'https://my.labguru.com/biocollections/proteins/{self.id}'
        # else:
        #     self.url = ''