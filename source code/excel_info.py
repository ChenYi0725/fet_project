import re

class ExcelInfo():
    def __init__(self,node_name,tac,longitude,latitude,bands,table_widget):
        self.node_name = node_name
        self.tac = tac
        self.longitude = self._get_latitude_or_longitude(longitude)
        self.latitude = self._get_latitude_or_longitude(latitude)
        self.table_widget = table_widget
        self.table_info = self._get_table_info()
        self.test = self._get_display_cell_name("700")
        self.bands = bands
        self.e_node_b_id = self._get_e_node_b_id()
        self.export_map = {
            "Node Name": self.node_name,
            "eNodeB ID": self.e_node_b_id,
            "TAC": self.tac,
        }
        self.get_export_data()

    def __str__(self):
        return(
        f"Node Name: {str(self.node_name)}\n"
        f"TAC: {str(self.tac)}\n"
        f"Longitude: {str(self.longitude)}\n"
        f"Latitude: {str(self.latitude)}\n"
        f"Bands: {str(self.bands)}\n"
        f"Table Info: {str(self.table_info)}\n"
        f"eNodeB ID:{str(self.e_node_b_id)}\n"
        f"input col:{str(self._get_filled_columns())}\n"
        f"test:{self.test}\n")
    
    def is_input_correct(self):
        node_name_pattern = r'^[A-Za-z]{2}\d{6}$'
        if not re.match(node_name_pattern,self.node_name):
            return False
        try:
            float(self.tac)
            float(self.latitude)
            float(self.longitude)
            for i in self.table_info:
                for j in i:
                    if not j == "":  
                        float(j)
            return True
        except:
            return False

    def _get_e_node_b_id(self):
        return self.node_name[2:]
    
    def _get_latitude_or_longitude(self,input): #input = self.latitude or self.longitude
            if not input:
                 return ""
            else:
                try:
                    return str(int(input))
                except ValueError:
                    try:
                        float_val = float(input)
                        formatted = f"{float_val:.6f}"
                        return formatted.replace(".", "")
                    except ValueError:
                         return ""
    
    def _get_display_latitude_or_longitude(self,input):
        output = ""
        lat_long = self._get_latitude_or_longitude(input)
        for i in range(len(self._get_filled_columns())):
            output = output +lat_long+"/"
        return output[:-1]

    def _get_filled_columns(self, is_return_index=False): 
        filled_columns = []
        for col in range(self.table_widget.columnCount()):
            full = True
            for row in range(self.table_widget.rowCount()):
                item = self.table_widget.item(row, col)
                if not item or item.text().strip() == "":
                    full = False
                    break
            if full:
                if is_return_index:
                    filled_columns.append(col)
                else:
                    header_item = self.table_widget.horizontalHeaderItem(col)
                    if header_item:
                        filled_columns.append(header_item.text())
                    else:
                        filled_columns.append(f"Column {col}")  
        return filled_columns
    
    def _get_column_data(self, column_name):
        col_index = -1
        for col in range(self.table_widget.columnCount()):
            header_item = self.table_widget.horizontalHeaderItem(col)
            if header_item and header_item.text().strip() == column_name:
                col_index = col
                break
        if col_index == -1:
            raise ValueError(f"找不到欄位名稱: {column_name}")
        col_data = []
        for row in range(self.table_widget.rowCount()):
            item = self.table_widget.item(row, col_index)
            if item:
                col_data.append(item.text().strip())
            else:
                col_data.append("")  # 空格子補空字串
        return col_data
    
    def _get_table_info(self):
        table_info = []
        for row in range(self.table_widget.rowCount()):
            row_data = []
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row, col)
                if item:
                    row_data.append(item.text())
                else:
                    row_data.append("")
            table_info.append(row_data)
        return table_info
    
 
    
    def _get_display_cell_name(self,band,is2600D6=False):
        band_map = {
                "700":"1",
                "700 5M":"8",
                "1800":"2",
                "2100":"5",
                "2600":"3",
                "2600 TDD":"4"
            }
        display_cell_name = ""
        for i in self._get_filled_columns():
            if is2600D6:
                i = chr(ord(i)+10)
            display_cell_name=display_cell_name+self.node_name+i+band_map[band]+"/"
        display_cell_name = display_cell_name[:-1]
        return display_cell_name         

    def _get_cell_id(self,band):
        band_map = {
            "700": "x1",
            "700 5M": "x8",
            "1800": "x2",
            "2100": "x5",
            "2600": "x3",
            "2600 TDD D5": "x4",
            "2600 TDD D6": "1x4"
        }
        if band not in band_map:
            return ""
        base = band_map[band]
        filled_cols = self._get_filled_columns(is_return_index=True)
        cell_ids =""
        for i in filled_cols:
            cell_id = base.replace("x", str(i + 1))
            cell_ids = cell_ids + cell_id+"/"
        return cell_ids[:-1]
    
    def _get_display_pss(self):
        display_pss = ""
        for i in self._get_filled_columns():
            display_pss= display_pss+str(self._get_column_data(i)[0])+"/"
        display_pss = display_pss[:-1]
        return display_pss


    def _enter_display_sss_or_rach(self,is_sss,keys):     # k -> list ,input map、key, get sss from table
        filled_cols = self._get_filled_columns()
        for i, key in enumerate(keys):  # 依序對應keys
            if i < len(filled_cols):
                col_name = filled_cols[i]
                if is_sss:
                    self.export_map[key] = self._get_column_data(col_name)[1] # sss
                else:
                    self.export_map[key] = self._get_column_data(col_name)[2] # rach
            else:
                self.export_map[key] = ""  # 不足的
        
    def get_export_data(self):
        if "700" in self.bands:
            self.export_map["EUtranCellFDD ID (L700)"]=self._get_display_cell_name("700")
            self.export_map["Cell ID (L700)"]=self._get_cell_id("700")
            self.export_map["Latitude"] = self._get_display_latitude_or_longitude(self.latitude)
            self.export_map["Longitude"] = self._get_display_latitude_or_longitude(self.longitude)
            self.export_map["Physical Layer Cell ID Group"] =self._get_display_pss()
            sss_700_keys = ["Physical Layer Sub Cell ID 1","Physical Layer Sub Cell ID 2","Physical Layer Sub Cell ID 3"]
            self._enter_display_sss_or_rach(is_sss = True,keys = sss_700_keys)
            rach_700_keys =["RACH root sequence number 1","RACH root sequence number 2","RACH root sequence number 3"]
            self._enter_display_sss_or_rach(is_sss = False,keys = rach_700_keys)
        if "700 5M" in self.bands:
            self.export_map["APT EUtranCellFDD ID (L700)"] = self._get_display_cell_name("700 5M")
            self.export_map["APT Cell ID (L700)"] =self._get_cell_id("700 5M")
        if "1800" in self.bands:
            self.export_map["EUtranCellFDD ID (L1800)"] = self._get_display_cell_name("1800")
            self.export_map["Cell ID (L1800)"] = self._get_cell_id("1800")
        if "2100" in self.bands:
            self.export_map["Latitude (L2100)"] = self._get_display_latitude_or_longitude(self.latitude)
            self.export_map["Longitude (L2100)"] = self._get_display_latitude_or_longitude(self.longitude)
            self.export_map["EUtranCellFDD ID (L2100)"] =self._get_display_cell_name("2100")
            self.export_map["Cell ID (L2100)"] =self._get_cell_id("2100")
            self.export_map["Physical Layer Cell ID Group (L2100)"] =self._get_display_pss()
            sss_2100_keys = ["Physical Layer Sub Cell ID 1 (L2100)","Physical Layer Sub Cell ID 2 (L2100)","Physical Layer Sub Cell ID 3 (L2100)","Physical Layer Sub Cell ID 4 (L2100)","Physical Layer Sub Cell ID 5 (L2100)","Physical Layer Sub Cell ID 6 (L2100)"]
            self._enter_display_sss_or_rach(is_sss=True,keys=sss_2100_keys)
            rach_2100_keys = ["RACH root sequence number 1 (L2100)","RACH root sequence number 2 (L2100)","RACH root sequence number 3 (L2100)","RACH root sequence number 4 (L2100)","RACH root sequence number 5 (L2100)","RACH root sequence number 6 (L2100)",]
            self._enter_display_sss_or_rach(is_sss=False,keys=rach_2100_keys)
        if "2600" in self.bands:
            self.export_map["Latitude (L2600)"] = self._get_display_latitude_or_longitude(self.latitude)
            self.export_map["Longitude (L2600)"] = self._get_display_latitude_or_longitude(self.longitude)
            self.export_map["EUtranCellFDD ID (L2600)"] = self._get_display_cell_name("2600")
            self.export_map["Cell ID (L2600)"] = self._get_cell_id("2600")
            sss_2600_keys =["Physical Layer Cell ID Group (L2600)","Physical Layer Sub Cell ID 1 (L2600)","Physical Layer Sub Cell ID 2 (L2600)","Physical Layer Sub Cell ID 3 (L2600)","Physical Layer Sub Cell ID 4 (L2600)","Physical Layer Sub Cell ID 5 (L2600)","Physical Layer Sub Cell ID 6 (L2600)",]
            self._enter_display_sss_or_rach(is_sss=True,keys=sss_2600_keys)
            rach_2600_keys =["RACH root sequence number 1 (L2600)","RACH root sequence number 2 (L2600)","RACH root sequence number 3 (L2600)"]
            self._enter_display_sss_or_rach(is_sss=False,keys=rach_2600_keys)
        if "2600 TDD" in self.bands:
            self.export_map["Latitude (TDD_L2600)_D5"] = self._get_display_latitude_or_longitude(self.latitude)
            self.export_map["Longitude (TDD_L2600)_D5"] = self._get_display_latitude_or_longitude(self.longitude)
            self.export_map["Latitude (TDD_L2600)_D6"] = self._get_display_latitude_or_longitude(self.latitude)
            self.export_map["Longitude (TDD_L2600)_D6"] = self._get_display_latitude_or_longitude(self.longitude)
            self.export_map["EUtranCellTDD ID (TDD_L2600)_D5"] = self._get_display_cell_name("2600 TDD")
            self.export_map["Cell ID (TDD_L2600)_D5"] = self._get_cell_id("2600 TDD D5")
            self.export_map["Physical Layer Cell ID Group (TDD_L2600)_D5"] = self._get_display_pss()
            sss_tdd_2600_keys = ["Physical Layer Sub Cell ID 1 (TDD_L2600)_D5","Physical Layer Sub Cell ID 2 (TDD_L2600)_D5","Physical Layer Sub Cell ID 3 (TDD_L2600)_D5","Physical Layer Sub Cell ID 4 (TDD_L2600)_D5","Physical Layer Sub Cell ID 5 (TDD_L2600)_D5","Physical Layer Sub Cell ID 6 (TDD_L2600)_D5",]
            self._enter_display_sss_or_rach(is_sss=True,keys=sss_tdd_2600_keys)
            rach_tdd_2600_keys =["RACH root sequence number 1 (TDD_L2600)_D5","RACH root sequence number 2 (TDD_L2600)_D5","RACH root sequence number 3 (TDD_L2600)_D5"]
            self._enter_display_sss_or_rach(is_sss=False,keys=rach_tdd_2600_keys)
            self.export_map["EUtranCellTDD ID (TDD_L2600)_D6"] = self._get_display_cell_name("2600 TDD",is2600D6=True)
            self.export_map["Cell ID (TDD_L2600)_D6"] = self._get_cell_id("2600 TDD D6")
            self.export_map["Physical Layer Cell ID Group (TDD_L2600)_D6"] = self._get_display_pss()
            sss_tdd_2600_d6_keys =["Physical Layer Sub Cell ID 1 (TDD_L2600)_D6","Physical Layer Sub Cell ID 2 (TDD_L2600)_D6","Physical Layer Sub Cell ID 3 (TDD_L2600)_D6","Physical Layer Sub Cell ID 4 (TDD_L2600)_D6","Physical Layer Sub Cell ID 5 (TDD_L2600)_D6","Physical Layer Sub Cell ID 6 (TDD_L2600)_D6",]
            self._enter_display_sss_or_rach(is_sss=True,keys=sss_tdd_2600_d6_keys)
            rach_tdd_2600_d6_keys =["RACH root sequence number 1 (TDD_L2600)_D6","RACH root sequence number 2 (TDD_L2600)_D6","RACH root sequence number 3 (TDD_L2600)_D6","RACH root sequence number 4 (TDD_L2600)_D6","RACH root sequence number 5 (TDD_L2600)_D6","RACH root sequence number 6 (TDD_L2600)_D6",]
            self._enter_display_sss_or_rach(is_sss=False,keys=rach_tdd_2600_d6_keys)
        print(f"\n\n\n\nexport data:{self.export_map}\n\n\n")

        




# output_map = {
#     "Node Name": "",
#     "eNodeB ID": "",
#     "TAC": "",
#     # ------
#     "EUtranCellFDD ID (L700)": "",
#     "Cell ID (L700)": "",
#     "Physical Layer Cell ID Group": "",
#     "Physical Layer Sub Cell ID 1": "",
#     "Physical Layer Sub Cell ID 2": "",
#     "Physical Layer Sub Cell ID 3": "",
#     "Latitude": "",
#     "Longitude": "",
#     "RACH root sequence number 1": "",
#     "RACH root sequence number 2": "",
#     "RACH root sequence number 3": "",
#     # -------
#     "APT EUtranCellFDD ID (L700)": "",
#     "APT Cell ID (L700)": "",
#     # -------
#     "EUtranCellFDD ID (L1800)": "",
#     "Cell ID (L1800)": "",
#     # -------
#     "EUtranCellFDD ID (L2600)":"",
#     "Cell ID (L2600)":"",
#     "Physical Layer Cell ID Group (L2600)":"",
#     "Physical Layer Sub Cell ID 1 (L2600)":"",
#     "Physical Layer Sub Cell ID 2 (L2600)":"",
#     "Physical Layer Sub Cell ID 3 (L2600)":"",
#     "Physical Layer Sub Cell ID 4 (L2600)":"",
#     "Physical Layer Sub Cell ID 5 (L2600)":"",
#     "Physical Layer Sub Cell ID 6 (L2600)":"",
#     "Latitude (L2600)":"",
#     "Longitude (L2600)":"",
#     "RACH root sequence number 1 (L2600)":"",
#     "RACH root sequence number 2 (L2600)":"",
#     "RACH root sequence number 3 (L2600)":"",
#     # --------
#     "EUtranCellTDD ID (TDD_L2600)_D5":"",
#     "Cell ID (TDD_L2600)_D5":"",
#     "Physical Layer Cell ID Group (TDD_L2600)_D5":"",
#     "Physical Layer Sub Cell ID 1 (TDD_L2600)_D5":"",
#     "Physical Layer Sub Cell ID 2 (TDD_L2600)_D5":"",
#     "Physical Layer Sub Cell ID 3 (TDD_L2600)_D5":"",
#     "Physical Layer Sub Cell ID 4 (TDD_L2600)_D5":"",
#     "Physical Layer Sub Cell ID 5 (TDD_L2600)_D5":"",
#     "Physical Layer Sub Cell ID 6 (TDD_L2600)_D5":"",
#     "Latitude (TDD_L2600)_D5":"",
#     "Longitude (TDD_L2600)_D5":"",
#     "RACH root sequence number 1 (TDD_L2600)_D5":"",
#     "RACH root sequence number 2 (TDD_L2600)_D5":"",
#     "RACH root sequence number 3 (TDD_L2600)_D5":"",
#     "EUtranCellTDD ID (TDD_L2600)_D6":"",
#     "Cell ID (TDD_L2600)_D6":"",
#     "Physical Layer Cell ID Group (TDD_L2600)_D6":"",
#     "Physical Layer Sub Cell ID 1 (TDD_L2600)_D6":"",
#     "Physical Layer Sub Cell ID 2 (TDD_L2600)_D6":"",
#     "Physical Layer Sub Cell ID 3 (TDD_L2600)_D6":"",
#     "Physical Layer Sub Cell ID 4 (TDD_L2600)_D6":"",
#     "Physical Layer Sub Cell ID 5 (TDD_L2600)_D6":"",
#     "Physical Layer Sub Cell ID 6 (TDD_L2600)_D6":"",
#     "Latitude (TDD_L2600)_D6":"",
#     "Longitude (TDD_L2600)_D6":"",
#     "RACH root sequence number 1 (TDD_L2600)_D6":"",
#     "RACH root sequence number 2 (TDD_L2600)_D6":"",
#     "RACH root sequence number 3 (TDD_L2600)_D6":"",
#     "RACH root sequence number 4 (TDD_L2600)_D6":"",
#     "RACH root sequence number 5 (TDD_L2600)_D6":"",
#     "RACH root sequence number 6 (TDD_L2600)_D6":"",
#     # ------
#     "EUtranCellFDD ID (L2100)":"",
#     "Cell ID (L2100)":"",
#     "Physical Layer Cell ID Group (L2100)":"",
#     "Physical Layer Sub Cell ID 1 (L2100)":"",
#     "Physical Layer Sub Cell ID 2 (L2100)":"",
#     "Physical Layer Sub Cell ID 3 (L2100)":"",
#     "Physical Layer Sub Cell ID 4 (L2100)":"",
#     "Physical Layer Sub Cell ID 5 (L2100)":"",
#     "Physical Layer Sub Cell ID 6 (L2100)":"",
#     "Latitude (L2100)":"",
#     "Longitude (L2100)":"",
#     "RACH root sequence number 1 (L2100)":"",
#     "RACH root sequence number 2 (L2100)":"",
#     "RACH root sequence number 3 (L2100)":"",
#     "RACH root sequence number 4 (L2100)":"",
#     "RACH root sequence number 5 (L2100)":"",
#     "RACH root sequence number 6 (L2100)":"",
# }   
        