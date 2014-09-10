import xlrd
import configparser

class FamilyPack():
    def __init__(self):
        self.host = ''
        self.address = ''
        self.amount = 0
        self.age_below_50 = 0
        self.age_between_50_60 = 0
        self.age_above_60 = 0        
        self.phone = ''

def LAP(string, log_list):
    print(string)
    log_list.append(string)

'''
def read_config_ids():
    try:
        f = open("config.ini", "r")
    except IOError:
        LAP("Config file missing!")
        return False
    else:
        f.close();
        config.read("config.ini")
'''
def read_excel(xlsx_file):
    try:
        workbook = xlrd.open_workbook(xlsx_file)
    except IOError:
        LAP("Target file missing!". log_list)
        return -1
    else:
        family_list = {}

        for sheet_index in range(workbook.nsheets):
            sheet = workbook.sheet_by_index(0)

            address_column_index = 6
            age_column_index = 4
            name_column_index = 1
            phone_column_index = 8
            
            for row_index in range(1, sheet.nrows):
                address_value = str(sheet.cell(row_index, address_column_index).value)
                age_value = int(sheet.cell(row_index, age_column_index).value)
                name_value = str(sheet.cell(row_index, name_column_index).value)
                phone_value = str(sheet.cell(row_index, phone_column_index).value)
                                
                if address_value not in family_list:
                    new_family = FamilyPack()
                    new_family.address = address_value
                    new_family.host = name_value
                    new_family.amount = 1
                    new_family.phone = phone_value
                    family_list[address_value] = new_family
                else:
                    current_family = family_list[address_value]
                    current_family.amount += 1
                    
                    if age_value < 50:
                        current_family.age_below_50 += 1
                    elif age_value >= 50 and age_value < 60:
                        current_family.age_between_50_60 += 1
                    elif age_value >= 60:
                        current_family.age_above_60 += 1
                    else:
                        LAP("Age is invalid!", log_list)
        
        
        for family in family_list.values():
            print(family.address)
            print(family.host)            
            print(family.amount)
            print(family.phone)    

#Main starts from here
xlsx_file = "1.xls"
log_list = []

read_excel(xlsx_file)

                         
                            
                                         
            
        
    
