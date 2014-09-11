import xlrd
#import configparser
import xlsxwriter

class FamilyPack():
    def __init__(self):
        self.host = ''
        self.address = ''
        self.amount = 0
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
        family_dict = {}


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
                            
            if address_value not in family_dict.keys():
                new_family = FamilyPack()
                new_family.address = address_value
                new_family.host = name_value
                new_family.amount = 1
                new_family.phone = phone_value
                family_dict[address_value] = new_family
            else:
                current_family = family_dict[address_value]
                current_family.amount += 1
                
                if age_value >= 50 and age_value < 60:
                    current_family.age_between_50_60 += 1
                elif age_value >= 60:
                    current_family.age_above_60 += 1
        
        for family in family_dict.values():
            print(family.host)
        
        return family_dict

def xlsx_write(family_dict):
    workbook = xlsxwriter.Workbook('output.xlsx')
    sheet = workbook.add_worksheet()
    
    title_list = ['Host', 'Address', 'Amount', 'Age<60', 'Age>60', 'Phone']
    for column_index in range(len(title_list)):
        sheet.write(0, column_index, title_list[column_index])
        
    row_index = 0
    for family in family_dict.values():
        row_index += 1
        sheet.write(row_index, 0, family.host)
        sheet.write(row_index, 1, family.address)
        sheet.write(row_index, 2, family.amount)
        sheet.write(row_index, 3, family.age_between_50_60)
        sheet.write(row_index, 4, family.age_above_60)
        sheet.write(row_index, 5, family.phone)
    
    workbook.close()     
#Main starts from here
xlsx_file = "1.xls"
log_list = []
family_dict = {}

family_dict = read_excel(xlsx_file)
xlsx_write(family_dict)