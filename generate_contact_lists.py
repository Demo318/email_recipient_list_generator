# xl_file = pd.ExcelFile(file_name)

# dfs = {sheet_name: xl_file.parse(sheet_name) 
#           for sheet_name in xl_file.sheet_names}

# export to csv: https://stackoverflow.com/questions/53767845/how-do-i-export-a-pandas-dataframe-to-microsoft-access

# convert csv to mdb (for Word to read): https://stackoverflow.com/questions/3627469/how-to-create-a-mdb-file-from-a-csv-file-in-python

# Other good Pandas excel reference: https://www.journaldev.com/33306/pandas-read_excel-reading-excel-file-in-python

import pandas 

print('lol')

master_contact_list = pandas.read_excel('Master Contact List.xlsx', sheet_name='Sheet1')

print(master_contact_list)