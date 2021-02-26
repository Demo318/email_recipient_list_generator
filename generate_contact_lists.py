# xl_file = pd.ExcelFile(file_name)

# dfs = {sheet_name: xl_file.parse(sheet_name) 
#           for sheet_name in xl_file.sheet_names}

# export to csv: https://stackoverflow.com/questions/53767845/how-do-i-export-a-pandas-dataframe-to-microsoft-access

# convert csv to mdb (for Word to read): https://stackoverflow.com/questions/3627469/how-to-create-a-mdb-file-from-a-csv-file-in-python

# Other good Pandas excel reference: https://www.journaldev.com/33306/pandas-read_excel-reading-excel-file-in-python

import pandas
from classes.ContactList import ContactList


master_contact_list = pandas.read_excel('Master Contact List.xlsx', sheet_name='Sheet1')

contact_lists = {}

for column_name in master_contact_list.columns[5:]:
    contact_lists[column_name] = ContactList(column_name)


for index, row in master_contact_list.iterrows():
    for sub_index, item in enumerate(row[5:]):
         if item == True:
            contact_lists[master_contact_list.columns[sub_index + 5]].add_contact(row)


for list_name in contact_lists:
    d_frame = contact_lists[list_name].dataframe()
    d_frame.to_csv(
        contact_lists[list_name].filename_string(),
        sep=',',
        encoding='utf=8'
        )

# TODO: Generate master email:BCC list