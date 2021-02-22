# xl_file = pd.ExcelFile(file_name)

# dfs = {sheet_name: xl_file.parse(sheet_name) 
#           for sheet_name in xl_file.sheet_names}

# export to csv: https://stackoverflow.com/questions/53767845/how-do-i-export-a-pandas-dataframe-to-microsoft-access

# convert csv to mdb (for Word to read): https://stackoverflow.com/questions/3627469/how-to-create-a-mdb-file-from-a-csv-file-in-python

# Other good Pandas excel reference: https://www.journaldev.com/33306/pandas-read_excel-reading-excel-file-in-python

import pandas 


master_contact_list = pandas.read_excel('Master Contact List.xlsx', sheet_name='Sheet1')

cp_list = {'First':[], 'Last':[], 'Company':[], 'Email':[]}
it_list = {'First':[], 'Last':[], 'Company':[], 'Email':[]}
voice_list = {'First':[], 'Last':[], 'Company':[], 'Email':[]}
wf_and_pp_list = {'First':[], 'Last':[], 'Company':[], 'Email':[]}

def add_to_sub_contact_list(row, target_list):
    target_list['First'].append(row['First'])
    target_list['Last'].append(row['Last'])
    target_list['Company'].append(row['Company'])
    target_list['Email'].append(row['Email'])

for index, row in master_contact_list.iterrows():
    if row['CP'] == True:
        add_to_sub_contact_list(row, cp_list)
    if row['IT'] == True:
        add_to_sub_contact_list(row, it_list)
    if row['Voice'] == True:
        add_to_sub_contact_list(row, voice_list)
    if row['WF & PP'] == True:
        add_to_sub_contact_list(row, wf_and_pp_list)


    

df_cp_list = pandas.DataFrame.from_dict(cp_list)
df_it_list = pandas.DataFrame.from_dict(it_list)
df_voice_list = pandas.DataFrame.from_dict(voice_list)
df_wf_and_pp_list = pandas.DataFrame.from_dict(wf_and_pp_list)

df_cp_list.to_csv('.\contact_lists\cp_contact_list.csv', sep=',', encoding='utf=8')
df_it_list.to_csv('.\contact_lists\it_contact_list.csv', sep=',', encoding='utf=8')
df_voice_list.to_csv('.\contact_lists\\voice_contact_list.csv', sep=',', encoding='utf=8')
df_wf_and_pp_list.to_csv('.\contact_lists\wf_and_pp_contact_list.csv', sep=',', encoding='utf=8')



# Export .csv

# Convert .csv to .mdb