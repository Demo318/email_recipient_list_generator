# A process to take one master list of contacts and
# generate sub-lists in individal CSV files to be used
# with Microsoft Word's mail merge feature.

import pandas
from classes.ContactList import ContactList


master_contact_list = pandas.read_excel(
    'Master Contact List.xlsx',
    sheet_name='Sheet1'
)

contact_lists = {}

for column_name in master_contact_list.columns[5:]:
    contact_lists[column_name] = ContactList(column_name)


for index, row in master_contact_list.iterrows():
    for sub_index, item in enumerate(row[5:]):
        if item is True:
            contact_lists[
                master_contact_list.columns[sub_index + 5]
            ].add_contact(row)


for list_name in contact_lists:
    d_frame = contact_lists[list_name].dataframe()
    d_frame.to_csv(
        contact_lists[list_name].filename_string(),
        sep=',',
        encoding='utf=8'
    )


# TODO: Generate master email:BCC list
