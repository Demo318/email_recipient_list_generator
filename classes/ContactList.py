import pandas


class ContactList:

    def __init__(self, category_name):
        self.category_name = category_name
        self.list_name = category_name + "_list"
        self.contacts_library = {
            'First': [], 'Last': [], 'Company': [],
            'Email': [], 'BCC': []
        }

    def add_contact(self, row):
        for index, item in enumerate(self.contacts_library):
            self.contacts_library[item].append(row[index])

    def dataframe(self):
        return pandas.DataFrame.from_dict(
            self.contacts_library
        )

    def filename_string(self):
        return r".\\contact_lists\\" + self.list_name + ".csv"
