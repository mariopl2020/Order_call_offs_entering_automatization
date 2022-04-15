"""Module to manage imported and exported files"""
import pandas as pd
#import openpyxl


class FilesProcessing():
    """Represents object what includes methods responsible for xlsx file processing to
    required dataframe shape"""

    def __init__(self):
        """Initialization of FileProcessing object"""

        self.data_frame = None

    def create_dataframe_from_file(self, filename):
        """Creates dataframe object from xlsx file

        Args:
            filename (str): name of file to be processed by program"""

        filepath = r'./data/' + filename
        dataframe_from_file = pd.read_excel(io=filepath)
        self.data_frame = dataframe_from_file

    def clean_columns_names(self):
        """Cleans dataframe's columns from whitespaces and instead of it gives underscore"""

        dataframe_column_names = [column.replace(" ", "_") for column in self.data_frame.columns]
        self.data_frame.columns = dataframe_column_names

    def give_complete_dataframe(self, filename):
        """Groups all key processing methods to get complete dataframe shape from file form"""

        self.create_dataframe_from_file(filename)
        self.clean_columns_names()

        return self.data_frame

    @staticmethod
    def write_dataframe_to_excel(dataframe, filename):
        """Writes provided dataframe into file 'output.xlsx'

        Args:
            dataframe (Dataframe): dataframe to be saved in xlsx file
            filename (str): name of file to be written"""

        filepath = r"./data/" + filename
        dataframe.to_excel(excel_writer=filepath, index=False)


class CallOffFileProcessing(FilesProcessing):
    """Represents object what includes methods responsible for xlsx file
    processing with call off data"""

    def remove_empty_lines(self):
        """Removes empty rows from call off dataframe"""

        mask = self.data_frame["SKU"].notnull()
        self.data_frame = self.data_frame[mask]

    def give_complete_dataframe(self, filename):
        """Groups all key processing methods to get complete call off dataframe shape from
        file form"""

        super().give_complete_dataframe(filename)
        self.remove_empty_lines()

        return self.data_frame
