from files_manager import CallOffFileProcessing

class CallOff():
	"""Represents data from call off template file"""

	def __init__(self):
		"""Initialization of CallOff object"""

		self.filename = r"calloff_template.xlsx"
		self.file_processor = CallOffFileProcessing()
		self.call_off_data_frame = None
		self.current_order = None
		self.current_order_no = None

	def get_data_frame(self):
		"""Getting completely processed dataframe starting from raw xlsx file"""

		self.call_off_data_frame = self.file_processor.give_complete_dataframe(self.filename)

	def give_single_order(self, row_index):
		"""Gives complete row from dataframe as Series object and order number

		Args:
			row_index (int): number of row index in dataframe"""

		self.current_order = self.call_off_data_frame.iloc[row_index]
		self.current_order_no = self.call_off_data_frame.iloc[row_index]["Customer_PO_number"]


# @TODO ABSTRACT CLASS for call_off and orders
