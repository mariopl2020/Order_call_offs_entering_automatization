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
		self.fully_taken_order = None

	def get_data_frame(self):
		"""Getting completely processed dataframe starting from raw xlsx file"""

		try:
			self.call_off_data_frame = self.file_processor.give_complete_dataframe(self.filename)
		except FileNotFoundError:
			raise FileNotFoundError("File calloff_template.xlsx does not exist. Provide it to run program")

	def give_single_order(self, row_index):
		"""Gives complete row from dataframe as Series object, order number and infromation if order will be fully taken\
		or partially

		Args:
			row_index (int): number of row index in dataframe"""

		self.current_order = self.call_off_data_frame.iloc[row_index]
		self.current_order_no = self.call_off_data_frame.iloc[row_index]["Customer_PO_number"]
		self.fully_taken_order = self.call_off_data_frame.iloc[row_index]["Remarks"]

