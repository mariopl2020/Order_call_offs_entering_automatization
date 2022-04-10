from files_manager import CallOffFileProcessing

class CallOff():
	"""Represents data from call off template file"""

	def __init__(self):
		"""Initialization of CallOff object"""

		self.filename = r"calloff_template.xlsx"
		self.file_processor = CallOffFileProcessing()
		self.call_off_data_frame = None

	def get_data_frame(self):
		"""Getting completely processed dataframe starting from raw xlsx file"""

		self.call_off_data_frame = self.file_processor.give_complete_dataframe(self.filename)

# @TODO ABSTRACT CLASS for call_off and orders
