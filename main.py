from files_manager import FilesProcessing
from orders import Orders


class Program():
	"""Represents complete form of program"""

	def __init__(self):
		"""Initialization of program"""

		self.order = Orders()
		self.call_off = self.order.call_off
		self.files_processing = FilesProcessing()

	def main_loop(self):
		"""Iterates on single order line from call off and calls main methods to process it"""

		for index in range(len(self.call_off.call_off_data_frame["Customer_PO_number"])):
			self.call_off.give_single_order(index)
			self.order.match_order_from_call_off()
			self.order.create_working_dataframe()

	def main_run(self):
		"""Collects all key method calls to run program completely"""

		self.call_off.get_data_frame()
		self.order.get_data_frame()


if __name__ == "__main__":
	program = Program()
	program.main_run()
	program.main_loop()
	program.files_processing.write_dataframe_to_excel(program.order.orders_data_frame, "output1.xlsx")
	program.files_processing.write_dataframe_to_excel(program.order.temporary_orders_to_process, "output2.xlsx")

