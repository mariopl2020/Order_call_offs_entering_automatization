from files_manager import FilesProcessing
from orders import Orders


class Program():
	"""Represents complete form of program"""

	def __init__(self):
		"""Initialization of program"""

		self.order = Orders()
		self.call_off = self.order.call_off
		self.files_processing = FilesProcessing()

	def main_run(self):
		"""Collects all key method calls to run program completely"""

		self.call_off.get_data_frame()
		self.order.get_data_frame()

	def main_loop(self):
		"""Iterates on single order line from call off and calls main methods to process it"""

		for index in range(len(self.call_off.call_off_data_frame["Customer_PO_number"])):
			self.call_off.give_single_order(index)
			self.order.match_order_from_call_off()
			self.order.create_working_dataframe()
			self.order.process_row_to_be_delivered()
			self.order.process_row_with_remaining_quantity()
			self.order.change_order_part_postfix()
			self.order.calculate_confirmed_part_row_index()
			self.order.assign_row_indexes()
			self.order.add_working_rows_to_orders_dataframe()

if __name__ == "__main__":
	program = Program()
	program.main_run()
	program.main_loop()
	program.files_processing.write_dataframe_to_excel(program.order.orders_data_frame, "processed_orders.xlsx")


