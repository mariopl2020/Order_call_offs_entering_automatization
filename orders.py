import pandas as pd
from call_off import  CallOff
from files_manager import FilesProcessing

class Orders():
	"""Represents data from orders file"""

	def __init__(self):
		"""Initialization of Orders object"""

		self.filename = r"orders.xlsx"
		self.call_off = CallOff()
		self.data_processor = FilesProcessing()
		self.orders_data_frame = None
		self.current_order = None
		self.current_order_index = None
		self.temporary_orders_to_process = None

	def get_data_frame(self):
		"""Getting completely processed dataframe starting from raw xlsx file"""

		self.orders_data_frame = self.data_processor.give_complete_dataframe(self.filename)

	def match_order_from_call_off(self):
		"""Finds row in order dataframe corresponding to order number from call off. Write it down and its index"""

		mask = (self.orders_data_frame["Purchase_order_number"] == self.call_off.current_order_no)
		self.current_order = self.orders_data_frame[mask]
		self.current_order_index = self.current_order.index

	def create_working_dataframe(self):
		"""Deletes order row from order dataframe and double it into working dataframe. It starts splitting order
		into rows what will be delivered and remaining part of order still to be called off"""

		self.orders_data_frame = self.orders_data_frame.drop(self.current_order_index)
		self.temporary_orders_to_process = pd.concat(objs=[self.current_order, self.current_order], ignore_index=True)
		self.temporary_orders_to_process["Delivered_(kg)"] = self.temporary_orders_to_process["Delivered_(kg)"].astype("int32")




