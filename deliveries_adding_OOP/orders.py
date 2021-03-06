"""Module to manage orders dataframe"""
import pandas as pd
from call_off import CallOff
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
        self.confirmed_part_row_index = None
        self.working_orders_dataframe = None

    def get_data_frame(self):
        """Getting completely processed dataframe starting from raw xlsx file"""

        try:
            self.orders_data_frame = self.data_processor.give_complete_dataframe(self.filename)
        except FileNotFoundError as orders_file_error:
            raise FileNotFoundError("File orders.xlsx does not exist. Provide it to"
                                    " run program") from orders_file_error

    def match_order_from_call_off(self):
        """Finds row in order dataframe corresponding to order number from call off.
        Write it down and its index"""

        mask = (self.orders_data_frame["Purchase_order_number"] == self.call_off.current_order_no)
        self.current_order = self.orders_data_frame[mask]
        self.current_order_index = self.current_order.index

    def create_working_dataframe(self):
        """Deletes order row from order dataframe. Checks if order will be fully or
        partially taken. Then decides to split it into 2 rows (to deliver and remaining
        part of order still to be called off) or 1 only to be delivered. Next they/it
        is assigned to working date frame
        """

        self.orders_data_frame = self.orders_data_frame.drop(self.current_order_index)
        if self.call_off.fully_taken_order == "complete":
            self.working_orders_dataframe = self.current_order.reset_index(drop=True)
        else:
            self.working_orders_dataframe = pd.concat(objs=[self.current_order, self.current_order],
                                                      ignore_index=True)
        self.working_orders_dataframe.loc[:, "Delivered_(kg)"] = \
            self.working_orders_dataframe["Delivered_(kg)"].astype("int32")

    def process_row_to_be_delivered(self):
        """Writes down values from order from call off dataframe as part of order to be delivered"""

        self.working_orders_dataframe.loc[0, "Current_delivery_date"] = \
            self.call_off.current_order["Confirmed_delivery_date"]
        self.working_orders_dataframe.loc[0, "Status"] = "confirmed"
        self.working_orders_dataframe.loc[0, "Required_delivery_date_OTIF1"] = \
            self.call_off.current_order["Requested_date"]
        self.working_orders_dataframe.loc[0, "Required_quantity_OTIF1"] = \
            self.call_off.current_order["Called_quantity_[kg]"]
        self.working_orders_dataframe.loc[0, "Ordered_(kg)"] = \
            self.call_off.current_order["Confirmed_quantity_[kg]"]
        self.working_orders_dataframe.loc[0, "Delivered_(kg)"] = \
            int(self.call_off.current_order["Confirmed_quantity_[kg]"])

    def process_row_with_remaining_quantity(self):
        """Corrects quantity values for row what indicates still remaining order's part"""

        self.working_orders_dataframe.loc[1, "Ordered_(kg)"] -= \
            self.working_orders_dataframe.loc[0, "Ordered_(kg)"]
        self.working_orders_dataframe.loc[1, "Delivered_(kg)"] = \
            self.working_orders_dataframe.loc[1, "Delivered_(kg)"] \
            - self.working_orders_dataframe.loc[0, "Delivered_(kg)"]
        if self.working_orders_dataframe.loc[1, "Delivered_(kg)"] < 0:
            self.working_orders_dataframe.loc[1, "Delivered_(kg)"] = 0

    def change_order_part_postfix(self):
        """Changes order number's postfixes what indicates order part's number"""

        if "/" in self.working_orders_dataframe.loc[1, "Purchase_order_number"]:
            labour_list = self.working_orders_dataframe.loc[1, "Purchase_order_number"].split("/")
            labour_list[1] = str(int(labour_list[1]) + 1)
            self.working_orders_dataframe.loc[1, "Purchase_order_number"] = \
                labour_list[0] + "/" + labour_list[1]
        else:
            self.working_orders_dataframe.loc[0, "Purchase_order_number"] += "/1"
            self.working_orders_dataframe.loc[1, "Purchase_order_number"] += "/2"

    def calculate_confirmed_part_row_index(self):
        """Calculates row index for new confirmed order's part based on sorted delivery dates
        in order's main dataframe """

        confirmed_part_delivery_date = self.working_orders_dataframe.loc[0, "Current_delivery_date"]
        mask = self.orders_data_frame["Current_delivery_date"] >= confirmed_part_delivery_date
        orders_with_equal_later_date = self.orders_data_frame[mask]
        self.confirmed_part_row_index = orders_with_equal_later_date.index[0] - 0.5

    def assign_row_indexes(self):
        """Assigns row indexes to part of orders """

        if self.call_off.fully_taken_order == "complete":
            self.working_orders_dataframe.index = [self.confirmed_part_row_index]
        else:
            self.working_orders_dataframe.index = [self.confirmed_part_row_index,
                                                   self.current_order_index[0]]

    def add_working_rows_to_orders_dataframe(self):
        """Rebuilds base orders dataframe by adding newly calculated order part's rows,
        sorts them by indexes and drops them"""

        self.orders_data_frame = pd.concat(objs=[self.orders_data_frame,
                                                 self.working_orders_dataframe])
        self.orders_data_frame = self.orders_data_frame.sort_index()
        self.orders_data_frame = self.orders_data_frame.reset_index(drop=True)

    def back_to_original_colnames(self):
        """Allows back to original column names with whitespaces"""

        dataframe_columns_names = [column.replace("_", " ") for column in
                                   self.orders_data_frame.columns]
        self.orders_data_frame.columns = dataframe_columns_names
