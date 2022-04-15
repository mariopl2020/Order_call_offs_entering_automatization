"""Module to process xlsx deliveries/ orders file"""
import pandas as pd


# import openpyxl


def create_dataframe_from_file(filename):
    """Creates dataframe object from xlsx file

    Args:
        filename (str): name of file to be processed by program"""

    filepath = r'./data/' + filename
    dataframe_from_file = pd.read_excel(io=filepath)
    data_frame = dataframe_from_file

    return data_frame


def clean_columns_names(data_frame):
    """Cleans dataframe's columns from whitespaces and instead of it gives underscore"""

    dataframe_column_names = [column.replace(" ", "_") for column in data_frame.columns]
    data_frame.columns = dataframe_column_names

    return data_frame


def remove_empty_lines(data_frame):
    """Removes empty rows from call off dataframe"""

    mask = data_frame["SKU"].notnull()
    data_frame = data_frame[mask]

    return data_frame


def give_complete_orders_dataframe(filename):
    """Groups all key processing methods to get complete dataframe shape from file form"""

    data_frame = create_dataframe_from_file(filename)
    data_frame = clean_columns_names(data_frame)

    return data_frame


def give_complete_call_off_dataframe(filename):
    """Groups all key processing methods to get complete dataframe shape from file form"""

    data_frame = create_dataframe_from_file(filename)
    data_frame = clean_columns_names(data_frame)
    data_frame = remove_empty_lines(data_frame)

    return data_frame


def write_dataframe_to_excel(dataframe, filename):
    """Writes provided dataframe into file 'output.xlsx'

    Args:
        dataframe (Dataframe): dataframe to be saved in xlsx file
        filename (str): name of file to be written"""

    filepath = r"./data/" + filename
    dataframe.to_excel(excel_writer=filepath, index=False)


def give_single_order(row_index, call_off_data_frame):
    """Gives complete row from dataframe as Series object and order number

    Args:
        row_index (int): number of row index in dataframe"""

    current_order = call_off_data_frame.iloc[row_index]
    current_order_no = call_off_data_frame.iloc[row_index]["Customer_PO_number"]

    return current_order, current_order_no


def match_order_from_call_off(orders_dataframe, current_call_off_order_no):
    """Finds row in order dataframe corresponding to order number from call off. Write it
    down and its index"""

    mask = (orders_dataframe["Purchase_order_number"] == current_call_off_order_no)
    current_order = orders_dataframe[mask]
    current_order_index = current_order.index

    return current_order, current_order_index


def create_working_dataframe(orders_dataframe, current_order_index, current_order):
    """Deletes order row from order dataframe and double it into working dataframe. It starts
    splitting order into rows what will be delivered and remaining part of order still to be
    called off"""

    orders_dataframe = orders_dataframe.drop(current_order_index)
    working_orders_dataframe = pd.concat(objs=[current_order, current_order], ignore_index=True)
    working_orders_dataframe["Delivered_(kg)"] = working_orders_dataframe["Delivered_(kg)"]. \
        astype("int32")

    return working_orders_dataframe, orders_dataframe


def process_row_to_be_delivered(working_orders_dataframe, current_call_off_order):
    """Writes down values from order from call off dataframe as part of order to be delivered"""

    working_orders_dataframe.loc[0, "Current_delivery_date"] = \
        current_call_off_order["Confirmed_delivery_date"]
    working_orders_dataframe.loc[0, "Status"] = "confirmed"
    working_orders_dataframe.loc[0, "Required_delivery_date_OTIF1"] = \
        current_call_off_order["Requested_date"]
    working_orders_dataframe.loc[0, "Required_quantity_OTIF1"] = \
        current_call_off_order["Called_quantity_[kg]"]
    working_orders_dataframe.loc[0, "Ordered_(kg)"] = \
        current_call_off_order["Confirmed_quantity_[kg]"]
    working_orders_dataframe.loc[0, "Delivered_(kg)"] = \
        int(current_call_off_order["Confirmed_quantity_[kg]"])

    return working_orders_dataframe


def process_row_with_remaining_quantity(working_orders_dataframe):
    """Corrects quantity values for row what indicates still remaining order's part"""

    working_orders_dataframe.loc[1, "Ordered_(kg)"] -= working_orders_dataframe.loc \
        [0, "Ordered_(kg)"]
    working_orders_dataframe.loc[1, "Delivered_(kg)"] = \
        working_orders_dataframe.loc[1, "Delivered_(kg)"] - working_orders_dataframe.loc \
            [0, "Delivered_(kg)"]
    if working_orders_dataframe.loc[1, "Delivered_(kg)"] < 0:
        working_orders_dataframe.loc[1, "Delivered_(kg)"] = 0

    return working_orders_dataframe


def change_order_part_postfix(working_orders_dataframe):
    """Changes order number's postfixes what indicates order part's number"""

    if "/" in working_orders_dataframe.loc[1, "Purchase_order_number"]:
        labour_list = working_orders_dataframe.loc[1, "Purchase_order_number"].split("/")
        labour_list[1] = str(int(labour_list[1]) + 1)
        working_orders_dataframe.loc[1, "Purchase_order_number"] = labour_list[0] \
                                                                   + "/" + labour_list[1]
    else:
        working_orders_dataframe.loc[0, "Purchase_order_number"] += "/1"
        working_orders_dataframe.loc[1, "Purchase_order_number"] += "/2"

    return working_orders_dataframe


def calculate_confirmed_part_row_index(working_orders_dataframe, orders_dataframe):
    """Calculates row index for new confirmed order's part based on sorted delivery
    dates in order's main dataframe """

    confirmed_part_delivery_date = working_orders_dataframe.loc[0, "Current_delivery_date"]
    mask = orders_dataframe["Current_delivery_date"] >= confirmed_part_delivery_date
    orders_with_equal_later_date = orders_dataframe[mask]
    confirmed_second_part_row_index = orders_with_equal_later_date.index[0] - 0.5

    return confirmed_second_part_row_index


def assign_row_indexes(working_orders_dataframe,
                       current_order_index,
                       confirmed_second_part_row_index):
    """Assigns row indexes to both parts of orders """

    working_orders_dataframe.index = [confirmed_second_part_row_index, current_order_index[0]]

    return working_orders_dataframe


def add_working_rows_to_orders_dataframe(orders_dataframe, working_orders_dataframe):
    """Rebuilds base orders dataframe by adding newly calculated order part's rows, sorts them by
    indexes and drops them"""

    orders_dataframe = pd.concat(objs=[orders_dataframe, working_orders_dataframe])
    orders_dataframe = orders_dataframe.sort_index()
    orders_dataframe = orders_dataframe.reset_index(drop=True)

    return orders_dataframe


def main_loop(call_off_dataframe, orders_dataframe):
    """Iterates on single order line from call off and calls main methods to process it"""

    for index in range(len(call_off_dataframe["Customer_PO_number"])):
        current_call_off_order, current_call_off_order_no =\
            give_single_order(index, call_off_dataframe)
        current_order, current_order_index =\
            match_order_from_call_off(orders_dataframe, current_call_off_order_no)
        working_orders_dataframe, orders_dataframe =\
            create_working_dataframe(orders_dataframe, current_order_index, current_order)
        working_orders_dataframe =\
            process_row_to_be_delivered(working_orders_dataframe, current_call_off_order)
        working_orders_dataframe = process_row_with_remaining_quantity(working_orders_dataframe)
        working_orders_dataframe = change_order_part_postfix(working_orders_dataframe)
        confirmed_second_part_row_index =\
            calculate_confirmed_part_row_index(working_orders_dataframe, orders_dataframe)
        working_orders_dataframe = assign_row_indexes(working_orders_dataframe, current_order_index,
                                                      confirmed_second_part_row_index)
        orders_dataframe =\
            add_working_rows_to_orders_dataframe(orders_dataframe, working_orders_dataframe)

    return orders_dataframe


orders = give_complete_orders_dataframe(r"orders.xlsx")
call_off = give_complete_call_off_dataframe(r"calloff_template.xlsx")
orders = main_loop(call_off, orders)
write_dataframe_to_excel(orders, "orders_output.xlsx")
