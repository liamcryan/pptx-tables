from pptx_tables.columns import Columns
from pptx_tables.rows import Rows


class Collection(object):
    """  This class represents a collection of data that is to be separated into rows and columns.

    Attributes:
        data: this is the data to be placed in the table
        columns: this is column information of the given data, a Columns object
        rows: this is the row information of the given data, a Rows object

    """
    def __init__(self, data):
        """ Instantiate the class with data.  this class gets called from a PptxTable.

        :param data: this is the data to be placed in the table
                    must look like [[1, 0], [0, 0], [2, 1]]
                    OR
                    must look like [{"item1": 1, "item2": 2}, {"item1": 3, "item2": 3}]
        """
        self.data = data
        self.columns = Columns(data)
        self.rows = Rows(data)

    def set_column_headers(self, headers):
        """ Updates the column index to account for the headers and updates the data self.data.

        Here are examples of data for reference:
        [[0, 1, 2],
            [3, 4, 5],
                [6, 7, 8]]

        [{"apples": 0, "bananas": 1, "pears": 2},
             {"apples": 3, "bananas": 4, "pears": 5},
             {"apples": 6, "bananas": 7, "pears": 8}]

        :param headers: a list of integers
        :return: None
        """
        if isinstance(self.data[0], list):
            self.data = [headers] + self.data

            increment = [i + 1 for i in self.rows.idx]
            self.rows.idx = [0] + increment

        elif isinstance(self.data[0], dict):
            datum = {}
            for i, key in enumerate(self.columns.idx):
                datum.update({key: headers[i]})
            self.data = [datum] + self.data

            increment = [i + 1 for i in self.rows.idx]
            self.rows.idx = [0] + increment


