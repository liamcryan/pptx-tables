from pptx_tables.columns import Columns
from pptx_tables.rows import Rows


class Collection(object):
    def __init__(self, data):
        self.data = data
        self.columns = Columns(data)
        self.rows = Rows(data)

    def set_column_headers(self, headers):
        """ Updates the column index to account for the headers and updates the data self.data.

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


