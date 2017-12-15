

class Columns(object):
    """  This class contains information related to a table's columns.

    Attributes:
        idx: this is the index information from the data

    """
    def __init__(self, data):
        """ Instantiate the class with data.

        :param data:  this is the data to be placed in the table.  this class gets called from a Collection.
                    must look like [[1, 0], [0, 0], [2, 1]]
                    OR
                    must look like [{"item1": 1, "item2": 2}, {"item1": 3, "item2": 3}]

        """
        self.idx = self._index(data)

    def _index(self, data):
        """ Index the columns.

        Here are examples of data for reference:
        [[0, 1, 2],
            [3, 4, 5],
                [6, 7, 8]]

        [{"apples": 0, "bananas": 1, "pears": 2},
             {"apples": 3, "bananas": 4, "pears": 5},
             {"apples": 6, "bananas": 7, "pears": 8}]

        """
        indexed = []
        for i, col in enumerate(data[0]):
            indexed.append(i) if isinstance(col, int) else indexed.append(col)
        return indexed

    def sort_order(self, order):
        """ Sets the new column index.

        Here are examples of data for reference:
        [[0, 1, 2],
            [3, 4, 5],
                [6, 7, 8]]

        [{"apples": 0, "bananas": 1, "pears": 2},
             {"apples": 3, "bananas": 4, "pears": 5},
             {"apples": 6, "bananas": 7, "pears": 8}]

        :param order: a list of integers
        :return: None
        """
        self.idx = order