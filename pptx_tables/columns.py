

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

        :param data: data passed into the class.
        :return: None
        """
        indexed = []
        for i, col in enumerate(data[0]):
            if isinstance(col, int):
                indexed.append(i)
            elif isinstance(col, str):
                indexed.append(col)
                indexed = sorted(indexed)
        return indexed

    def sort_order(self, order):
        """ Sets the new column index.

        :param order: a list of integers
        :return: None
        """
        self.idx = order
