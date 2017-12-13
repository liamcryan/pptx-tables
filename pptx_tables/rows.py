

class Rows(object):
    def __init__(self, data):
        self.idx = self._index(data)

    def _index(self, data):
        """ Index the rows.

        [[0, 1, 2],
            [3, 4, 5],
                [6, 7, 8]]

        [{"apples": 0, "bananas": 1, "pears": 2},
             {"apples": 3, "bananas": 4, "pears": 5},
             {"apples": 6, "bananas": 7, "pears": 8}]

        """
        indexed = []
        for i, row in enumerate(data):
            indexed.append(i)
        return indexed

    def sort_order(self, order):
        """ Sets the new row index.

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
