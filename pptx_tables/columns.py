from pptx_tables.collections import Collection


class Columns(Collection):
    def __init__(self, data):
        Collection.__init__(self, data)
        self.idx = self.index_by(by=self.__repr__())
        self.order = None

    def __repr__(self):
        return "columns"

    def set_order(self, order):
        self.order = order
        