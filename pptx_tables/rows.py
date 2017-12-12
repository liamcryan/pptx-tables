from pptx_tables.collections import Collection


class Rows(Collection):
    def __init__(self, data):
        Collection.__init__(self, data)
        self.idx = self.index_by(by=self.__repr__())

    def __repr__(self):
        return "rows"
