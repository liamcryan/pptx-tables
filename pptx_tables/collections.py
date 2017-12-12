

class Collection(object):
    def __init__(self, data):
        self.data = data
        self.idx = None
        self.head = None
        self.tail = None

    def index_by(self, by):
        if by == "rows":
            indexed = {}
            for i, r in enumerate(self.data):
                indexed.update({i: r})
            return indexed

        if by == "columns":
            def get_column_qty(_data):
                """ this is for lists """
                _cols = 0
                for _r in _data:
                    if len(_r) > _cols:
                        _cols = len(_r)
                return _cols

            def get_column_keys(_data):
                """ this is for dictionaries """
                _cols = []
                for _r in _data:
                    for _c in _r:
                        if _c not in _cols:
                            _cols.append(_c)
                return sorted(_cols)

            if isinstance(self.data[0], dict):
                cols = get_column_keys(self.data)
                indexed = {i: [] for i in cols}

                for r in self.data:
                    for c in r:
                        indexed[c].append(r[c])

                return indexed

            elif isinstance(self.data[0], list):
                cols = get_column_qty(self.data)
                indexed = {i: [] for i in range(0, cols)}

                for r in self.data:
                    for j, c in enumerate(r):
                        indexed[j].append(c)

                return indexed

    def slice(self, subset=None):
        if isinstance(self.idx, list):
            sliced = []
            for i, key in enumerate(self.idx):
                sliced.append({i: self.idx[i]})
            return sliced

        elif isinstance(self.idx, dict):
            if not subset:
                return self.idx
            else:
                sliced = []
                for key in subset:
                    sliced.append({key: self.idx[key]})
                return sliced

    def set_head(self, alias):
        self.head = alias

    def set_tail(self, alias):
        self.tail = alias
