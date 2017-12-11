from pptx import Presentation
from pptx.util import Inches


class PptxTable(object):
    def __init__(self, data, col_subset=None, row_subset=None, col_headers=None, presentation=None):
        self.col_subset = [] if not col_subset else col_subset
        self.row_subset = [] if not row_subset else row_subset
        self.col_headers = col_headers

        self.table_args = None
        self.col_alias = None
        self.col_widths = None
        self.fixed_row_width = .38

        self.table = None
        self.prs = presentation

        self.row_qty = None
        self.col_qty = None
        self.data = self.process_data(data)

    def create_table(self, slide_index=0):
        if not self.prs:
            if slide_index > 0:
                raise Exception("slide index too high")
            prs = Presentation()
            self.prs = prs
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[1])
        else:
            if slide_index > len(list(self.prs.slides._sldIdLst)) - 1:
                raise Exception("slide index too high")
            slide = self.prs.slides[slide_index]

        shapes = slide.shapes
        table = shapes.add_table(*self.table_args).table

        for i, row in enumerate(self.data):
            for j, col in enumerate(row):
                table.cell(i, j).text = str(self.data[i][j])

        self.table = table

    def save_table(self, file_name):
        self.prs.save(file_name)

    def set_table_location(self, left, top, width, height):
        self.table_args = [self.row_qty, self.col_qty, left, top, width, height]

    def set_column_width_proportions(self, proportion_list):
        self.col_widths = proportion_list

    def reset_fixed_row_width(self, row_width):
        self.fixed_row_width = row_width

    def set_column_alias(self, alias):  # {"apples": "Apples", "bananas": "Bananas"}
        self.col_alias = alias

    def process_data(self, d):
        if (isinstance(d, list) or isinstance(d, tuple)) and len(d) > 0:
            if self.row_subset:
                self.row_qty = len(self.row_subset) + 1
            else:
                self.row_qty = len(d) + 1
                self.row_subset = [i for i in range(0, self.row_qty)]

            if isinstance(d[0], list) or isinstance(d[0], tuple) or isinstance(d[0], dict):

                if self.col_subset:
                    self.col_qty = len(self.col_subset)
                else:
                    self.col_qty = len(d[0])
                    if not isinstance(d[0], dict):
                        self.col_subset = [i for i in range(0, self.col_qty)]
                    else:
                        self.col_subset = sorted([k for k in d[0]])

            else:
                raise Exception("must be a list, tuple or dict")
        else:
            raise Exception("must be a list or tuple")

        processed_data = []
        # if dict: create header row of keys
        # if list or tuple, create header row of index

        for i, row in enumerate(d):
            if i in self.row_subset:
                row_j = []
                for j, col in enumerate(row):
                    if not isinstance(d[i], dict):
                        if j in self.col_subset:
                            row_j.append(col)
                    else:
                        if col in self.col_subset:
                            row_j.append(d[i][col])

                processed_data.append(row_j)

        if self.col_headers:
            if not isinstance(self.col_headers[0], dict):
                processed_data = [self.col_headers] + processed_data
            else:
                col_headers = []
                for header in self.col_headers:
                    for key in header.keys():
                        col_headers.append(header[key])
                processed_data = [col_headers] + processed_data
        else:
            if not isinstance(d[0], dict):
                col_headers = self.col_subset
                processed_data = [col_headers] + processed_data
            else:
                col_headers = []
                for key in d[0]:
                    col_headers.append(key)
                processed_data = [col_headers] + processed_data

        return processed_data


if __name__ == "__main__":
    data2 = [{"apples": 1, "bananas": 0, "pears": 1},
             {"bananas": 1, "apples": 0, "pears": 1},
             {"apples": 1, "bananas": 0, "pears": 2}]

    tbl = PptxTable(data2, col_subset=["bananas", "pears"],
                    row_subset=[0, 1], col_headers=[{"bananas": "Bananas"},
                                                    {"pears": "Pears"}])
    tbl.set_table_location(Inches(0), Inches(3), Inches(5), Inches(2))
    tbl.create_table(slide_index=0)
    tbl.save_table("test3.pptx")
