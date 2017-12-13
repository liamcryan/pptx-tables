from pptx import Presentation
from pptx.util import Inches

from pptx_tables.collections import Collection


class PptxTable(object):
    def __init__(self, data, presentation=None):
        self.prs = presentation
        self.table_args = None
        self.pptx_table = None

        self.collection = Collection(data=data)

    def create_table(self, slide_index=0):
        """ Creates a slide if needed then adds table according to the table_args provided.

        :param slide_index: slide you want to add to...only really needed if you have an existing presentation
        :return: None
        """
        if not self.prs:
            if slide_index > 0:
                raise Exception("slide index provided is greater than the number of slides in the report")
            prs = Presentation()
            self.prs = prs
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[1])
        else:
            if slide_index > len(list(self.prs.slides._sldIdLst)) - 1:
                raise Exception("slide index provided is greater than the number of slides in the report")
            slide = self.prs.slides[slide_index]

        shapes = slide.shapes
        table = shapes.add_table(*self.table_args).table

        self.pptx_table = table

    def populate_table(self):
        """ Puts data in a table.

        [[0, 1, 2],
            [3, 4, 5],
                [6, 7, 8]]

        [{"apples": 0, "bananas": 1, "pears": 2},
             {"apples": 3, "bananas": 4, "pears": 5},
             {"apples": 6, "bananas": 7, "pears": 8}]

        :return: None
        """

        for i, row in enumerate(self.collection.rows.idx):
            for j, col in enumerate(self.collection.columns.idx):
                self.pptx_table.cell(i, j).text = str(self.collection.data[row][col])

    def save_pptx(self, file_name):
        self.prs.save(file_name)

    def set_table_location(self, left, top, width, height):
        self.table_args = [len(self.collection.rows.idx), len(self.collection.columns.idx), left, top, width, height]


if __name__ == "__main__":
    data1 = [[0, 1, 2], [3, 4, 5], [6, 7, 8]]

    tbl1 = PptxTable(data1)
    tbl1.collection.columns.sort_order([2, 1, 0])
    tbl1.collection.rows.sort_order([2, 1, 0])
    tbl1.collection.set_column_headers(["column1", "column2", "column3"])  # this must be after sort order...before set table location
    tbl1.set_table_location(Inches(0), Inches(1), Inches(5), Inches(2))
    tbl1.create_table(slide_index=0)
    tbl1.populate_table()
    tbl1.save_pptx("test1.pptx")

    data2 = [{"apples": 0, "bananas": 1, "pears": 2},
             {"apples": 3, "bananas": 4, "pears": 5},
             {"apples": 6, "bananas": 7, "pears": 8}]

    tbl2 = PptxTable(data2)
    tbl2.collection.rows.sort_order([2, 1, 0])  # sorting must be done before setting the headers...
    tbl2.collection.columns.sort_order(["pears", "bananas", "apples"])
    tbl2.collection.set_column_headers(["Pears", "Bananas", "Apples"])

    tbl2.set_table_location(Inches(0), Inches(3), Inches(5), Inches(2))  # set table location after setting headers...must be
    tbl2.create_table(slide_index=0)  # more args...table_args, column_headers...this would prevent user from issues in comments above
    tbl2.populate_table()
    tbl2.save_pptx("test2.pptx")
