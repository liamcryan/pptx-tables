from pptx import Presentation
from pptx.util import Inches

from pptx_tables.collections import Collection


class PptxTable(object):
    """  This class represents a PowerPoint Table.

    Attributes:
        prs:  a Python-pptx presentation
        table_args:  the arguments needed for Python-pptx presentation.add_table method
        pptx_table:  the table returned from the Python-pptx presentation.add_table method

    """
    def __init__(self, data, presentation=None):
        """   Instantiate the class with data and an optional presentation.

        :param data: this is the data to be placed in the table
                    must look like [[1, 0], [0, 0], [2, 1]]
                    OR
                    must look like [{"item1": 1, "item2": 2}, {"item1": 3, "item2": 3}]
        :param presentation: this is the Python-pptx presentation.  If not provided, a default presentation
                            will be created.
        """
        self.prs = presentation
        self.table_args = None
        self.pptx_table = None

        self.collection = Collection(data=data)

    def _add_table(self, slide_index=0):
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
        self.set_table_size(len(self.collection.rows.idx), len(self.collection.columns.idx))
        table = shapes.add_table(*self.table_args).table
        self.pptx_table = table

    def create_table(self, slide_index=0, rows_sort_order=None, columns_sort_order=None, column_headers=None):
        """ Sorts the rows/columns. Provides column headers.  Creates table. Puts data in a table.

        Here are examples of data for reference:
        [[0, 1, 2],
            [3, 4, 5],
                [6, 7, 8]]

        [{"apples": 0, "bananas": 1, "pears": 2},
             {"apples": 3, "bananas": 4, "pears": 5},
             {"apples": 6, "bananas": 7, "pears": 8}]

        :param slide_index:  what slide do you want to put the table on?
        :param rows_sort_order:  list, how to sort rows
        :param columns_sort_order: list, how to sort columns
        :param column_headers: list, what to call columns (dependent upon sorting of columns) for example:
                                columns are sorted in this order : [0, 1, 2],
                                columns headers should be something like : ["column_0", "column_1", "column_2"]
                                columns are sorted in this order : [2, 1, 0],
                                columns headers should be something like : ["column_2", "column_1", "column_0"]
        :return: None
        """
        if rows_sort_order:
            self.collection.rows.sort_order(rows_sort_order)
        if columns_sort_order:
            self.collection.columns.sort_order(columns_sort_order)
        if column_headers:
            self.collection.set_column_headers(column_headers)

        self._add_table(slide_index)

        for i, row in enumerate(self.collection.rows.idx):
            for j, col in enumerate(self.collection.columns.idx):
                self.pptx_table.cell(i, j).text = str(self.collection.data[row][col])

    def save_pptx(self, file_name):
        self.prs.save(file_name)

    def set_table_location(self, left, top, width, height):
        self.table_args = [None, None, left, top, width, height]

    def set_table_size(self, rows, columns):
        self.table_args[0] = rows
        self.table_args[1] = columns

if __name__ == "__main__":
    data1 = [[0, 1, 2], [3, 4, 5], [6, 7, 8]]

    tbl1 = PptxTable(data1)
    tbl1.set_table_location(Inches(0), Inches(1), Inches(5), Inches(2))
    tbl1.create_table(slide_index=0,
                      rows_sort_order=[2, 1, 0],
                      columns_sort_order=[2, 1, 0],
                      column_headers=["column1", "column2", "column3"])
    tbl1.save_pptx("test1.pptx")

    data2 = [{"apples": 0, "bananas": 1, "pears": 2},
             {"apples": 3, "bananas": 4, "pears": 5},
             {"apples": 6, "bananas": 7, "pears": 8}]

    tbl2 = PptxTable(data2)

    tbl2.set_table_location(Inches(0), Inches(3), Inches(5), Inches(2))
    tbl2.create_table(slide_index=0,
                      rows_sort_order=[2, 1, 0],
                      columns_sort_order=["pears", "bananas", "apples"],
                      column_headers=["Pears", "Bananas", "Apples"])
    tbl2.save_pptx("test2.pptx")
