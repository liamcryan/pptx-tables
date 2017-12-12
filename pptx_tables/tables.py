from pptx import Presentation
from pptx.util import Inches

from pptx_tables.columns import Columns
from pptx_tables.rows import Rows


class PptxTable(object):
    def __init__(self, data, presentation=None):
        self.prs = presentation
        self.table_args = None

        self.rows = Rows(data=data)
        self.columns = Columns(data=data)

        self.table_data = None

    def merge(self):
        table = []

        for i in self.rows.idx:
            if isinstance(self.rows.idx[i], dict):
                if self.columns.order:
                    self.rows.idx[i] = [self.rows.idx[i][key] for key in self.columns.order]
                else:
                    self.rows.idx[i] = [self.rows.idx[i][key] for key in sorted(self.rows.idx[i].keys())]
            elif isinstance(self.rows.idx[i], list):
                if self.columns.order:
                    self.rows.idx[i] = [self.rows.idx[i][j] for j in self.columns.order]

            if self.rows.head and self.rows.tail:
                if self.columns.head and i == 0:
                    table.append([None] + self.columns.head + [None])
                table.append([self.rows.head[i]] + self.rows.idx[i] + [self.rows.tail[i]])
                if self.columns.tail and i == len(self.rows.idx):
                    table.append([None] + self.columns.tail + [None])

            elif self.rows.head:
                if self.columns.head and i == 0:
                    table.append([None] + self.columns.head)
                table.append([self.rows.head[i]] + self.rows.idx[i])
                if self.columns.tail and i == len(self.rows.idx):
                    table.append([None] + self.columns.tail)

            elif self.rows.tail:
                if self.columns.head and i == 0:
                    table.append(self.columns.head + [None])
                table.append(self.rows.idx[i] + [self.rows.tail[i]])
                if self.columns.tail and i == len(self.rows.idx):
                    table.append(self.columns.tail + [None])
            else:
                table.append(self.rows.idx[i])

        self.table_data = table

    def create_table(self, slide_index=0):
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

        for i, row in enumerate(self.table_data):
            for j, col in enumerate(row):
                table.cell(i, j).text = str(self.table_data[i][j]) if self.table_data[i][j] is not None else ""

    def save_pptx(self, file_name):
        self.prs.save(file_name)

    def set_table_location(self, left, top, width, height):
        self.table_args = [len(self.table_data), len(self.table_data[0]), left, top, width, height]


if __name__ == "__main__":
    data1 = [[0, 1], [1, 1], [2, 1]]

    data2 = [{"apples": 1, "bananas": 0, "pears": 1},
             {"bananas": 1, "apples": 0, "pears": 1},
             {"apples": 1, "bananas": 0, "pears": 2}]

    tbl = PptxTable(data1)
    tbl.columns.set_order([1, 0])
    tbl.merge()

    tbl.set_table_location(Inches(0), Inches(1), Inches(5), Inches(2))
    tbl.create_table(slide_index=0)
    tbl.save_pptx("test1.pptx")

    tbl = PptxTable(data2)
    tbl.columns.set_head(["Apples", "Pears", "Bananas"])  # head and order should be linked somehow.
    tbl.columns.set_order(["apples", "pears", "bananas"])
    tbl.rows.set_head([0, 1, 2])
    tbl.merge()
    tbl.set_table_location(Inches(0), Inches(3), Inches(5), Inches(2))
    tbl.create_table(slide_index=0)
    tbl.save_pptx("test2.pptx")
