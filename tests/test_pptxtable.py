from pptx.util import Inches

from pptx_tables.tables import PptxTable


class TestPptxTable:
    presentation = None
    data1 = [[1, 2],
             [3, 4],
             [5, 6]]

    data2 = [{"apples": 1, "bananas": 0, "pears": 1},
             {"bananas": 1, "apples": 0, "pears": 1},
             {"apples": 1, "bananas": 0, "pears": 2}]

    def test_data1_no_arguments(self):
        tbl = PptxTable(self.data1, presentation=self.presentation)
        tbl.set_table_location(Inches(0), Inches(2), Inches(5), Inches(1))
        tbl.create_table(slide_index=0)
        tbl.save_pptx("test2.pptx")

    def test_data2(self):
        tbl = PptxTable(self.data2, presentation=self.presentation)
        tbl.set_table_location(Inches(0), Inches(3), Inches(5), Inches(2))
        tbl.create_table(slide_index=0)
        tbl.save_pptx("test4.pptx")
