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

    def test_data1_col_subset_row_subset_col_header(self):
        tbl = PptxTable(self.data1, col_subset=[0], row_subset=[1, 2], col_headers=["column1"])
        tbl.set_table_location(Inches(0), Inches(0), Inches(5), Inches(1))
        tbl.create_table(slide_index=0)
        tbl.save_table("test1.pptx")
        self.presentation = tbl.prs

    def test_data1_no_arguments(self):
        tbl = PptxTable(self.data1, presentation=self.presentation)
        tbl.set_table_location(Inches(0), Inches(2), Inches(5), Inches(1))
        tbl.create_table(slide_index=0)
        tbl.save_table("test2.pptx")

    def test_data2_col_subset_row_subset_col_header(self):
        tbl = PptxTable(self.data2, presentation=self.presentation, col_subset=["bananas", "pears"],
                        row_subset=[0, 1], col_headers=[{"bananas": "Bananas"},
                                                        {"pears": "Pears"}])
        tbl.set_table_location(Inches(0), Inches(3), Inches(5), Inches(2))
        tbl.create_table(slide_index=0)
        tbl.save_table("test3.pptx")

    def test_data2_no_arguments(self):
        tbl = PptxTable(self.data2, presentation=self.presentation)
        tbl.set_table_location(Inches(0), Inches(3), Inches(5), Inches(2))
        tbl.create_table(slide_index=0)
        tbl.save_table("test4.pptx")
