import os

from pptx_tables import PptxTable

from pptx.util import Inches, Pt  # this comes from Python-pptx


here = os.path.abspath(os.path.dirname(__file__))

data1 = [[0, 1, 2],
         [3, 4, 5],
         [6, 7, 8]]

tbl1 = PptxTable(data1)
tbl1.create_table()

tbl1.save_pptx(os.path.join(here, "docs", "test1.pptx"))

tbl2 = PptxTable(data1)
tbl2.set_table_location(left=Inches(0), top=Inches(3), width=Inches(5))
tbl2.set_formatting(font_size=Pt(11), row_height=Inches(1))
tbl2.create_table(slide_index=0)
tbl2.save_pptx(os.path.join(here, "docs", "test2.pptx"))

tbl3 = PptxTable(data1)
tbl3.set_table_location(left=Inches(2), top=Inches(3), width=Inches(5))
tbl3.set_formatting(font_size=Pt(11), row_height=Inches(1))
tbl3.create_table(slide_index=0,
                  columns_headers=["column0", "column1", "column2"])
tbl3.save_pptx(os.path.join(here, "docs", "test3.pptx"))

tbl4 = PptxTable(data1)
tbl4.set_table_location(left=Inches(2), top=Inches(3), width=Inches(5))
tbl4.set_formatting(font_size=Pt(11), row_height=Inches(1))
tbl4.create_table(slide_index=0,
                  columns_sort_order=[2, 1, 0],
                  # notice the column headers need to be changed to match the column sort order
                  columns_headers=["column2", "column0", "column1"])
tbl4.save_pptx(os.path.join(here, "docs", "test4.pptx"))

tbl5 = PptxTable(data1)
tbl5.set_table_location(left=Inches(2), top=Inches(3), width=Inches(5))
tbl5.set_formatting(font_size=Pt(11), row_height=Inches(1))
tbl5.create_table(slide_index=0,
                  columns_sort_order=[2, 1, 0],
                  # notice the column headers need to be changed to match the column sort order
                  columns_headers=["column2", "column0", "column1"],
                  # the numbers in the list correspond to the weight given to each column, 1 means unchanged
                  columns_widths_weight=[.75, .75, 1.5])
tbl5.save_pptx(os.path.join(here, "docs", "test5.pptx"))

data2 = [{"apples": 0, "bananas": 1, "pears": 2},
         {"apples": 3, "bananas": 4, "pears": 5},
         {"apples": 6, "bananas": 7, "pears": 8}]  # get the presentation containing the previous table
presentation = tbl5.prs
tbl6 = PptxTable(data2, presentation)
tbl6.set_table_location(left=Inches(2), top=Inches(1), width=Inches(4))
tbl6.create_table(slide_index=0,
                  # default sort order is alphabetically on the keys,
                  # so the column headers should be alphabetical in this case
                  columns_headers=["Apples", "Bananas", "Pears"])
tbl6.save_pptx(os.path.join(here, "docs", "test6.pptx"))
