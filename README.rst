==================
Python Pptx Tables
==================

Use Python-pptx-tables to create tables more easily through Python-pptx.

Example
-------

>>>  from pptx_tables import PptxTable
>>>
>>>  data1 = [[1, 2],
>>>             [3, 4],
>>>             [5, 6]]
>>>
>>>  tbl = PptxTable(data1)
>>>  tbl.set_table_location(Inches(0), Inches(0), Inches(5), Inches(1))
>>>  tbl.create_table(slide_index=0)
>>>  tbl.save_table("test1.pptx")

If you are already working on a PowerPoint, you can create a table on your current slide like this:

>>>  from pptx import Presentation
>>>  prs = Presentation()
>>>  slide = prs.slides.add_slide(self.prs.slide_layouts[1])  # suppose this adds the 3rd slide on your presentation
>>>
>>>  tbl2 = PptxTable(data1, presentation=prs)
>>>  tbl2.set_table_location(Inches(0), Inches(0), Inches(5), Inches(1))
>>>  tbl2.create_table(slide_index=2)  # slide number is one more than slide index


Features
--------

Provide data formatted as [[0, 0], [1, 1]] or [{"col1": 1, "col2": 0}, {"col2": 0, "col2": 1}]

Subset the columns/rows you want to display

Provide custom headers
