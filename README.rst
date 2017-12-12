===========
Pptx Tables
===========

Use pptx_tables to create tables more easily through python-pptx.


Features
--------

- Provide data formatted as a list of lists or a list of dictionaries

- Subset the columns/rows you want to display

- Provide custom headers

- Provide custom sort order on columns


Example
-------

>>> from pptx_tables import PptxTable
>>> from pptx.util import Inches
>>>
>>> data1 = [[1, 2],
>>>            [3, 4],
>>>            [5, 6]]
>>>
>>> tbl = PptxTable(data1)
>>> tbl.merge()
>>> tbl.set_table_location(Inches(0), Inches(0), Inches(5), Inches(1))
>>> tbl.create_table(slide_index=0)
>>> tbl.save_pptx("test1.pptx")

If you are already working on a PowerPoint, you can create a table on your current slide like this:

>>> from pptx import Presentation
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(self.prs.slide_layouts[1])  # suppose this adds the 3rd slide on your presentation
>>>
>>> tbl = PptxTable(data1, presentation=prs)
>>> tbl.merge()
>>> tbl.set_table_location(Inches(0), Inches(0), Inches(5), Inches(1))
>>> tbl.create_table(slide_index=2)

Here is data formatted as a list of un-nested dictionaries

>>> data2 = [{"apples": 1, "bananas": 2, "carrots": 6, "potatoes": 4},
>>>             {"apples": 6, "bananas": 3, "carrots": 4, "potatoes"; 3}]
>>> tbl = PptxTable(data2)
>>> tbl.columns.slice(["apples", "bananas"])
>>> tbl.columns.set_order(["bananas", "apples"])
>>> tbl.columns.set_head(["Bananas", "Apples"])
>>> tbl.merge()
>>> tbl.set_table_location(Inches(0), Inches(0), Inches(5), Inches(1))
>>> tbl.create_table()


Roadmap
-------

- set columns width
- set rows width
- font size
- cell alignment
