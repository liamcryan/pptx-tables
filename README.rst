===========
Pptx Tables
===========

Use pptx_tables to create tables more easily through python-pptx.


Features
========

- Provide data formatted as a list of lists or a list of dictionaries

- Provide custom headers

- Provide custom sort order on columns

- Set columns width

- Set font size

- Set cell alignment

- Set the row height


Samples
=======

Create a table of data on a slide
---------------------------------

>>>    from pptx_tables import PptxTable
>>>
>>>    data1 = [[0, 1, 2],
>>>             [3, 4, 5],
>>>             [6, 7, 8]]
>>>
>>>    tbl1 = PptxTable(data1)
>>>    tbl1.create_table()
>>>    tbl1.save_pptx("test1.pptx")

.. image:: /docs/test1.png


Set location of table and provide some formatting
-------------------------------------------------

>>>    from pptx.util import Inches, Pt  # this comes from Python-pptx
>>>
>>>    tbl2 = PptxTable(data1)
>>>    tbl2.set_table_location(left=Inches(0), top=Inches(3), width=Inches(5))
>>>    tbl2.set_formatting(font_size=Pt(7), row_height=Inches(.5))
>>>    tbl2.create_table(slide_index=0)
>>>    tbl2.save_pptx("test2.pptx")

.. image:: /docs/test2.png

Create column headers
---------------------

>>>    tbl3 = PptxTable(data1)
>>>    tbl3.set_table_location(left=Inches(0), top=Inches(3), width=Inches(5))
>>>    tbl3.set_formatting(font_size=Pt(7), row_height=Inches(.5))
>>>    tbl3.create_table(slide_index=0,
>>>                      columns_headers=["column0", "column1", "column2"])
>>>    tbl3.save_pptx("test3.pptx")

.. image:: /docs/test3.png

Sort columns
------------

>>>    tbl4 = PptxTable(data1)
>>>    tbl4.set_table_location(left=Inches(0), top=Inches(3), width=Inches(5))
>>>    tbl4.set_formatting(font_size=Pt(9), row_height=Inches(.5))
>>>    tbl4.create_table(slide_index=0,
>>>                      columns_sort_order=[2, 1, 0],
>>>                      # notice the column headers need to be changed to match the column sort order
>>>                      columns_headers=["column2", "column0", "column1"])
>>>    tbl4.save_pptx("test4.pptx")

.. image:: /docs/test4.png

Set column widths
-----------------

>>>    tbl5 = PptxTable(data1)
>>>    tbl5.set_table_location(left=Inches(0), top=Inches(3), width=Inches(5))
>>>    tbl5.set_formatting(font_size=Pt(9), row_height=Inches(.5))
>>>    tbl5.create_table(slide_index=0,
>>>                      columns_sort_order=[2, 1, 0],
>>>                      # notice the column headers need to be changed to match the column sort order
>>>                      columns_headers=["column2", "column0", "column1"],
>>>                      # the numbers in the list correspond to the weight given to each column, 1 means unchanged
>>>                      columns_widths_weight=[.75, .75, 1.5])
>>>    tbl5.save_pptx("test5.pptx")

.. image:: /docs/test5.png

Add another table to the same slide
-----------------------------------

>>>    # here is some new data, oh by the way, it's also formatted differently
>>>    data2 = [{"apples": 0, "bananas": 1, "pears": 2},
>>>             {"apples": 3, "bananas": 4, "pears": 5},
>>>             {"apples": 6, "bananas": 7, "pears": 8}]
>>>
>>>    # get the presentation containing the previous table
>>>    presentation = tbl5.prs
>>>    tbl6 = PptxTable(data2, presentation)
>>>    tbl6.set_table_location(left=Inches(0), top=Inches(5), width=Inches(4))
>>>    tbl6.create_table(slide_index=0,
>>>                      # default sort order is alphabetically on the keys,
>>>                      # so the column headers should be alphabetical in this case
>>>                      columns_headers=["Apples", "Bananas", "Pears"])
>>>    tbl6.save_pptx("test6.pptx")

.. image:: /docs/test6.png

Transpose a table
-----------------

>>> tbl7 = PptxTable(data2)
>>> tbl7.set_table_location(left=Inches(2), top=Inches(1), width=Inches(4))
>>> tbl7.create_table(slide_index=0,
>>>                   columns_headers=["Apples", "Bananas", "Pears"],  # column headers become the row headers
>>>                   columns_widths_weight=[1.5, .5, .5, .5],  # since transpose need 4 columns weights instead of 3
>>>                   transpose=True)
>>> tbl7.save_pptx(os.path.join(here, "docs", "test7.pptx"))

.. image:: /docs/test7.png