#!/usr/bin/env python

#This is small auto Python script file to copying only worksheet of XLS to new worksheet in new XLS.
#Created by Tommas Huang 
#Created date: 2020-01-31

#You may need to install the additional plugin depending on which filetype you're working with: pip install pyexcel pyexcel-xls

#pyexcel provides one application programming interface to read, manipulate and write data in various excel formats.
import pyexcel as pe

#Save output-book path.
outputbook = "merged.xls"
#input-books is a dictionary mapping of the workbooks you wish to read in against their sheet.
inputbooks = {
  "/Users/TommasHuang/Documents/Test/test1.xls" : 'lab',
  "/Users/TommasHuang/Documents/Test/test2.xls" : 'drugs',
  "/Users/TommasHuang/Documents/Test/test3.xls" : 'drugs2',
}

merged_book = None
for book, sheet in inputbooks.iteritems():
  wb = pe.get_book(file_name=book)
  if merged_book is None:
    merged_book = wb[sheet]
  else:
    merged_book = merged_book + wb[sheet]
    
merged_book.save_as(outputbook)
   