# ils2py

This is an implementation of an Integrated Library System using web2py.

This is a personal project for organizing the book library for LASFS (http://lasfs.org/)


web2py is a python web framework - see http://web2py.com/


Current implementation:

The current way that items (books, magazines, DVDs) are tracked is through a spreadsheet.

There are seperate tabs of the spreadsheet for categories of items.

The columns have information, such as `WHO` the item is checked out, which librarian (`LIB`) checked the item out, when it was checked out, etc.

When a `Book` item is checked out, the following fields have information:

* `OUT`
* `DUE`
* `Borrowed By`
* `LIB`
* `RETURN`
* `LOCATION`
* `TYPE`
* `Number`
* `TITLE`
* `AUTHOR`
* `COAUTHOR`
* `Comments`
* `PUBLISHER`
* `SERIES`
* `ENTERED`
* `ISBN`
* `Donor`
* `MSRP`

The Magazine item has similar fields:

* `OUT`
* `DUE`
* `WHO`
* `LIB`
* `RETURNED`
* `Discard?`
* `LOCATION`
* `NUMBER`
* `! TITLE`
* `YEAR`
* `MONTH`
* `VOLUME`
* `VOLNUM`
* `WHOLE`
* `COMMENTS`
* `ENTERED`
* `More Comments`

## Stages and Migration plan

Import the library and normalize data

* publishers
* locations
* types

Dependencies:

* xlrd - http://www.python-excel.org/ https://pypi.python.org/pypi/xlrd - for importing current library  
