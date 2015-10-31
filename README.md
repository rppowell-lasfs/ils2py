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

The plan is to have stages for migration of the current library system into the new database scheme.

```
ils          Book       Magazine       migration-stage1(import)                      migration-stage2(normalize-1)
circulation  OUT        OUT            ils_item.item_state - ils_item_state          ils_item.item_state - ils_item_state
circulation  DUE        DUE            ils_loan                                      ils_loan
circulation  WHO        WHO            ils_loan                                      ils_loan
circulation  LIB        LIB            ils_loan                                      ils_loan
circulation  RETURN     RETURNED       ils_loan                                      ils_loan
circulation  LOCATION   LOCATION       ils_item.item_location - ils_item_location    item_location
biblio       TYPE                      ils_item.item_type - ils_item_type            ils_biblio.biblio_type -  ils_biblio_type
circulation  NUMBER     NUMBER         ils_item.item_id                              ils_item.item_id
biblio       TITLE      TITLE          ils_item.item_title                           ils_biblio.biblio_title
biblio                  YEAR           ils_item.item_year                            ils_biblio.biblio_year
biblio                  MONTH          ils_item.item_month                           ils_biblio.biblio_month - ils_biblio_month
biblio                  VOLUME         ils_item.item_volume                          ils_biblio.biblio_volume
biblio                  VOLNUM         ils_item.item_volnum                          ils_biblio.biblio_volnum
biblio                  WHOLE          ils_item.item_whole                           ils_biblio_tag - complete
circulation             Discard?       ils_item.item_condition - ils_item_condition  ils_item.item_condition - ils_item_condition
biblio       AUTHOR                    ils_item.item_author - ils_item_person        biblio_person / biblio_person_type
biblio       COAUTHOR                  ils_item.item_coauthor - ils_item_person      biblio_person / biblio_person_type
circulation  COMMENTS   COMMENTS       ils_item.item_comments                        ils_item.item_comments
biblio       PUBLISHER                 ils_item.item_publisher                       biblio_publisher
biblio       SERIES                    ils_item.item_series - ils_item_series        biblio_tag - series
circulation  ENTERED    ENTERED        ils_item.item_entered                         ils_item.item_entered
biblio       ISBN                      ils_item.item_isbn                            ils_biblio.biblio_isbn
circulation  Donor                     ils_item.item_donor                           ils_item.item_donor
circulation  MSRP                      ils_item.item_msrp                            ils_item.item_msrp
circulation             More Comments  ils_item.item_comments2                       ils_biblio_tag
```

The 1st phase - import the library

The 1st phase will be entirely based on the information currently available per item entry from the excel spreadsheet.
This will be used to build the `ils_item` item and corresponding `ils_item` entries:
* `Number`     - `ils_item.item_id`
* `TITLE`      - `ils_item.item_title`
* `TYPE`       - `ils_item.item_type` - `ils_item_type`
* `AUTHOR`     - `ils_item.item_author` - `ils_item_person`
* `COAUTHOR`   - `ils_item.item_coauthor` - `ils_item_person`
* `PUBLISHER`  - `ils_item.item_publisher` - `ils_item_person`
* `SERIES`     - `ils_item.item_series`
* `ISBN`       - `ils_item.item_isbn`
* `MSRP`       - `ils_item.item_msrp`


The 2nd phase - normalize biblio

The 2nd phase will be to build the `ils_biblio` entries from the current `ils_item_*` entries:

* `TITLE`      - `ils_biblio.biblio_title`
* `TYPE`       - `ils_biblio.biblio_type` - `ils_biblio_type`
* `PUBLISHER`  - `ils_biblio_publisher`
* `AUTHOR`     - `ils_biblio_person` - `ils_biblio_x_person`/`ils_biblio_person_type`
* `COAUTHOR`   - `ils_biblio_person` - `ils_biblio_x_person`/`ils_biblio_person_type`
* `SERIES`     - `ils_biblio_tag`
* `ISBN`       - `ils_item.item_isbn`
* `MSRP`       - `ils_item.item_msrp`





Import the library and normalize data

* publishers
* locations
* types

Dependencies:

* xlrd - http://www.python-excel.org/ https://pypi.python.org/pypi/xlrd - for importing current library  
