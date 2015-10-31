#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import re
import datetime

import logging

from import_library_books_xls import *
from import_library_magazines_xls import *

"""

Modules for importing the library from xls spreadsheet

xls spreadsheet row  -->  entry  --> collation

RawEntry
* TODO: invalid entry checks
** Invalid Data - TEXT
*** is not XL_CELL_TEXT
*** is XL_CELL_EMPTY

* Invalid Data - DATE
* Invalid Data - NUMBER

collation
* duplicate entry
* entry locations, types, categories, etc

"""

logger = logging.getLogger("web2py.app.ils2py.import_library_xls")
logger.setLevel(logging.DEBUG)

class RawEntryDataText:
    def read(self, cell):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".read()")
        logger.setLevel(logging.DEBUG)

        self.cell = cell
        if (self.cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.cell.value])
            self.data = None
            return None

        elif (self.cell.ctype == xlrd.XL_CELL_TEXT):
            self.data = self.cell.value.encode('utf-8')
            return self.cell.value.encode('utf-8')

        elif (self.cell.ctype == xlrd.XL_CELL_EMPTY):
            self.data = None
            return None

        else:
            self.data = None
            raise Exception(str(self.cell))


class RawEntryDataNumber:
    def read(self, cell):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".read()")
        logger.setLevel(logging.DEBUG)
        self.cell = cell

        if (self.cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.cell.value])
            return None

        elif (self.cell.ctype == xlrd.XL_CELL_NUMBER):
            return int(self.cell.value)

        elif (self.cell.ctype == xlrd.XL_CELL_EMPTY):
            return None

        else:
            raise Exception(str(self.cell))


class RawEntryDataDate:
    def read(self, cell, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".read()")
        logger.setLevel(logging.DEBUG)

        self.cell = cell
        if (self.cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.cell.value])
            return None
    
        elif (self.cell.ctype == xlrd.XL_CELL_DATE and self.cell.value != ''):
            return datetime.datetime(*xlrd.xldate_as_tuple(self.cell.value, xlrdworkbook.datemode))
    
        elif (self.cell.ctype == xlrd.XL_CELL_TEXT):
            if (self.cell.value == '' or self.cell.value == ' '):
                return None
            else:
                #raise Exception(str(self.cell))
                return None
    
        elif (self.cell.ctype == xlrd.XL_CELL_EMPTY):
            return None
    
        else:
            raise Exception(str(self.cell))

class RawEntryDataType:
    TEXT, NUMBER, DATE = range(3)

    @staticmethod
    def datatype_string_to_enum(s):
        return ['TEXT','NUMBER','DATE'].index(s)

    def __init__(self, name, index, datatype):
        self.name=name
        self.index=index
        self.datatype=datatype

    def read(self, row, xlrdworkbook, rowindex):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".read()")
        logger.setLevel(logging.DEBUG)

        self.cell = row[self.index]
        if self.datatype==self.TEXT:
            try:
                r = RawEntryDataText()
                r.read(self.cell)
                data = r.data
            except Exception as e:
                logger.debug(" ".join((str(x) for x in (e, self.name, self.index, self.cell, row, rowindex))))
                data = None

        elif self.datatype==self.NUMBER:
            try:
                r = RawEntryDataNumber()
                data =r.read(self.cell)
            except Exception as e:
                logger.debug(" ".join((str(x) for x in (e, self.name, self.index, self.cell, row, rowindex))))
                data = None
    
        elif self.datatype==self.DATE:
            try:
                r = RawEntryDataDate()
                data = r.read(self.cell, xlrdworkbook)
            except Exception as e:
                logger.debug(" ".join((str(x) for x in (e, self.name, self.index, self.cell, row, rowindex))))
                data = None

        return data


class RawEntry:
    ENTRY_FORMAT=[]

    def extract_data(self, rowindex, row, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".__init__()")
        logger.setLevel(logging.DEBUG)

        self.row = row
        self.entry={}

        for entry in self.ENTRY_FORMAT:
            data = entry.read(row, xlrdworkbook, rowindex)
            self.entry[entry.name] = data


class MemberEntry(RawEntry):
    ENTRY_FORMAT = [
            RawEntryDataType('NAME', 0, RawEntryDataType.TEXT),
    ]

    def __init__(self, index, row, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".__init__()")
        logger.setLevel(logging.DEBUG)

        self.row = row
        self.entry={}

        self.extract_data(index, row, xlrdworkbook)

class VideoEntry(RawEntry):
    ENTRY_FORMAT = [
            RawEntryDataType('TITLE', 0, RawEntryDataType.TEXT),
            RawEntryDataType('LC', 0, RawEntryDataType.TEXT),
            RawEntryDataType('!', 0, RawEntryDataType.TEXT),
            RawEntryDataType('LOCATION', 0, RawEntryDataType.TEXT),
            RawEntryDataType('Borrowed By', 0, RawEntryDataType.TEXT),
            RawEntryDataType('Checkout', 0, RawEntryDataType.TEXT),
            RawEntryDataType('Due', 0, RawEntryDataType.DATE),
            RawEntryDataType('Returned', 0, RawEntryDataType.DATE),
            RawEntryDataType('Libr', 0, RawEntryDataType.DATE),
            RawEntryDataType('COMMENT1', 0, RawEntryDataType.DATE),
            RawEntryDataType('COMMENT2', 0, RawEntryDataType.DATE),
    ]

    def __init__(self, index, row, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".__init__()")
        logger.setLevel(logging.DEBUG)

        self.row = row
        self.entry={}

        self.extract_data(index, row, xlrdworkbook)


def check_headers(xlrdworkbook):
    column_names = [
        'CK OUT',
        'DUE',
        'WHO',
        'LIB',
        'RETURN',
        'LOCATION',
        'TYPE',
        'NUMBER',
        'TITLE',
        'AUTHOR',
        'COAUTHOR',
        'COMMENTS',
        'PUBLISHER',
        'SERIES',
        'ENTERED'
    ]
    logger = logging.getLogger("import_library_books_xls.check_headers")
    logger.setLevel(logging.DEBUG)
    #sheet = xlrdworkbook.sheet_by_name('BOOK')
    sheet = xlrdworkbook.sheet_by_index(0)
    logger.debug("Sheet:{0}".format(sheet.name))
    for i, column_name in enumerate(column_names):
        c = sheet.cell(0, i)
        logger.debug("Checking Headers {0} - '{1}' == '{2}'".format(i, column_name, c.value))
        if c.value != column_name:
            return False
    return True


class LibraryXLS:
    """
    category     Book       Magazine
    circulation  OUT        OUT
    circulation  DUE        DUE
    circulation  WHO        WHO
    circulation  LIB        LIB
    circulation  RETURN     RETURNED
    circulation  LOCATION   LOCATION
    biblio       TYPE
    circulation  NUMBER     NUMBER
    biblio       TITLE      TITLE
    biblio       YEAR
    biblio       MONTH
    biblio       VOLUME
    biblio       VOLNUM
    biblio       WHOLE
    circulation  Discard?
    biblio       AUTHOR
    biblio       COAUTHOR
    circulation  COMMENTS   COMMENTS
    biblio       PUBLISHER
    biblio       SERIES
    circulation  ENTERED    ENTERED
    biblio       ISBN
    circulation  Donor
    circulation  MSRP
    circulation  More       Comments
    
    entries
    numbers
    duplicates

    locations
    types
    authors
    coauthors
    publishers
    discards
    years
    months
    volumes
    volnums
    series_completed

    """
    def __init__(self):
        self.library_entries=[]
        self.library_numbers={}
        self.library_duplicates=[]

        self.library_locations={}
        self.library_types={}
        self.library_authors={}
        self.library_comments={}
        self.library_coauthors={}
        self.library_publishers={}

        self.library_discard={}
        self.library_year={}
        self.library_month={}
        self.library_volume={}
        self.library_volnum={}
        self.library_series_completed={}

        self.library_circulations=[]
        self.library_members={}
        self.library_librarians={}

    def add_xls_book_entry(self, entry, entry_index):

        if ((entry['NUMBER']!='') and (entry['TITLE']!='')):
            self.library_entries.append(entry)

        if entry['NUMBER'] in self.library_numbers:
            self.library_duplicates.append(
                    [entry['NUMBER'], [self.library_numbers[entry['NUMBER']], entry_index]]
            )
        else:
            self.library_numbers[entry['NUMBER']] = entry_index

        self.library_locations[entry['LOCATION']] = \
            self.library_locations.get(entry['LOCATION'], []) + [entry['NUMBER']]
            
        self.library_types[entry['TYPE']] = \
            self.library_types.get(entry['TYPE'], []) + [entry['NUMBER']]

        # bibliography info
        if 'AUTHOR' in entry:
            self.library_authors[entry['AUTHOR']] = \
                self.library_authors.get(entry['AUTHOR'], []) + [entry['NUMBER']]
    
        if 'COAUTHOR' in entry:
            self.library_coauthors[entry['coauthor']] = \
                self.library_coauthors.get(entry['coauthor'], []) + [entry['number']]
    
        if 'PUBLISHER' in entry:
            self.library_publishers[entry['PUBLISHER']] = \
                self.library_publishers.get(entry['PUBLISHER'], []) + [entry['PUBLISHER']]

        # circulation info
        if 'WHO' in entry:
            self.library_members[entry['WHO']] = \
                self.library_members.get(entry['WHO'], []) + [entry['number']]

        if 'LIB' in entry:
            self.library_librarians[entry['LIB']] = \
                self.library_librarians.get(entry['LIB'], []) + [entry['number']]


        if ('OUT' in entry and 'DUE' in entry and 'WHO' in entry):
            if ('RETURN' in entry):
                pass
            else:
                self.library_circulations.append(entry)

    def import_library_books_xls(self, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".import_library_books_xls()")
        logger.setLevel(logging.DEBUG)

        sheet = xlrdworkbook.sheet_by_name('Books')

        for i in range(1,sheet.nrows):
            entry = BookEntry.process_row_data(sheet, i, xlrdworkbook)

            if not entry:
                continue

            self.add_xls_book_entry(entry, i)


def import_library_books_xls(xlrdworkbook):
    """Import workflow:
    For XLS sheet named 'Books'

    """
    logger = logging.getLogger("import_library_books_xls.import_library_books_xls")
    logger.setLevel(logging.DEBUG)

    sheet = xlrdworkbook.sheet_by_name('Books')

    entries = []
    circulations = []

    entry_numbers = {}
    duplicate_entries = []

    entry_locations = {}
    entry_types = {}
    entry_authors = {}
    entry_coauthors = {}
    entry_publishers = {}

    entry_members = {}
    entry_librarians = {}

    libraryxls = LibraryXLS()

    for i in range(1,sheet.nrows):
        entry = BookEntry.process_row_data(sheet, i, xlrdworkbook)

        # sanity / type check cell data
        
        if not entry:
            continue

        libraryxls.add_xls_entry(entry, i)

        if (entry['number']) in entry_numbers:
            logger.debug(" ".join((str(x) for x in (
                "duplicate entry:", i, entry
            ))))
            duplicate_entries.append([entry['number'], [entry_numbers[entry['number']], i]])
            continue
        else:
            entry_numbers[entry['number']]=i

        entry_locations[entry['location']] = \
            entry_locations.get(entry['location'], []) + [entry['number']]
            
        entry_types[entry['type']] = \
            entry_types.get(entry['type'], []) + [entry['number']]

        if ((entry['number']!='') and (entry['title']!='')):
            entries.append(entry)

        # bibliography info
        if 'author' in entry:
            entry_authors[entry['author']] = \
                entry_authors.get(entry['author'], []) + [entry['number']]
    
        if 'coauthor' in entry:
            entry_coauthors[entry['coauthor']] = \
                entry_coauthors.get(entry['coauthor'], []) + [entry['number']]
    
        if 'publisher' in entry:
            entry_publishers[entry['publisher']] = \
                entry_publishers.get(entry['publisher'], []) + [entry['publisher']]

        # circulation info
        if 'WHO' in entry:
            entry_members[entry['WHO']] = \
                entry_members.get(entry['WHO'], []) + [entry['number']]

        if 'LIB' in entry:
            entry_librarians[entry['LIB']] = \
                entry_librarians.get(entry['LIB'], []) + [entry['number']]


        if ('OUT' in entry and 'DUE' in entry and 'WHO' in entry):
            if ('RETURN' in entry):
                pass
            else:
                print "circulation:", i, entry
                circulation = entry.copy()
                circulations.append(entry)

    if (False):
        print "\nentries:"
        for i, entry in enumerate(entries):
            print i, entry

    if (True):
        print "\nduplicate_entries:", duplicate_entries
        print "\t".join(('number', 'index', 'author', 'title'))
        for k in duplicate_entries:
            i = k[0];
            i1 = k[1][0]
            i2 = k[1][1]
            print "\t".join((str(i), str(i1+1), sheet.row(i1)[BookEntry.row_index_AUTHOR].value, sheet.row(i1)[BookEntry.row_index_TITLE].value))
            print "\t".join((str(i), str(i2+1), sheet.row(i2)[BookEntry.row_index_AUTHOR].value, sheet.row(i2)[BookEntry.row_index_TITLE].value))
            print ""

    if (False):
        print "\nentry_locations:"
        for t in sorted([key for key in entry_locations.keys()]):
            print "'{}'".format(t), len(entry_locations[t]), entry_locations[t] if (len(entry_locations[t]) < 10) else entry_locations[t][:10] + ["..."]

    if (False):
        print "\nentry_types:"
        for t in sorted([key for key in entry_types.keys()]):
            print "'{}'".format(t), len(entry_types[t]), entry_types[t] if (len(entry_types[t]) < 10) else entry_types[t][:10] + ["..."]

    # bibliography
    if (False):
        print "\nentry_authors:"
        for t in sorted([key for key in entry_authors.keys()]):
            print "'{}'".format(t), len(entry_authors[t]), entry_authors[t] if (len(entry_authors[t]) < 10) else entry_authors[t][:10] + ["..."]

    if (False):
        print "\nentry_coauthors:"
        for t in sorted([key for key in entry_coauthors.keys()]):
            print "'{}'".format(t), len(entry_coauthors[t]), entry_coauthors[t] if (len(entry_coauthors[t]) < 10) else entry_coauthors[t][:10] + ["..."]

    if (False):
        print "\nentry_publishers:"
        for t in sorted([key for key in entry_publishers.keys()]):
            print "'{}'".format(t), len(entry_publishers[t]), entry_publishers[t] if (len(entry_publishers[t]) < 10) else entry_publishers[t][:10] + ["..."]

    # circulation
    if (False):
        print "\nentry_members:"
        for t in sorted([key for key in entry_members.keys()]):
            print "'{}'".format(t), len(entry_members[t]), entry_members[t] if (len(entry_members[t]) < 10) else entry_members[t][:10] + ["..."]

    if (False):
        print "\nentry_librarians:"
        for t in sorted([key for key in entry_librarians.keys()]):
            print "'{}'".format(t), len(entry_librarians[t]), entry_librarians[t] if (len(entry_librarians[t]) < 10) else entry_librarians[t][:10] + ["..."]

    if (False):
        print "\ncirculations:"
        for i, circulation in enumerate(circulations):
            print i, circulation

    return {
        # 'raw_entries': raw_entries,
#        'entries': entries,
#        'duplicates': duplicate_entries,
#        'locations': entry_locations,
#        'types': entry_types,
#        'authors': entry_authors,
#        'coauthors': entry_coauthors,
#        'publishers': entry_publishers,
#        'members': entry_members,
#        'librarians': entry_librarians,
#        'circulations': circulations,
    }


def import_library_magazines_xls(xlrdworkbook):
    """Import workflow:
    For XLS sheet named 'Magazines'

    """
    logger = logging.getLogger("import_library_books_xls.import_library_books_xls")
    logger.setLevel(logging.DEBUG)

    sheet = xlrdworkbook.sheet_by_name('Magazines')

    entries = []
    circulations = []

    entry_numbers = {}
    duplicate_entries = []

    entry_titles = {}
    entry_locations = {}
    entry_discard = {}
    entry_year = {}
    entry_month = {}
    entry_volume = {}
    entry_volnum = {}

    entry_members = {}
    entry_librarians = {}

    for i in range(1,sheet.nrows):
        entry = MagazineEntry.process_row_data(sheet, i, xlrdworkbook)

        # sanity / type check cell data
        
        if not entry:
            continue

        #logger.debug(entry)

        if (entry['number']) in entry_numbers:
            #logger.debug(" ".join((str(x) for x in ( "duplicate entry:", i, entry))))
            duplicate_entries.append([entry['number'], [entry_numbers[entry['number']], i]])
            continue
        else:
            entry_numbers[entry['number']]=i

        entry_titles[entry['title']] = \
            entry_titles.get(entry['title'], []) + [entry['number']]

        entry_locations[entry['location']] = \
            entry_locations.get(entry['location'], []) + [entry['number']]
            
        entry_discard[entry['discard']] = \
            entry_discard.get(entry['discard'], []) + [entry['number']]

        entry_year[entry['year']] = \
            entry_year.get(entry['year'], []) + [entry['number']]

        entry_month[entry['month']] = \
            entry_month.get(entry['month'], []) + [entry['number']]

        entry_volume[entry['volume']] = \
            entry_volume.get(entry['volume'], []) + [entry['number']]

        entry_volnum[entry['volnum']] = \
            entry_volnum.get(entry['volnum'], []) + [entry['number']]

        if ((entry['number']!='') and (entry['title']!='')):
            entries.append(entry)

    if (True):
        print "\nentry_title:"
        for t in sorted([key for key in entry_titles.keys()]):
            print "'{}'".format(t), len(entry_titles[t]), entry_titles[t] if (len(entry_titles[t]) < 10) else entry_titles[t][:10] + ["..."]

    if (True):
        print "\nentry_locations:"
        for t in sorted([key for key in entry_locations.keys()]):
            print "'{}'".format(t), len(entry_locations[t]), entry_locations[t] if (len(entry_locations[t]) < 10) else entry_locations[t][:10] + ["..."]

    if (True):
        print "\nentry_discard:"
        for t in sorted([key for key in entry_discard.keys()]):
            print "'{}'".format(t), len(entry_discard[t]), entry_discard[t] if (len(entry_discard[t]) < 10) else entry_discard[t][:10] + ["..."]

    if (True):
        print "\nentry_discard:"
        for t in sorted([key for key in entry_discard.keys()]):
            print "'{}'".format(t), len(entry_discard[t]), entry_discard[t] if (len(entry_discard[t]) < 10) else entry_discard[t][:10] + ["..."]

    if (True):
        print "\nentry_month:"
        for t in sorted([key for key in entry_month.keys()]):
            print "'{}'".format(t), len(entry_month[t]), entry_month[t] if (len(entry_month[t]) < 10) else entry_month[t][:10] + ["..."]

    return {
        'locations': entry_locations,
#        'discard': entry_discard,
#        'year': entry_year,
#        'month': entry_month,
#        'volume': entry_volume,
#        'volnum': entry_volnum,
    }


