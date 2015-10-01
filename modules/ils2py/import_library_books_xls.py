#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import re
import datetime

import logging

"""

Modules for importing the library from xls spreadsheet

xls spreadsheet row  -->  entry

RawEntry
* Invalid Data - TEXT
** is not XL_CELL_TEXT
** is XL_CELL_EMPTY

* Invalid Data - DATE
* Invalid Data - NUMBER


"""


logger = logging.getLogger("web2py.app.ils2py.import_library_books_xls")
logger.setLevel(logging.DEBUG)

class RawEntryDataType:
    def __init__(self, column_name, index, datatype):
        self.column_name = column_name
        self.index=index
        self.datatype=datatype

class RawEntry:
    ENTRY_FORMAT=[]

    def extract_data(self, index, row, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".__init__()")
        logger.setLevel(logging.DEBUG)

        self.row = row
        self.entry={}

        for n,i,t in self.ENTRY_FORMAT:
            cell = row[i]
            if t=='TEXT':
                if (cell.ctype == xlrd.XL_CELL_TEXT):
                    self.entry[n] = cell.value.encode('utf-8')
                    #logger.info(str(i)+":"+" ".join((str(x) for x in (n, cell, index, row))))
                elif (cell.ctype == xlrd.XL_CELL_EMPTY):
                    #logger.debug(" ".join((str(x) for x in (n, cell, index, row))))
                    self.entry[n] = None
                else:
                    #logger.debug(" ".join((str(x) for x in (n, cell, index, row))))
                    self.entry[n] = None
            elif t=='NUMBER':
                if (cell.ctype == xlrd.XL_CELL_NUMBER):
                    self.entry[n] = int(cell.value)
                    #logger.info(str(i)+":"+" ".join((str(x) for x in (n, cell, index, row))))
                elif (cell.ctype == xlrd.XL_CELL_EMPTY):
                    self.entry[n] = None
                else:
                    logger.debug("ctype != xlrd.XL_CELL_NUMBER:" + str(cell.ctype))
                    logger.debug(" ".join((str(x) for x in (n, cell, index, row))))
                    self.entry[n] = None
            elif t=='DATE':
                if (cell.ctype == xlrd.XL_CELL_DATE and cell.value != ''):
                    self.entry[n] = datetime.datetime(*xlrd.xldate_as_tuple(cell.value, xlrdworkbook.datemode))
                    #logger.info(str(i)+":"+" ".join((str(x) for x in (n, cell, index, row))))
                elif (cell.ctype == xlrd.XL_CELL_TEXT):
                    if (cell.value == '' or cell.value == ' '):
                        self.entry[n] = None
                    else:
                        logger.debug(" ".join((str(x) for x in (n, cell, index, row))))
                        self.entry[n] = None
                elif (cell.ctype == xlrd.XL_CELL_EMPTY):
                    self.entry[n] = None
                else:
                    logger.debug("ctype != xlrd.XL_CELL_DATE:" + str(cell.ctype))
                    logger.debug(" ".join((str(x) for x in (n, cell, index, row))))
                    self.entry[n] = None


class BookEntry(RawEntry):
    """
    circulation info:
            ('CK_OUT',    0, 'DATE'),
            ('DUE',       1, 'DATE'),
            ('WHO',       2, 'TEXT'),
            ('LIB',       3, 'TEXT'),
            ('RETURN',    4, 'DATE'),
            ('ENTERED',  14, 'TEXT'),
            ('ISBN',     15, 'TEXT'),
            ('DONOR',    16, 'TEXT'),
            ('MSRP',     17, 'TEXT'),

    item info:
            ('LOCATION',  5, 'TEXT'),
            ('NUMBER',    7, 'NUMBER'),

    bibio info:
            ('TYPE',      6, 'TEXT'),
            ('TITLE',     8, 'TEXT'),
            ('AUTHOR',    9, 'TEXT'),
            ('COAUTHOR', 10, 'TEXT'),
            ('COMMENTS', 11, 'TEXT'),
            ('PUBLISHER',12, 'TEXT'),
            ('SERIES',   13, 'TEXT'),
    ]

    """

    row_index_CK_OUT    =  0;  str_CK_OUT    = 'CK_OUT'
    row_index_DUE       =  1;  str_DUE       = 'DUE'
    row_index_WHO       =  2;  str_WHO       = 'WHO'
    row_index_LIB       =  3;  str_LIB       = 'LIB'
    row_index_RETURN    =  4;  str_RETURN    = 'RETURN'
    row_index_LOCATION  =  5;  str_LOCATION  = 'LOCATION'
    row_index_TYPE      =  6;  str_TYPE      = 'TYPE'
    row_index_NUMBER    =  7;  str_NUMBER    = 'NUMBER'
    row_index_TITLE     =  8;  str_TITLE     = 'TITLE'
    row_index_AUTHOR    =  9;  str_AUTHOR    = 'AUTHOR'
    row_index_COAUTHOR  = 10;  str_COAUTHOR  = 'COAUTHOR'
    row_index_COMMENTS  = 11;  str_COMMENTS  = 'COMMENTS'
    row_index_PUBLISHER = 12;  str_PUBLISHER = 'PUBLISHER'
    row_index_SERIES    = 13;  str_SERIES    = 'SERIES'
    row_index_ENTERED   = 14;  str_ENTERED   = 'ENTERED'
    row_index_ISBN      = 15;  str_ISBN      = 'ISBN'
    row_index_DONOR     = 16;  str_DONOR     = 'Donor'
    row_index_MSRP      = 18;  str_MSRP      = 'MSRP'

    ENTRY_FORMAT=[
            ('CK_OUT',    0, 'DATE'),
            ('DUE',       1, 'DATE'),
            ('WHO',       2, 'TEXT'),
            ('LIB',       3, 'TEXT'),
            ('RETURN',    4, 'DATE'),
            ('LOCATION',  5, 'TEXT'),
            ('TYPE',      6, 'TEXT'),
            ('NUMBER',    7, 'NUMBER'),
            ('TITLE',     8, 'TEXT'),
            ('AUTHOR',    9, 'TEXT'),
            ('COAUTHOR', 10, 'TEXT'),
            ('COMMENTS', 11, 'TEXT'),
            ('PUBLISHER',12, 'TEXT'),
            ('SERIES',   13, 'TEXT'),
            ('ENTERED',  14, 'TEXT'),
            ('ISBN',     15, 'TEXT'),
            ('DONOR',    16, 'TEXT'),
            ('MSRP',     17, 'TEXT'),
    ]

    def __init__(self, index, row, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".__init__()")
        logger.setLevel(logging.DEBUG)
        self.extract_data(index, row, xlrdworkbook)

    @staticmethod
    def check_sheet(sheet):
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
        logger = logging.getLogger("import_library_books_xls.BookEntry.check_sheet()")
        logger.setLevel(logging.DEBUG)
        #sheet = xlrdworkbook.sheet_by_index(0)
        logger.debug("Sheet:{0}".format(sheet.name))
        for i, column_name in enumerate(column_names):
            c = sheet.cell(0, i)
            logger.debug("Checking Headers {0} - '{1}' == '{2}'".format(i, column_name, c.value))
            if c.value != column_name:
                return False
        return True


    @staticmethod
    def process_row_data(sheet, i, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls.BookEntry.process_row_data({})".format(i))
        logger.setLevel(logging.DEBUG)
    
        row = sheet.row(i)
        entry = {}
        # sanity / type check cell data
    
        if ((row[BookEntry.row_index_NUMBER].ctype == xlrd.XL_CELL_EMPTY) and
            (row[BookEntry.row_index_LOCATION].ctype == xlrd.XL_CELL_EMPTY) and
            (row[BookEntry.row_index_TYPE].ctype == xlrd.XL_CELL_EMPTY) and
            (row[BookEntry.row_index_TITLE].ctype == xlrd.XL_CELL_EMPTY)):
            return {}
    
        b = BookEntry(i, row, xlrdworkbook)
    
        entry['index'] = i
        entry['number'] = b.entry.get('NUMBER', None)
        entry['location'] = b.entry.get('LOCATION', None)
        entry['type'] = b.entry.get('TYPE', None)
        entry['title'] = b.entry.get('TITLE', None)
        entry['author'] = b.entry.get('AUTHOR', None)
        entry['coauthor'] = b.entry.get('COAUTHOR', None)
        entry['comments'] = b.entry.get('COMMENTS', None)
        entry['publisher'] = b.entry.get('PUBLISHER', None)
        entry['series'] = b.entry.get('SERIES', None)
        entry['OUT'] = b.entry.get('CK_OUT', None)
        entry['DUE'] = b.entry.get('DUE', None)
        entry['WHO'] = b.entry.get('WHO', None)
        entry['RETURN'] = b.entry.get('RETURN', None)
    
        return entry

class MagazineEntry(RawEntry):
    """
    circulation info
            ('OUT',        0, 'DATE'),
            ('DUE',        1, 'DATE'),
            ('WHO',        2, 'TEXT'),
            ('LIB',        3, 'TEXT'),
            ('RETURNED',   4, 'DATE'),

    item info:
            ('Discard',    5, 'TEXT'),
            ('LOCATION',   6, 'TEXT'),
            ('NUMBER',     7, 'NUMBER'),
            ('COMMENTS1', 14, 'TEXT'),

    bibio info:
            ('TITLE',      8, 'TEXT'),
            ('YEAR',       9, 'TEXT'),
            ('MONTH',     10, 'TEXT'),
            ('VOLUME',    11, 'TEXT'),
            ('VOLNUM',    12, 'TEXT'),
            ('WHOLE',     13, 'TEXT'),
            ('ENTERED',   15, 'DATE'),
            #('BLANK1',    16, 'TEXT'),
            ('COMMENTS2', 17, 'TEXT'),
            #('BLANK2',    18, 'TEXT'),
            #('BLANK3',    19, 'TEXT'),
    """

    row_index_OUT        = 0
    row_index_DUE        = 1
    row_index_WHO        = 2
    row_index_LIB        = 3
    row_index_RETURNED   = 4
    row_index_Discard    = 5
    row_index_LOCATION   = 6
    row_index_NUMBER     = 7
    row_index_TITLE      = 8
    row_index_YEAR       = 9
    row_index_MONTH     = 10
    row_index_VOLUME    = 11
    row_index_VOLNUM    = 12
    row_index_WHOLE     = 13
    row_index_COMMENTS1 = 14
    row_index_ENTERED   = 15
    row_index_BLANK1    = 16
    row_index_COMMENTS2 = 17
    row_index_BLANK2    = 18
    row_index_BLANK3    = 19
    ENTRY_FORMAT = [
            ('OUT',        0, 'DATE'),
            ('DUE',        1, 'DATE'),
            ('WHO',        2, 'TEXT'),
            ('LIB',        3, 'TEXT'),
            ('RETURNED',   4, 'DATE'),
            ('Discard',    5, 'TEXT'),
            ('LOCATION',   6, 'TEXT'),
            ('NUMBER',     7, 'NUMBER'),
            ('TITLE',      8, 'TEXT'),
            ('YEAR',       9, 'TEXT'),
            ('MONTH',     10, 'TEXT'),
            ('VOLUME',    11, 'TEXT'),
            ('VOLNUM',    12, 'TEXT'),
            ('WHOLE',     13, 'TEXT'),
            ('COMMENTS1', 14, 'TEXT'),
            ('ENTERED',   15, 'DATE'),
            #('BLANK1',    16, 'TEXT'),
            ('COMMENTS2', 17, 'TEXT'),
            #('BLANK2',    18, 'TEXT'),
            #('BLANK3',    19, 'TEXT'),
    ]

    def __init__(self, index, row, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".__init__()")
        logger.setLevel(logging.DEBUG)
        self.extract_data(index, row, xlrdworkbook)

    @staticmethod
    def process_row_data(sheet, i, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls.MagazineEntry.process_row_data({})".format(i))
        logger.setLevel(logging.DEBUG)
    
        row = sheet.row(i)
        entry = {}
        # sanity / type check cell data
    
        #if ((row[MagazineEntry.row_index_NUMBER].ctype == xlrd.XL_CELL_EMPTY) and
        #    (row[MagazineEntry.row_index_LOCATION].ctype == xlrd.XL_CELL_EMPTY)):
        #    return {}
    
        b = MagazineEntry(i, row, xlrdworkbook)
    
        entry['index'] = i
        entry['number'] = b.entry.get('NUMBER', None)
        entry['location'] = b.entry.get('LOCATION', None)
        entry['discard'] = b.entry.get('Discard', None)
        entry['title'] = b.entry.get('TITLE', None)
        entry['year'] = b.entry.get('YEAR', None)
        entry['month'] = b.entry.get('MONTH', None)
        entry['volume'] = b.entry.get('VOLUME', None)
        entry['volnum'] = b.entry.get('VOLNUM', None)
        entry['completed'] = b.entry.get('WHOLE', None)
        entry['comments1'] = b.entry.get('COMMENTS1', None)
        entry['entered'] = b.entry.get('ENTERED', None)
        entry['comments2'] = b.entry.get('COMMENTS2', None)
        return entry

class MemberEntry(RawEntry):
    ENTRY_FORMAT = [
            ('NAME', 0, 'TEXT'),
    ]

    def __init__(self, index, row, xlrdworkbook):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__+".__init__()")
        logger.setLevel(logging.DEBUG)

        self.row = row
        self.entry={}

        self.extract_data(index, row, xlrdworkbook)

class VideoEntry(RawEntry):
    ENTRY_FORMAT = [
            ('TITLE', 0, 'TEXT'),
            ('LC', 0, 'TEXT'),
            ('!', 0, 'TEXT'),
            ('LOCATION', 0, 'TEXT'),
            ('Borrowed By', 0, 'TEXT'),
            ('Checkout', 0, 'TEXT'),
            ('Due', 0, 'DATE'),
            ('Returned', 0, 'DATE'),
            ('Libr', 0, 'DATE'),
            ('COMMENT1', 0, 'DATE'),
            ('COMMENT2', 0, 'DATE'),
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

    for i in range(1,sheet.nrows):
        entry = BookEntry.process_row_data(sheet, i, xlrdworkbook)

        # sanity / type check cell data
        
        if not entry:
            continue

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
