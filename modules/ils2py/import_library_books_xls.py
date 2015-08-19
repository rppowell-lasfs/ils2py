#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import re
import datetime

import logging

logger = logging.getLogger("web2py.app.ils2py.import_library_books_xls")
logger.setLevel(logging.DEBUG)

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
    'SERIES'
    'ENTERED'
]

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


def check_headers(xlrdworkbook):
    logger.debug("check_headers()")
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
    logger.debug("import_library_books_xls()")

    sheet = xlrdworkbook.sheet_by_name('Books')
    logger.debug("Importing Sheet:{0}".format(sheet.name))

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
        row = sheet.row(i)

        entry = {}
        circulation = {}

        # sanity / type check cell data

        if ((row[row_index_NUMBER].ctype == xlrd.XL_CELL_EMPTY) and
            (row[row_index_LOCATION].ctype == xlrd.XL_CELL_EMPTY) and
            (row[row_index_TYPE].ctype == xlrd.XL_CELL_EMPTY) and
            (row[row_index_TITLE].ctype == xlrd.XL_CELL_EMPTY)):
            continue

        number_cell = row[row_index_NUMBER]
        if (number_cell.ctype != xlrd.XL_CELL_NUMBER):
            logger.debug("number_cell.ctype != xlrd.XL_CELL_NUMBER; number_cell.ctype=" + str(number_cell.ctype))
            logger.debug(" ".join((str(x) for x in (number_cell, i, row))))
            continue
        else:
            entry['number'] = int(number_cell.value)

        location_cell = row[row_index_LOCATION]
        if (location_cell.ctype != xlrd.XL_CELL_TEXT):
            logger.debug("location_cell.ctype != xlrd.XL_CELL_TEXT; location_cell.ctype=" + str(location_cell.ctype))
            logger.debug(" ".join((str(x) for x in (location_cell, i, row))))
            break
        elif (location_cell.ctype == xlrd.XL_CELL_EMPTY):
            logger.debug("location_cell.ctype == xlrd.XL_CELL_EMPTY; location_cell.ctype=" + str(location_cell.ctype))
            logger.debug(" ".join((str(x) for x in (location_cell, i, row))))
            break
        else:
            entry['location'] = location_cell.value.encode('utf-8')

        type_cell = row[row_index_TYPE]
        if not (type_cell.ctype == xlrd.XL_CELL_EMPTY or type_cell.ctype == xlrd.XL_CELL_TEXT):
            logger.debug("!(type_cell.ctype == xlrd.XL_CELL_EMPTY or xlrd.XL_CELL_TEXT; type_cell.ctype=" + str(type_cell.ctype))
            logger.debug(" ".join((str(x) for x in (type_cell, i, row))))
            break
        else:
            entry['type'] = type_cell.value.encode('utf-8')

        title_cell = row[row_index_TITLE]
        if (title_cell.ctype != xlrd.XL_CELL_TEXT):
            logger.debug("title_cell.ctype != xlrd.XL_CELL_TEXT; title_cell.ctype=" + str(title_cell.ctype))
            logger.debug(" ".join((str(x) for x in (title_cell, i, row))))
            break
        else:
            entry['title'] = title_cell.value.encode('utf-8')


        if (False):
            print [i, number_cell.value, location_cell.value, type_cell.value, title_cell.value]


        author_cell = row[row_index_AUTHOR]
        entry['author'] = author_cell.value.encode('utf-8') if author_cell.value != '' else ''

        coauthor_cell = row[row_index_COAUTHOR]
        entry['coauthor'] = coauthor_cell.value.encode('utf-8') if coauthor_cell.value != '' else ''

        publisher_cell = row[row_index_PUBLISHER]
        if (publisher_cell.ctype == xlrd.XL_CELL_TEXT):
            entry['publisher'] = publisher_cell.value.encode('utf-8') if publisher_cell.value != '' else ''
        elif (publisher_cell.ctype == xlrd.XL_CELL_NUMBER):
            entry['publisher'] = str(publisher_cell.value)
        elif (publisher_cell.ctype == xlrd.XL_CELL_EMPTY):
            entry['publisher'] = ''
        else:
            print 'publisher', i, publisher_cell
            entry['publisher'] = ''


        if (entry['number']) in entry_numbers:
            logger.debug(" ".join((str(x) for x in (
                "duplicate entry:", entry['number'], [i, number_cell.value, location_cell.value, type_cell.value, title_cell.value]
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
        entry_authors[entry['author']] = \
            entry_authors.get(entry['author'], []) + [entry['number']]

        entry_coauthors[entry['coauthor']] = \
            entry_coauthors.get(entry['coauthor'], []) + [entry['number']]

        entry_publishers[entry['publisher']] = \
            entry_publishers.get(entry['publisher'], []) + [entry['publisher']]

        # circurlation info
        out_cell = row[row_index_CK_OUT]
        due_cell = row[row_index_DUE]
        who_cell = row[row_index_WHO]
        lib_cell = row[row_index_LIB]
        return_cell = row[row_index_RETURN]

        if (out_cell.ctype == xlrd.XL_CELL_DATE and out_cell.value != ''):
            circulation['OUT'] = datetime.datetime(*xlrd.xldate_as_tuple(out_cell.value, xlrdworkbook.datemode))
        elif (out_cell.ctype == xlrd.XL_CELL_TEXT and out_cell.value == ''):
            pass
        #elif (out_cell.ctype == xlrd.XL_CELL_TEXT and out_cell.value == ' '):
        #    pass
        elif (out_cell.ctype == xlrd.XL_CELL_EMPTY):
            pass
        else:
            print 'OUT', i, out_cell, row
    
        if (due_cell.ctype == xlrd.XL_CELL_DATE and due_cell.value != ''):
            circulation['DUE'] = datetime.datetime(*xlrd.xldate_as_tuple(due_cell.value, xlrdworkbook.datemode))
        elif (due_cell.ctype == xlrd.XL_CELL_TEXT and due_cell.value == ''):
            pass
        #elif (due_cell.ctype == xlrd.XL_CELL_TEXT and due_cell.value == ' '):
        #    pass
        elif (due_cell.ctype == xlrd.XL_CELL_EMPTY):
            pass
        else:
            print 'DUE', i, due_cell, row

        if (who_cell.ctype == xlrd.XL_CELL_TEXT and who_cell.value != ''):
            circulation['WHO'] = who_cell.value.encode('utf-8')
        elif (who_cell.ctype == xlrd.XL_CELL_EMPTY):
            pass
        else:
            print 'WHO', i, who_cell, row

        if 'WHO' in circulation:
            entry_members[circulation['WHO']] = \
                entry_members.get(circulation['WHO'], []) + [entry['number']]

        if (lib_cell.ctype == xlrd.XL_CELL_TEXT and lib_cell.value != ''):
            circulation['LIB'] = lib_cell.value.encode('utf-8')
        elif (lib_cell.ctype == xlrd.XL_CELL_EMPTY):
            pass
        else:
            print 'LIB', i, lib_cell

        if 'LIB' in circulation:
            entry_librarians[circulation['LIB']] = \
                entry_librarians.get(circulation['LIB'], []) + [entry['number']]

        if (return_cell.ctype == xlrd.XL_CELL_DATE and return_cell.value != ''):
            circulation['RETURN'] = datetime.datetime(*xlrd.xldate_as_tuple(return_cell.value, xlrdworkbook.datemode))
        elif (return_cell.ctype == xlrd.XL_CELL_TEXT and return_cell.value == ''):
            pass
        elif (return_cell.ctype == xlrd.XL_CELL_EMPTY):
            pass
        else:
            print 'RETURN', i, return_cell

        if ('OUT' in circulation and 'DUE' in circulation and 'WHO' in circulation):
            if ('RETURN' in circulation):
                pass
            else:
                print "cirulation:", i, row
                circulation['number'] = entry['number']
                circulations.append(circulation)

    if (False):
        print "\nentries:"
        for i, entry in enumerate(entries):
            print i, entry

    if (False):
        print "\nduplicate_entries:", duplicate_entries
        for k in duplicate_entries:
            i = k[0];
            i1 = k[1][0]
            i2 = k[1][1]
            print "\t".join((str(i), str(i1+1), sheet.row(i1)[row_index_AUTHOR].value, sheet.row(i1)[row_index_TITLE].value))
            print "\t".join((str(i), str(i2+1), sheet.row(i2)[row_index_AUTHOR].value, sheet.row(i2)[row_index_TITLE].value))

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


    return {
        # 'raw_entries': raw_entries,
        'entries': entries,
        'duplicates': duplicate_entries,
        'locations': entry_locations,
        'types': entry_types,
        'authors': entry_authors,
        'coauthors': entry_coauthors,
        'publishers': entry_publishers,
        'members': entry_members,
        'librarians': entry_librarians,
        'circulations': circulations,
    }

