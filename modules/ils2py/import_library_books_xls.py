#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import re
import datetime

import logging

logger = logging.getLogger("web2py.app.ils2py.import_library_books_xls")
logger.setLevel(logging.DEBUG)

class BookEntry:

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
    """

    logger = logging.getLogger("web2py.app.ils2py.import_library_books_xls")
    logger.setLevel(logging.DEBUG)

    def __init__(self, index, row, xlrdworkbook):

        self.index = index
        self.row = row
        self.xlrdworkbook = xlrdworkbook

    def read_Book(self):
        logger = logging.getLogger("import_library_books_xls."+self.__class__.__name__)
        logger.setLevel(logging.DEBUG)

        self.entry={}
        try:
            self.entry['index'] = self.index
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
        try:
            self.entry['CKOUT'] = self.read_Book_CKOUT()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['CKOUT'] = None
        try:
            self.entry['DUE'] = self.read_Book_DUE()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['DUE'] = None
        try:
            self.entry['WHO'] = self.read_Book_WHO()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['WHO'] = None
        try:
            self.entry['LIB'] = self.read_Book_LIB()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['LIB'] = None
        try:
            self.entry['RETURN'] = self.read_Book_RETURN()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['RETURN'] = None
        try:
            self.entry['LOCATION'] = self.read_Book_LOCATION()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['LOCATION'] = None
        try:
            self.entry['TYPE'] = self.read_Book_TYPE()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['TYPE'] = None
        try:
            self.entry['NUMBER'] = self.read_Book_NUMBER()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['NUMBER'] = None
        try:
            self.entry['TITLE'] = self.read_Book_TITLE()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['TITLE'] = None
        try:
            self.entry['AUTHOR'] = self.read_Book_AUTHOR()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['AUTHOR'] = None
        try:
            self.entry['COAUTHOR'] = self.read_Book_COAUTHOR()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['COAUTHOR'] = None
        try:
            self.entry['COMMENTS'] = self.read_Book_COMMENTS()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['COMMENTS'] = None
        try:
            self.entry['PUBLISHER'] = self.read_Book_PUBLISHER()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['PUBLISHER'] = None
        try:
            self.entry['SERIES'] = self.read_Book_SERIES()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['SERIES'] = None
        try:
            self.entry['ENTERED'] = self.read_Book_ENTERED()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['ENTERED'] = None
        try:
            self.entry['ISBN'] = self.read_Book_ISBN()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['ISBN'] = None
        try:
            self.entry['DONOR'] = self.read_Book_DONOR()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['DONOR'] = None
        try:
            self.entry['MSRP'] = self.read_Book_MSRP()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['MSRP'] = None

        return self.entry

    def read_Book_CKOUT(self):
        self.CKOUT_cell = self.row[0]
        if (self.CKOUT_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.cell.value])
            self.CKOUT_data=None
    
        elif (self.CKOUT_cell.ctype == xlrd.XL_CELL_DATE and self.CKOUT_cell.value != ''):
            self.CKOUT_data=datetime.datetime(*xlrd.xldate_as_tuple(self.CKOUT_cell.value, self.xlrdworkbook.datemode))
    
        elif (self.CKOUT_cell.ctype == xlrd.XL_CELL_TEXT):
            if (self.CKOUT_cell.value == '' or self.CKOUT_cell.value == ' '):
                self.CKOUT_data=None
            else:
                #raise Exception(str(self.cell))
                self.CKOUT_data=None
    
        elif (self.CKOUT_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.CKOUT_data=None
    
        else:
            self.CKOUT_data=None
            raise Exception(str(self.cell))

        return self.CKOUT_data

    def read_Book_DUE(self):
        self.DUE_cell = self.row[1]
        if (self.DUE_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.cell.value])
            self.DUE_data=None
    
        elif (self.DUE_cell.ctype == xlrd.XL_CELL_DATE and self.DUE_cell.value != ''):
            self.DUE_data=datetime.datetime(*xlrd.xldate_as_tuple(self.DUE_cell.value, self.xlrdworkbook.datemode))
    
        elif (self.DUE_cell.ctype == xlrd.XL_CELL_TEXT):
            if (self.DUE_cell.value == '' or self.DUE_cell.value == ' '):
                self.DUE_data=None
            else:
                #raise Exception(str(self.cell))
                self.DUE_data=None
    
        elif (self.DUE_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.DUE_data=None
    
        else:
            self.DUE_data = None
            raise Exception(str(self.DUE_cell))

        return self.DUE_data

    def read_Book_WHO(self):
        self.WHO_cell = self.row[2]
        if (self.WHO_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.cell.value])
            self.WHO_data = None

        elif (self.WHO_cell.ctype == xlrd.XL_CELL_TEXT):
            self.WHO_data = self.WHO_cell.value.encode('utf-8').strip()

        elif (self.WHO_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.WHO_data = None

        else:
            self.WHO_data = None
            raise Exception(str(self.WHO_cell))

        return self.WHO_data

    def read_Book_LIB(self):
        self.LIB_cell = self.row[3]
        if (self.LIB_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.LIB_cell.value])
            self.LIB_data = None

        elif (self.LIB_cell.ctype == xlrd.XL_CELL_TEXT):
            self.LIB_data = self.LIB_cell.value.encode('utf-8').strip()

        elif (self.LIB_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.LIB_data = str(self.LIB_cell.value)

        elif (self.LIB_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.LIB_data = None

        else:
            self.LIB_data = None
            raise Exception(str(self.LIB_cell))

        return self.LIB_data

    def read_Book_RETURN(self):
        self.RETURN_cell = self.row[4]
        if (self.RETURN_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.cell.value])
            self.RETURN_data=None
    
        elif (self.RETURN_cell.ctype == xlrd.XL_CELL_DATE and self.RETURN_cell.value != ''):
            self.RETURN_data=datetime.datetime(*xlrd.xldate_as_tuple(self.RETURN_cell.value, self.xlrdworkbook.datemode))
    
        elif (self.RETURN_cell.ctype == xlrd.XL_CELL_TEXT):
            if (self.RETURN_cell.value == '' or self.RETURN_cell.value == ' '):
                self.RETURN_data=None
            else:
                #raise Exception(str(self.cell))
                self.RETURN_data=None
    
        elif (self.RETURN_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.RETURN_data=None
    
        else:
            raise Exception(str(self.RETURN_cell))

        return self.RETURN_data


    def read_Book_LOCATION(self):
        self.LOCATION_cell = self.row[5]
        if (self.LOCATION_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.LOCATION_cell.value])
            self.LOCATION_data = None

        elif (self.LOCATION_cell.ctype == xlrd.XL_CELL_TEXT):
            self.LOCATION_data = self.LOCATION_cell.value.encode('utf-8').strip()

        elif (self.LOCATION_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.LOCATION_data = None

        else:
            self.LOCATION_data = None
            raise Exception(str(self.LOCATION_cell))

        return self.LOCATION_data

    def read_Book_TYPE(self):
        self.TYPE_cell = self.row[6]
        if (self.TYPE_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.TYPE_cell.value])
            self.TYPE_data = None

        elif (self.TYPE_cell.ctype == xlrd.XL_CELL_TEXT):
            self.TYPE_data = self.TYPE_cell.value.encode('utf-8').strip()

        elif (self.TYPE_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.TYPE_data = None

        else:
            self.TYPE_data = None
            raise Exception(str(self.TYPE_cell))

        return self.TYPE_data

    def read_Book_NUMBER(self):
        self.NUMBER_cell = self.row[7]

        if (self.NUMBER_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.NUMBER_cell.value])
            self.NUMBER_data = None

        elif (self.NUMBER_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.NUMBER_data = int(self.NUMBER_cell.value)

        elif (self.NUMBER_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.NUMBER_data = None

        else:
            self.NUMBER_data = None
            raise Exception(str(self.NUMBER_cell))

        return self.NUMBER_data

    def read_Book_TITLE(self):
        self.TITLE_cell = self.row[8]
        if (self.TITLE_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.TITLE_cell.value])
            self.TITLE_data = None

        elif (self.TITLE_cell.ctype == xlrd.XL_CELL_TEXT):
            self.TITLE_data = self.TITLE_cell.value.encode('utf-8').strip()

        elif (self.TITLE_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.TITLE_data = None

        else:
            self.TITLE_data = None
            raise Exception(str(self.TITLE_cell))

        return self.TITLE_data

    def read_Book_AUTHOR(self):
        self.AUTHOR_cell = self.row[9]
        if (self.AUTHOR_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.AUTHOR_cell.value])
            self.AUTHOR_data = None

        elif (self.AUTHOR_cell.ctype == xlrd.XL_CELL_TEXT):
            self.AUTHOR_data = self.AUTHOR_cell.value.encode('utf-8').strip()

        elif (self.AUTHOR_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.AUTHOR_data = None

        else:
            self.AUTHOR_data = None
            raise Exception(str(self.AUTHOR_cell))

        return self.AUTHOR_data

    def read_Book_COAUTHOR(self):
        self.COAUTHOR_cell = self.row[10]
        if (self.COAUTHOR_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.COAUTHOR_cell.value])
            self.COAUTHOR_data = None

        elif (self.COAUTHOR_cell.ctype == xlrd.XL_CELL_TEXT):
            self.COAUTHOR_data = self.COAUTHOR_cell.value.encode('utf-8').strip()

        elif (self.COAUTHOR_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.COAUTHOR_data = None

        else:
            self.COAUTHOR_data = None
            raise Exception(str(self.COAUTHOR_cell))

        return self.COAUTHOR_data

    def read_Book_COMMENTS(self):
        self.COMMENTS_cell = self.row[11]
        if (self.COMMENTS_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.COMMENTS_cell.value])
            self.COMMENTS_data = None

        elif (self.COMMENTS_cell.ctype == xlrd.XL_CELL_TEXT):
            self.COMMENTS_data = self.COMMENTS_cell.value.encode('utf-8').strip()

        elif (self.COMMENTS_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.COMMENTS_data = str(self.COMMENTS_cell.value)

        elif (self.COMMENTS_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.COMMENTS_data = None

        else:
            self.COMMENTS_data = None
            raise Exception(str(self.COMMENTS_cell))

        return self.COMMENTS_data

    def read_Book_PUBLISHER(self):
        self.PUBLISHER_cell = self.row[12]
        if (self.PUBLISHER_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.PUBLISHER_cell.value])
            self.PUBLISHER_data = None

        elif (self.PUBLISHER_cell.ctype == xlrd.XL_CELL_TEXT):
            self.PUBLISHER_data = self.PUBLISHER_cell.value.encode('utf-8').strip()

        elif (self.PUBLISHER_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.PUBLISHER_data = str(self.PUBLISHER_cell.value)

        elif (self.PUBLISHER_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.PUBLISHER_data = None

        else:
            self.PUBLISHER_data = None
            raise Exception(str(self.PUBLISHER_cell))

        return self.PUBLISHER_data

    def read_Book_SERIES(self):
        self.SERIES_cell = self.row[13]
        if (self.SERIES_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.SERIES_cell.value])
            self.SERIES_data = None

        elif (self.SERIES_cell.ctype == xlrd.XL_CELL_TEXT):
            self.SERIES_data = self.SERIES_cell.value.encode('utf-8').strip()

        elif (self.SERIES_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.SERIES_data = None

        else:
            self.SERIES_data = None
            raise Exception(str(self.SERIES_cell))

        return self.SERIES_data

    def read_Book_ENTERED(self):
        self.ENTERED_cell = self.row[14]
        if (self.ENTERED_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.ENTERED_cell.value])
            self.ENTERED_data=None
    
        elif (self.ENTERED_cell.ctype == xlrd.XL_CELL_DATE and self.ENTERED_cell.value != ''):
            self.ENTERED_data=datetime.datetime(*xlrd.xldate_as_tuple(self.ENTERED_cell.value, self.xlrdworkbook.datemode))
    
        elif (self.ENTERED_cell.ctype == xlrd.XL_CELL_TEXT):
            if (self.ENTERED_cell.value == '' or self.ENTERED_cell.value == ' '):
                self.ENTERED_data=None
            else:
                #raise Exception(str(self.ENTERED_cell))
                self.ENTERED_data=None
    
        elif (self.ENTERED_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.ENTERED_data=None
    
        else:
            self.ENTERED_data=None
            raise Exception(str(self.ENTERED_cell))

        return self.ENTERED_data

    def read_Book_ISBN(self):
        self.ISBN_cell = self.row[15]
        if (self.ISBN_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.ISBN_cell.value])
            self.ISBN_data = None

        elif (self.ISBN_cell.ctype == xlrd.XL_CELL_TEXT):
            self.ISBN_data = self.ISBN_cell.value.encode('utf-8').strip()

        elif (self.ISBN_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.ISBN_data = str(int(self.ISBN_cell.value))

        elif (self.ISBN_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.ISBN_data = None

        else:
            self.ISBN_data = None
            raise Exception(str(self.ISBN_cell))

        return self.ISBN_data

    def read_Book_DONOR(self):
        self.DONOR_cell = self.row[16]
        if (self.DONOR_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.DONOR_cell.value])
            self.DONOR_data = None

        elif (self.DONOR_cell.ctype == xlrd.XL_CELL_TEXT):
            self.DONOR_data = self.DONOR_cell.value.encode('utf-8').strip()

        elif (self.DONOR_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.DONOR_data = None

        else:
            self.DONOR_data = None
            raise Exception(str(self.DONOR_cell))

        return self.DONOR_data

    def read_Book_MSRP(self):
        self.MSRP_cell = self.row[18]

        if (self.MSRP_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.MSRP_cell.value])
            self.MSRP_data = None

        elif (self.MSRP_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.MSRP_data = str(self.MSRP_cell.value)

        elif (self.MSRP_cell.ctype == xlrd.XL_CELL_TEXT):
            self.MSRP_data = self.MSRP_cell.value.encode('utf-8').strip()

        elif (self.MSRP_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.MSRP_data = None

        else:
            self.MSRP_data = None
            raise Exception(str(self.MSRP_cell))

        return self.MSRP_data


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
    
        #if ((row[BookEntry.row_index_NUMBER].ctype == xlrd.XL_CELL_EMPTY) and
        #    (row[BookEntry.row_index_LOCATION].ctype == xlrd.XL_CELL_EMPTY) and
        #    (row[BookEntry.row_index_TYPE].ctype == xlrd.XL_CELL_EMPTY) and
        #    (row[BookEntry.row_index_TITLE].ctype == xlrd.XL_CELL_EMPTY)):
        #    return {}
    
        b = BookEntry(i, row, xlrdworkbook)
        entry=b.read_Book()

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

