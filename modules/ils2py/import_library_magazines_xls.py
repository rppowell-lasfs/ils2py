#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import re
import datetime

import logging

logger = logging.getLogger("web2py.app.ils2py.import_library_magazines_xls")
logger.setLevel(logging.DEBUG)

class MagazineEntry:
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
            RawEntryDataType('OUT',        0, RawEntryDataType.DATE),
            RawEntryDataType('DUE',        1, RawEntryDataType.DATE),
            RawEntryDataType('WHO',        2, RawEntryDataType.TEXT),
            RawEntryDataType('LIB',        3, RawEntryDataType.TEXT),
            RawEntryDataType('RETURNED',   4, RawEntryDataType.DATE),
            RawEntryDataType('Discard',    5, RawEntryDataType.TEXT),
            RawEntryDataType('LOCATION',   6, RawEntryDataType.TEXT),
            RawEntryDataType('NUMBER',     7, RawEntryDataType.NUMBER),
            RawEntryDataType('TITLE',      8, RawEntryDataType.TEXT),
            RawEntryDataType('YEAR',       9, RawEntryDataType.TEXT),
            RawEntryDataType('MONTH',     10, RawEntryDataType.TEXT),
            RawEntryDataType('VOLUME',    11, RawEntryDataType.TEXT),
            RawEntryDataType('VOLNUM',    12, RawEntryDataType.TEXT),
            RawEntryDataType('WHOLE',     13, RawEntryDataType.TEXT),
            RawEntryDataType('COMMENTS1', 14, RawEntryDataType.TEXT),
            RawEntryDataType('ENTERED',   15, RawEntryDataType.DATE),
            #('BLANK1',    16, 'TEXT'),
            RawEntryDataType('COMMENTS2', 17, RawEntryDataType.TEXT),
            #('BLANK2',    18, 'TEXT'),
            #('BLANK3',    19, 'TEXT'),
    ]
    """

    def __init__(self, index, row, xlrdworkbook):
        self.index = index
        self.row = row
        self.xlrdworkbook = xlrdworkbook

    def read_Magazine(self):
        self.entry={}
        self.entry['index'] = self.index
        try:
            self.entry['OUT'] = self.read_Magazine_OUT()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['OUT'] = None
        try:
            self.entry['DUE'] = self.read_Magazine_DUE()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['DUE'] = None
        try:
            self.entry['WHO'] = self.read_Magazine_WHO()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['WHO'] = None
        try:
            self.entry['LIB'] = self.read_Magazine_LIB()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['LIB'] = None
        try:
            self.entry['RETURNED'] = self.read_Magazine_RETURNED()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['RETURNED'] = None
        try:
            self.entry['Discard'] = self.read_Magazine_Discard()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['Discard'] = None
        try:
            self.entry['LOCATION'] = self.read_Magazine_LOCATION()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['LOCATION'] = None
        try:
            self.entry['NUMBER'] = self.read_Magazine_NUMBER()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['NUMBER'] = None
        try:
            self.entry['TITLE'] = self.read_Magazine_TITLE()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['TITLE'] = None
        try:
            self.entry['YEAR'] = self.read_Magazine_YEAR()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['YEAR'] = None
        try:
            self.entry['MONTH'] = self.read_Magazine_MONTH()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['MONTH'] = None
        try:
            self.entry['VOLUME'] = self.read_Magazine_VOLUME()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['VOLUME'] = None
        try:
            self.entry['VOLNUM'] = self.read_Magazine_VOLNUM()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['VOLNUM'] = None
        try:
            self.entry['WHOLE'] = self.read_Magazine_WHOLE()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['WHOLE'] = None
        try:
            self.entry['COMMENTS1'] = self.read_Magazine_COMMENTS1()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['COMMENTS'] = None
        try:
            self.entry['ENTERED'] = self.read_Magazine_ENTERED()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['ENTERED'] = None
        try:
            self.entry['COMMENTS2'] = self.read_Magazine_COMMENTS2()
        except Exception as e:
            logger.exception("Error processing entry {0}:{1}".format(self.index + 1, e))
            self.entry['COMMENTS2'] = None

        return self.entry

    def read_Magazine_OUT(self):
        self.OUT_cell = self.row[0]
        if (self.OUT_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.OUT_cell.value])
            self.OUT_data=None
    
        elif (self.OUT_cell.ctype == xlrd.XL_CELL_DATE and self.OUT_cell.value != ''):
            self.OUT_data=datetime.datetime(*xlrd.xldate_as_tuple(self.OUT_cell.value, self.xlrdworkbook.datemode))
    
        elif (self.OUT_cell.ctype == xlrd.XL_CELL_TEXT):
            if (self.OUT_cell.value == '' or self.OUT_cell.value == ' '):
                self.OUT_data=None
            else:
                #raise Exception(str(self.cell))
                self.OUT_data=None
    
        elif (self.OUT_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.OUT_data=None
    
        else:
            self.OUT_data=None
            raise Exception(str(self.cell))

        return self.OUT_data

    def read_Magazine_DUE(self):
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

    def read_Magazine_WHO(self):
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

    def read_Magazine_LIB(self):
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

    def read_Magazine_RETURNED(self):
        self.RETURNED_cell = self.row[4]
        if (self.RETURNED_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.cell.value])
            self.RETURNED_data=None
    
        elif (self.RETURNED_cell.ctype == xlrd.XL_CELL_DATE and self.RETURNED_cell.value != ''):
            self.RETURNED_data=datetime.datetime(*xlrd.xldate_as_tuple(self.RETURNED_cell.value, self.xlrdworkbook.datemode))
    
        elif (self.RETURNED_cell.ctype == xlrd.XL_CELL_TEXT):
            if (self.RETURNED_cell.value == '' or self.RETURNED_cell.value == ' '):
                self.RETURNED_data=None
            else:
                #raise Exception(str(self.cell))
                self.RETURNED_data=None
    
        elif (self.RETURNED_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.RETURNED_data=None
    
        else:
            raise Exception(str(self.RETURNED_cell))

        return self.RETURNED_data

    def read_Magazine_Discard(self):
        self.Discard_cell = self.row[5]
        if (self.Discard_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.cell.value])
            self.Discard_data=None
    
        elif (self.Discard_cell.ctype == xlrd.XL_CELL_DATE and self.Discard_cell.value != ''):
            self.Discard_data=datetime.datetime(*xlrd.xldate_as_tuple(self.Discard_cell.value, self.xlrdworkbook.datemode))
    
        elif (self.Discard_cell.ctype == xlrd.XL_CELL_TEXT):
            if (self.Discard_cell.value == '' or self.Discard_cell.value == ' '):
                self.Discard_data=None
            else:
                #raise Exception(str(self.cell))
                self.Discard_data=None
    
        elif (self.Discard_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.Discard_data=None
    
        else:
            raise Exception(str(self.Discard_cell))

        return self.Discard_data
    def read_Magazine_LOCATION(self):
        self.LOCATION_cell = self.row[6]
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

    def read_Magazine_NUMBER(self):
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

    def read_Magazine_TITLE(self):
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

    def read_Magazine_YEAR(self):
        self.YEAR_cell = self.row[9]
        if (self.YEAR_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.YEAR_cell.value])
            self.YEAR_data = None

        elif (self.YEAR_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.YEAR_data = int(self.YEAR_cell.value)

        elif (self.YEAR_cell.ctype == xlrd.XL_CELL_TEXT):
            self.YEAR_data = self.YEAR_cell.value.encode('utf-8').strip()

        elif (self.YEAR_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.YEAR_data = None

        else:
            self.YEAR_data = None
            raise Exception(str(self.YEAR_cell))

        return self.YEAR_data

    def read_Magazine_MONTH(self):
        self.MONTH_cell = self.row[10]
        if (self.MONTH_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.MONTH_cell.value])
            self.MONTH_data = None

        elif (self.MONTH_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.MONTH_data = int(self.MONTH_cell.value)

        elif (self.MONTH_cell.ctype == xlrd.XL_CELL_TEXT):
            self.MONTH_data = self.MONTH_cell.value.encode('utf-8').strip()

        elif (self.MONTH_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.MONTH_data = None

        else:
            self.MONTH_data = None
            raise Exception(str(self.MONTH_cell))

        return self.MONTH_data

    def read_Magazine_VOLUME(self):
        self.VOLUME_cell = self.row[11]
        if (self.VOLUME_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.VOLUME_cell.value])
            self.VOLUME_data = None

        elif (self.VOLUME_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.VOLUME_data = int(self.VOLUME_cell.value)

        elif (self.VOLUME_cell.ctype == xlrd.XL_CELL_TEXT):
            self.VOLUME_data = self.VOLUME_cell.value.encode('utf-8').strip()

        elif (self.VOLUME_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.VOLUME_data = None

        else:
            self.VOLUME_data = None
            raise Exception(str(self.VOLUME_cell))

        return self.VOLUME_data

    def read_Magazine_VOLNUM(self):
        self.VOLNUM_cell = self.row[12]
        if (self.VOLNUM_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.VOLNUM_cell.value])
            self.VOLNUM_data = None

        elif (self.VOLNUM_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.VOLNUM_data = int(self.VOLNUM_cell.value)

        elif (self.VOLNUM_cell.ctype == xlrd.XL_CELL_TEXT):
            self.VOLNUM_data = self.VOLNUM_cell.value.encode('utf-8').strip()

        elif (self.VOLNUM_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.VOLNUM_data = None

        else:
            self.VOLNUM_data = None
            raise Exception(str(self.VOLNUM_cell))

        return self.VOLNUM_data

    def read_Magazine_WHOLE(self):
        self.WHOLE_cell = self.row[13]
        if (self.WHOLE_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.WHOLE_cell.value])
            self.WHOLE_data = None

        elif (self.WHOLE_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.WHOLE_data = int(self.WHOLE_cell.value)

        elif (self.WHOLE_cell.ctype == xlrd.XL_CELL_TEXT):
            self.WHOLE_data = self.WHOLE_cell.value.encode('utf-8').strip()

        elif (self.WHOLE_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.WHOLE_data = None

        else:
            self.WHOLE_data = None
            raise Exception(str(self.WHOLE_cell))

        return self.WHOLE_data

    def read_Magazine_COMMENTS1(self):
        self.COMMENTS1_cell = self.row[14]
        if (self.COMMENTS1_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.COMMENTS1_cell.value])
            self.COMMENTS1_data = None

        elif (self.COMMENTS1_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.COMMENTS1_data = int(self.COMMENTS1_cell.value)

        elif (self.COMMENTS1_cell.ctype == xlrd.XL_CELL_TEXT):
            self.COMMENTS1_data = self.COMMENTS1_cell.value.encode('utf-8').strip()

        elif (self.COMMENTS1_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.COMMENTS1_data = None

        else:
            self.COMMENTS1_data = None
            raise Exception(str(self.COMMENTS1_cell))

        return self.COMMENTS1_data

    def read_Magazine_ENTERED(self):
        self.ENTERED_cell = self.row[15]
        if (self.ENTERED_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:"+ xlrd.error_text_from_code[self.cell.value])
            self.ENTERED_data=None
    
        elif (self.ENTERED_cell.ctype == xlrd.XL_CELL_DATE and self.ENTERED_cell.value != ''):
            self.ENTERED_data=datetime.datetime(*xlrd.xldate_as_tuple(self.ENTERED_cell.value, self.xlrdworkbook.datemode))
    
        elif (self.ENTERED_cell.ctype == xlrd.XL_CELL_TEXT):
            if (self.ENTERED_cell.value == '' or self.ENTERED_cell.value == ' '):
                self.ENTERED_data=None
            else:
                #raise Exception(str(self.cell))
                self.ENTERED_data=None
    
        elif (self.ENTERED_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.ENTERED_data=None
    
        else:
            self.ENTERED_data = None
            raise Exception(str(self.ENTERED_cell))

        return self.ENTERED_data

    def read_Magazine_COMMENTS2(self):
        self.COMMENTS2_cell = self.row[17]
        if (self.COMMENTS2_cell.ctype == xlrd.XL_CELL_ERROR):
            #logger.error("XL_CELL_ERROR:" + xlrd.error_text_from_code[self.COMMENTS2_cell.value])
            self.COMMENTS2_data = None

        elif (self.COMMENTS2_cell.ctype == xlrd.XL_CELL_NUMBER):
            self.COMMENTS2_data = int(self.COMMENTS2_cell.value)

        elif (self.COMMENTS2_cell.ctype == xlrd.XL_CELL_TEXT):
            self.COMMENTS2_data = self.COMMENTS2_cell.value.encode('utf-8').strip()

        elif (self.COMMENTS2_cell.ctype == xlrd.XL_CELL_EMPTY):
            self.COMMENTS2_data = None

        else:
            self.COMMENTS2_data = None
            raise Exception(str(self.COMMENTS2_cell))

        return self.COMMENTS2_data


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

