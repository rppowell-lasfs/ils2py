# -*- coding: utf-8 -*-
import xlrd
import ils2py
import ils2py.import_library_books_xls

def import_library_books_from_xls():
    messages = ["Importing XLS"]
    form = FORM(
        INPUT(_name='xls',_type='file'),
        INPUT(_type='submit')
    )
    data = {}
    if form.process().accepted:
        response.flash = 'form accepted'
        #tmp_str = request.vars.xls.file.read()
        #book = xlrd.open_workbook(request.vars.xls.file.read())
        # encoding_override="cp1252"
        book = xlrd.open_workbook(file_contents=request.vars.xls.file.read(), encoding_override="cp1252")
        messages.append("book={}".format(str(book)))

        #request._vars['xls'] = 'Undefined'
        #print str(request._vars['xls'])
        #print str(response._vars)

        ils2py.import_library_books_xls.check_headers(book)
        data = ils2py.import_library_books_xls.import_library_books_xls(book)

        # import bibliographies to database
        if (False):
            if (True):
                for t in data['types']:
                    if t == '':
                        pass
                    if not db(db.ils_item_type.name==t).select():
                        db.ils_item_type.insert(name=t)
                        messages.append("inserting type '%s'"%(t))
            if (True):
                for location in data['locations']:
                    if location == '':
                        pass
                    if not db(db.ils_item_location.name==location).select():
                        db.ils_item_location.insert(name=location)
                        messages.append("inserting location '%s'"%(location))
            if (True):
                for publisher in sorted(data['publishers'].keys()):
                    if publisher == '':
                        pass
                    if not db(db.ils_publisher.name==publisher).select():
                        db.ils_publisher.insert(name=publisher)
                        messages.append("inserting publisher '%s'"%(publisher))
    
            if (True):
                for author in data['authors']:
                    if author == '':
                        pass
                    if not db(db.ils_person.full_name==author).select():
                        db.ils_person.insert(full_name=author)
                        messages.append("inserting bibliography_person '%s'"%(author))

            # import items to database
            if (True):
                messages.append("processing %d entries"%(len(data['entries'])))
                for i, entry in enumerate(data['entries']):
                    if db(db.ils_item.item_id==entry['number']).select():
                        messages.append("WARNING: entry %s(%d) exists"%(str(entry['title']), entry['number']))
                    else:
                        print "inserting", i, entry
                        e = db.ils_item.insert(
                            item_id = entry['number'],
                            item_title = entry['title'],
                            item_type = db(db.ils_item_type.name==entry['type']).select()[0],
                            item_location = db(db.ils_item_location.name==entry['location']).select()[0],
                            item_publisher = db(db.ils_publisher.name==entry['publisher']).select()[0],
                            item_author = db(db.ils_person.full_name==entry['author']).select()[0]
                        )
                        messages.append("inserting item: %s"%([entry['title'],entry['number']]))

        # circulations
        if (False):
            if (True):
                for member in sorted(data['members'].keys()):
                    if member == '': 
                        pass
                    if ',' in member:
                        member_last, member_first = member.split(',', 2)
                        if not db(db.auth_user.first_name == member_first and db.auth_user.last_name== member_last).select():
                            db.auth_user.insert(first_name=member_first, last_name=member_last)
                            messages.append("auth_user inserting %s, %s"%(member_last, member_first))
                    else:
                        if not db(db.auth_user.username == member).select():
                            db.auth_user.insert(username=member)
                            messages.append("auth_user inserting %s"%(member))
    
            if (True):
                for librarian in sorted(data['librarians'].keys()):
                    if librarian == '':
                        pass
                    if not db(db.auth_user.username == librarian).select():
                        db.auth_user.insert(username=librarian)
                        messages.append("auth_user inserting %s"%(librarian))

        # disable toolbar - binary file upload gets in the way

    elif form.errors:
        response.flash = 'form has errors'

    return dict(form=form, data=data, message=PRE("\n".join(messages)))


