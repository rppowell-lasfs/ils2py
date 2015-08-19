# -*- coding: utf-8 -*-

# try something like
import xlrd

from gluon.custom_import import track_changes; track_changes(True)

import ils2py
import ils2py.db_defaults

import logging

logger = logging.getLogger("web2py.app.ils2py")
logger.setLevel(logging.DEBUG)

#import import_library_Books_xls

def index(): return dict(message="hello from _dev.py")


def db_reset(): 
    logger.debug("db_reset()")
    messages = ["Resetting database..."]
    for table in db.tables:
        messages.append("db.{table}.truncate()".format(table=table))
        db[table].truncate()

    logger.debug("\n".join(messages))
    return dict(message=PRE("\n".join(messages)))


def db_setup_scaffolding(): 
    messages = ["Setting up database..."]
    for ADMIN_GROUP in ils2py.db_defaults.ILS_ADMIN_GROUPS:
        db.auth_group.insert(
            role=ADMIN_GROUP['role'], description=ADMIN_GROUP['description']
        )
        messages.append("db.auth_group.insert(%s)"%(str(ADMIN_GROUP)))

    for n, s in [ 
            ('ils_item_type', ils2py.db_defaults.ILS_ITEM_TYPES), 
            ('ils_item_location', ils2py.db_defaults.ILS_ITEM_LOCATIONS), 
            ('ils_item_state', ils2py.db_defaults.ILS_ITEM_STATES), 
            ('ils_item_event_type', ils2py.db_defaults.ILS_ITEM_EVENT_TYPES), 
            ('ils_person_type', ils2py.db_defaults.ILS_PERSON_TYPES), 
    ]:
        for i in s:
            db[n].insert(
                name=i['name'], description=i['description']
            )
            messages.append("db.%s.insert(%s)"%(n,str(i)))

    return dict(message=PRE("\n".join(messages)))

def db_setup_test_users():
    messages = ["Setting up test users..."]
    
    auth.add_membership(auth.id_group(role='Head Librarian'),
        db.auth_user.insert(
            username='HeadLibrarian', password=db.auth_user.password.validate('password')[0],
            email='headlibrarian@test.com', first_name='HeadLibrarian', last_name='HeadLibrarian'
        )
    )
    messages.append("Created 'HeadLibrarian'")

    for i in range(1,4):
        db.auth_user.insert(
            username='User{:03d}'.format(i), password=db.auth_user.password.validate('password{:03d}'.format(i))[0],
            email='user{:03d}@test.com'.format(i), first_name='First{:03d}'.format(i), last_name='Last{:03d}'.format(i)
        )
        messages.append("Created 'User{:03d}'".format(i))

    return dict(message=PRE("\n".join(messages)))

def db_setup_test_items():
    messages = ["Setting up test items..."]

    messages.append("Setting up persons...")
    for i in range(1,4):
        id = db.ils_person.insert(
            full_name='PersonFull{:03d}'.format(i), 
            search_name='PersonSearch{:03d}'.format(i), 
            first_name='PersonFirst{:03d}'.format(i), 
            last_name='PersonLast{:03d}'.format(i), 
        )
        messages.append("Created 'Person{:03d} ({})'".format(i, id))

    messages.append("Setting up publishers...")
    for i in range(1,4):
        id = db.ils_publisher.insert(
            name='Publisher{:03d}'.format(i)
        )
        messages.append("Created 'Publisher{:03d} ({})'".format(i, id))

    messages.append("Setting up book items...")

    return dict(message=PRE("\n".join(messages)))

