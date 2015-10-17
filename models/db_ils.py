# -*- coding: utf-8 -*-

import re

def natural_key(string_):
    return [int(s) if s.isdigit() else s for s in re.split(r'(\d+)', string_)]


################################################################################

db.define_table(
    'ils_biblio_type',
    Field('name', 'string'),
    Field('description', 'string'),
    format = '%(name)s'
)
sorted_ils_biblio_types = sorted(db(db.ils_biblio_type.id > 0).select(), key=lambda x: natural_key(x.name))
    
db.define_table(
    'ils_biblio',
    Field('biblio_title', 'string'),
    Field('biblio_type', 'string'),
    Field('biblio_isbn', 'string'),
)

db.ils_biblio.biblio_type.widget = lambda f, v: SELECT(['']+[OPTION(i.name, _value=i.id) for i in sorted_ils_biblio_types], _name=f.name, _id="%s_%s" % (f._tablename, f.name), _value=v, value=v)

db.define_table(
    'ils_biblio_publisher',
    Field('name', 'string'),
    format = '%(name)s'
)

db.define_table(
    'ils_biblio_person',
    Field('full_name', 'string'),
    Field('search_name', 'string'),
    Field('first_name', 'string'),
    Field('last_name', 'string'),
    format = '%(full_name)s'
)

db.define_table(
    'ils_biblio_person_type',
    Field('name', 'string'),
    Field('description', 'string'),
    format = '%(name)s'
)

db.define_table(
    'ils_biblio_x_person',
    Field('ils_biblio', db.ils_biblio),
    Field('ils_biblio_person', db.ils_biblio_person),
    Field('ils_biblio_person_type', db.ils_biblio_person_type),
)

#db.ils_biblio_x_person.ils_biblio.requires=IS_IN_DB(db,'ils_biblio', '%(title)s'))
#db.ils_biblio_x_person.ils_biblio_person_type.requires=IS_EMPTY_OR(IS_IN_DB(db,'ils_bibio_person_type', '%(name)s'))

db.define_table(
    'ils_biblio_tag',
    Field('name', 'string'),
    Field('description', 'string'),
    Field('parent', 'reference ils_biblio_tag')
)

db.define_table(
    'ils_biblio_x_tag',
    Field('ils_biblio', db.ils_biblio),
    Field('ils_biblio_tag', db.ils_biblio_tag)
)

################################################################################

db.define_table(
    'ils_item_type',
    Field('name', 'string'),
    Field('description', 'string'),
    format = '%(name)s'
)

sorted_ils_item_types = sorted(db(db.ils_item_type.id > 0).select(), key=lambda x: natural_key(x.name))

db.define_table(
    'ils_item_location',
    Field('name', 'string'),
    Field('description', 'string'),
    format = '%(name)s'
)

sorted_ils_item_locations = sorted(db(db.ils_item_location.id > 0).select(), key=lambda x: natural_key(x.name))


db.define_table(
    'ils_item_state',
    Field('name', 'string'),
    Field('description', 'string'),
    format = '%(name)s'
)

db.define_table(
    'ils_item_publisher',
    Field('name', 'string'),
    Field('description', 'string'),
    format = '%(name)s'
)

db.define_table(
    'ils_item_person',
    Field('full_name', 'string'),
    format = '%(full_name)s'
)

db.define_table(
    'ils_item',
    Field('item_id', 'string'),
    Field('item_title', 'string'),
    Field('item_type', db.ils_item_type),
    Field('item_location', db.ils_item_location),
    Field('item_state', db.ils_item_state),
    Field('item_publisher', db.ils_item_publisher),
    Field('item_author', db.ils_item_person),
    Field('item_coauthor', db.ils_item_person),
    Field('item_series', 'string'),
    Field('item_isbn', 'string'),
    Field('item_msrp', 'string'),
    Field('item_biblio', db.ils_biblio),
    format = '%(item_id)s:%(item_title)s'
)

db.ils_item.item_id.requires=IS_NOT_IN_DB(db,'ils_item.item_id')
db.ils_item.item_type.requires=IS_IN_DB(db,'ils_item_type.id', '%(name)s')
db.ils_item.item_location.requires=IS_IN_DB(db,'ils_item_location.id', '%(name)s')
db.ils_item.item_state.requires=IS_EMPTY_OR(IS_IN_DB(db,'ils_item_state.id', '%(name)s'))
db.ils_item.item_biblio.requires=IS_EMPTY_OR(IS_IN_DB(db,'ils_biblio.id', '%(biblio_title)s'))

#db.ils_item.item_type.widget = lambda f, v: SELECT([OPTION(i.name, _value=i.id) for i in sorted_ils_item_types], _name=f.name, _id="%s_%s" % (f._tablename, f.name), _value=v, value=v)
db.ils_item.item_type.widget = lambda f, v: SELECT(['']+[OPTION(i.name, _value=i.id) for i in sorted_ils_item_types], _name=f.name, _id="%s_%s" % (f._tablename, f.name), _value=v, value=v)

##db.ils_item.item_location.widget = lambda f, v: SELECT([OPTION(i.name, _value=i.id) for i in sorted_ils_item_locations], _name=f.name, _id="%s_%s" % (f._tablename, f.name), _value=v, value=v)
db.ils_item.item_location.widget = lambda f, v: SELECT(['']+[OPTION(i.name, _value=i.id) for i in sorted_ils_item_locations], _name=f.name, _id="%s_%s" % (f._tablename, f.name), _value=v, value=v)

################################################################################

db.define_table(
    'ils_item_tag',
    Field('name', 'string'),
    Field('description', 'string'),
    Field('parent', 'reference ils_item_tag')
)

db.define_table(
    'ils_item_x_tag',
    Field('ils_item', db.ils_item),
    Field('ils_item_tag', db.ils_item_tag)
)

################################################################################

db.define_table(
    'ils_item_event_type',
    Field('name', 'string'),
    Field('description', 'string'),
    format = '%(name)s'
)

db.define_table(
    'ils_item_event',
    Field('ils_item_event_type', db.ils_item_event_type),
)

db.define_table(
    'ils_item_circulation',
    Field('ils_item', db.ils_item),
    Field('checked_out_on', 'datetime'),
    Field('checked_out_by', db.auth_user),
    Field('checked_out_to', db.auth_user),
    Field('checked_in_on', 'datetime'),
    Field('checked_in_by', db.auth_user),
)

################################################################################

db.define_table(
    'ils_cart',
    Field('person', db.auth_user),
    auth.signature
)

db.define_table(
    'ils_cart_x_item',
    Field('ils_cart', db.ils_cart),
    Field('ils_item', db.ils_item)
)

