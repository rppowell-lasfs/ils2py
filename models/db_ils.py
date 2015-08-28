# -*- coding: utf-8 -*-

import re

def natural_key(string_):
    return [int(s) if s.isdigit() else s for s in re.split(r'(\d+)', string_)]


################################################################################

db.define_table(
    'ils_publisher',
    Field('name'),
    format = '%(name)s'
)

db.define_table(
    'ils_person',
    Field('full_name'),
    Field('search_name'),
    Field('first_name'),
    Field('last_name'),
    format = '%(full_name)s'
)


db.define_table(
    'ils_item_x_person_type',
    Field('name'),
    Field('description'),
    format = '%(name)s'
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
    'ils_item',
    Field('item_id', 'string'),
    Field('item_title', 'string'),
    Field('item_type', db.ils_item_type),
    Field('item_location', db.ils_item_location),
    Field('item_state', db.ils_item_state),
    Field('item_publisher', db.ils_publisher),
    Field('item_author', db.ils_person),
    Field('item_isbn10', 'string'),
    Field('item_isbn13', 'string'),
    format = '%(item_id)s:%(item_title)s'
)

db.ils_item.item_id.requires=IS_NOT_IN_DB(db,'ils_item.item_id')
db.ils_item.item_type.requires=IS_IN_DB(db,'ils_item_type.id', '%(name)s')
db.ils_item.item_location.requires=IS_IN_DB(db,'ils_item_location.id', '%(name)s')
db.ils_item.item_state.requires=IS_EMPTY_OR(IS_IN_DB(db,'ils_item_state.id', '%(name)s'))

#db.ils_item.item_type.widget = lambda f, v: SELECT([OPTION(i.name, _value=i.id) for i in sorted_ils_item_types], _name=f.name, _id="%s_%s" % (f._tablename, f.name), _value=v, value=v)
db.ils_item.item_type.widget = lambda f, v: SELECT(['']+[OPTION(i.name, _value=i.id) for i in sorted_ils_item_types], _name=f.name, _id="%s_%s" % (f._tablename, f.name), _value=v, value=v)

##db.ils_item.item_location.widget = lambda f, v: SELECT([OPTION(i.name, _value=i.id) for i in sorted_ils_item_locations], _name=f.name, _id="%s_%s" % (f._tablename, f.name), _value=v, value=v)
db.ils_item.item_location.widget = lambda f, v: SELECT(['']+[OPTION(i.name, _value=i.id) for i in sorted_ils_item_locations], _name=f.name, _id="%s_%s" % (f._tablename, f.name), _value=v, value=v)


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
    'ils_item_loan',
    Field('ils_item', db.ils_item),
    Field('checked_out_on', 'datetime'),
    Field('checked_out_by', db.auth_user),
    Field('checked_out_to', db.auth_user),
    Field('checked_in_on', 'datetime'),
    Field('checked_in_by', db.auth_user),
)

db.define_table(
    'ils_item_tag',
    Field('name', 'string'),
    Field('description', 'string'),
    Field('parent', 'reference ils_item_tag')
)

db.define_table(
    'ils_item_x_item_tag',
    Field('ils_item', db.ils_item),
    Field('ils_item_tag', db.ils_item)
)

################################################################################


db.define_table(
    'ils_item_x_person',
    Field('item', db.ils_item),
    Field('person', db.ils_person),
    Field('relationship', db.ils_item_x_person_type),
)
