# -*- coding: utf-8 -*-
 
# use generic views
response.generic_patterns = ['*']

# reload modules automatically when changes are detected
from gluon.custom_import import track_changes
track_changes(True)


