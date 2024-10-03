DB_FILES = [

]

CHATS_IDS = ''

try:
    from _constants import *
except ModuleNotFoundError:
    pass

if not CHATS_IDS:
    raise NotImplementedError('Declaration of the constants CHATS_IDS is required')


