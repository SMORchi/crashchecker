def logincisco():

    import os
    import re

    LOGIN = os.getlogin()
    MATCH = re.search('(\w+)-(\w+)', LOGIN)
    USERNAME = MATCH.group(2) + '-' + MATCH.group(1)

    return USERNAME

