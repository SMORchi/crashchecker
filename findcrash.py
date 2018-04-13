def findcrash (flash):

    import re


    match = re.search('crashinfo', flash)


    return match
