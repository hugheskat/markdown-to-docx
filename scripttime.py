def runningtime(starttime, endtime):

    # imports
    import datetime
    import re
    # get the number of seconds as an hhmmss string
    hhmmss = str(datetime.timedelta(seconds=(endtime - starttime)))
    # set default string values
    strhr = ' hrs'
    strmin = ' mins'
    strsec = ' secs'
    # get time data
    hrs = str(int(re.search(r'(\d+):(\d+):(\d+)', hhmmss).group(1)))
    if hrs == '1':
        strhr = ' hr'
    mins = str(int(re.search(r'(\d+):(\d+):(\d+)', hhmmss).group(2)))
    if mins == '1':
        strmin = ' min'
    secs = str(int(re.search(r'(\d+):(\d+):(\d+)', hhmmss).group(3)))
    if secs == '1':
        strsec = ' sec'
    # contruct the user-friendly string
    if hrs == '0':
        if mins == '0' and secs == '0':
            nicetime = 'less than a second'
        elif mins == '0' and secs != '0':
            nicetime = secs + strsec
        else:
            if secs == 0:
                nicetime = mins + strmin
            else:
                nicetime = mins + strmin + ' ' + secs + strsec
    else:
        if mins == '0':
            if secs == '0':
                nicetime = hrs + strhr
            else:
                nicetime = hrs + strhr + ' ' + mins + strmin + ' ' + secs + strsec
        else:
            if secs == '0':
                nicetime = hrs + strhr + ' ' + mins + strmin
            else:
                nicetime = hrs + strhr + ' ' + mins + strmin + ' ' + secs + strsec
    # and send it back
    return nicetime
