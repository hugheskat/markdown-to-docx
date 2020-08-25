def getmessagetext(filecount):

    msgtxt = ''
    if filecount == 0:
        msgtxt = '0 files processed in '
    elif filecount == 1:
        msgtxt = '1 file processed in '
    else:
        msgtxt = str(filecount) + ' files processed in '

    # and send it back
    return msgtxt
