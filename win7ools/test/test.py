# -*- coding: utf-8 -*-
"""
Created on Tue Jul 08 19:00:09 2014

@author: AEOLIA
"""



#fname = r"C:\Python27\Lib\site-packages\wintools\test\test.txt"
#attr = ctypes.windll.kernel32.GetFileAttributesW(unicode(fname))
#print(attr)
#print(bool(attr & 2))
#print(bool(attr & 16))
#print(bool(attr & 32))


#1 - read only file
#2 - hidden file/directory
#4 - OS file/directory
#16 - directory
#32 - archive file/directory
#64 - reserved
#128 - file that does not have other attributes set.  Used alone
#256 - used for temporary storage
#512 - sparse file
#1024 - reparse point or symbolic link
#2048 - compressed file/directory
#4096 - data of file not available
#8192 - not indexed
#16384 - encrypted file/directory
#32768 - configured with integrity
#131072 - cannot be read by scrubber

#windows CMD commands to look into:
#sfc/scannow

#USE:
#cipher /w:DRIVE_LETTER:\DIRECTORY (permanently delete deleted data)

#hostmonster shared server: 74.220.215.245

#ping --> parse for connection speed
#pathping --> useful, but has 225 second wait
#tracert -->
#   hop#    time(ms)    time(ms)    time(ms)    Hostname [IP]
#                                                (or just IP)
#assoc --> .EXT=Program.Name.Thing
#Tasklist --> name, PID, session, session #, memory usage
#Tasklist /v --> verbose information including window name and PID.
#Tasklist /v /fi "STATUS eq running" --> conditional filtering
#netstat -an --> protocol, local IP, foreign IP, state(TCP only)