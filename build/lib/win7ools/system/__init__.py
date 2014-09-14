# -*- coding: utf-8 -*-
'''System -- Classes to work with Windows systems'''

import _winreg as reg
import cPickle
import ctypes
import multiprocessing as multi
import os
import os.path as path
import platform
import re
import string
import subprocess
import time
from collections import defaultdict
from distutils.log import warn
from PIL import Image
from subprocess import call
from wintools.lib import img_loads
from wintools.lib import log
from wintools.lib import truncate
from wintools.ipl import IPL
from wintools.reg import get_values as values
from wintools.reg import RegistryKeys
from wintools.reg import LOCAL_MACHINE, CLASSES_ROOT, CURRENT_USER, ALL_ACCESS
from wintools.sec import md5

class Metadata(object):
    '''
    Class to set and get metadata from objects.
    EXIV from images, ADS from files, etc...
    '''
    def __init__(self, pathname=''):
        self.description = 'Hello World'

class Monitor(object):
    '''Object for retrieving attributes of a monitor.'''
    def __init__(self):
        screen_width = ctypes.windll.user32.GetSystemMetrics(0)
        screen_height = ctypes.windll.user32.GetSystemMetrics(1)
        #number of monitors
        #monitor_number = ctypes.windll.user32.GetSystemMetrics(80)
        self.width = screen_width
        self.height = screen_height
        aspect_ratio = truncate(float(screen_width) / screen_height, 2)
        self.aspect_ratio = aspect_ratio
        
    def logon_dimensions(self):
        dimensions = {'0.60': (768, 1280),\
                      '0.63': (900, 1440),\
                      '0.75': (960, 1280),\
                      '0.80': (1024, 1280),\
                      '1.25': (1280, 1024),\
                      '1.33': (1024, 768),\
                      '1.60': (1440, 900),\
                      '1.66': (1280, 768),\
                      '1.77': (1360, 768)}  
                          
        return dimensions[str(self.aspect_ratio)]
        
class Pictures(object):
    '''
    Stores serialized copies of logon screen, user account icons, and 
    ObjectDock PNG files. Also stores paths to useful Windows picture folders
    and drive icons.
    '''   
    def __init__(self, system_drive='C:'):
        self.home = path.expanduser("~")
        self.profile = path.join(system_drive,\
            r'\ProgramData\Microsoft\User Account Pictures\Default Pictures')       
        self.wallpapers_path = path.join(system_drive,\
            '\\Windows\\Web\\Wallpaper', os.environ['computername'])
        if not path.exists(self.wallpapers_path):
            os.makedirs(self.wallpapers_path)
            os.startfile(self.wallpapers_path)
        self.logon_screen = system_drive +\
                            r'\Windows\Sysnative\oobe\Info\backgrounds'
        self.logon_screen_path = self.logon_screen + r'\backgroundDefault.jpg'
        self.icon_cache = self.home + '\\AppData\\Local\\IconCache.db'
        self.get_drive_icon_data()
        self.user = []

    def clear_icon_cache(self):
        try:
            os.remove(self.icon_cache)
            return True
        except:
            return False         
        
    def get_drive_icon_data(self, keys=RegistryKeys()):
        '''
        Load drive icon data into pictures object
        '''
        self.drives = []
        drives = keys.drives
        temp = values(drives)
        value = 'RegistryName'
        labels = [x[value] for x in temp if len(x[value]) == 1]
        for label in labels:
            drive_path = values('%s\%s'%(drives, label))[0]['Values']['']
            drive_name = path.basename(drive_path)
            self.drives.append({'label': label,
                                'name': drive_name,
                                'path': drive_path}) 
                
class System(object):
    '''
    Object to manage computer settings and enable sync/backup/restore (SBR)
    
    Software Information:
        chrome bookmarks: 
            appdata + '\Local\Google\Chrome\User Data\Default\Bookmarks'
        objectdock:
            appdata + '\Local\Stardock\ObjectDockPlus'
        launchy:
            can be found in program_files
    '''   
    def __init__(self, debug=False):
        self.name = os.environ['computername']
        self.os = platform.system() + ' ' + platform.release()
        self.system_drive = os.environ['systemdrive']
        self.winevt = os.environ['windir'] + '\\Sysnative\\winevt\\Logs'
        self.home = path.expanduser("~")
        self.appdata = self.home + '\\AppData'
        #Recently opened files from Windows Explorer
        self.recent = self.appdata + '\\\Roaming\\Microsoft\\Windows\\Recent'
        self.program_files = os.environ['programfiles']
        self.monitor = Monitor()
        if not debug: self.pictures = Pictures(system_drive=self.system_drive) 
        else: self.pictures = []
        self.cpu_count = multi.cpu_count()
        self.ipl = IPL()
        self.keys = RegistryKeys()
     
    def __str__(self):
        return self.__repr__()
    
    def __repr__(self):
        info = '%s (%s)'%(self.name, self.ipl.date)
        info += '; OS: %s'%self.os
        info += '; CPU Count: %s'%self.cpu_count
        info += '; Programs: %s'%len(self.ipl['Programs'])
        info += '; Modules: %s'%len(self.ipl['Modules'])
        return info       

    def enable_custom_logon_screen(self):
        '''    
        Set OEMBackground and useOEMBackground registry keys
        to allow custom windows logon screen.
        Returns values of reg keys as [OEMBackground, UseOEMBackground]
        ***This does not copy an image to the appropriate folder***
        
        Sysnative is used instead of System32 since 32-bit apps are re-directed
        on 64-bit machines and do not have direct access to the 64-bit reserved
        System32 folder.  32-bit apps are redirected to SysWOW64.
        '''
        OEM =''
        useOEM = ''
        with reg.ConnectRegistry(None, LOCAL_MACHINE) as y:
            #Set OEMBackground reg key
            try:
                with reg.OpenKey(y, self.keys.logon, 0, ALL_ACCESS) as logon:
                    reg.SetValueEx(logon, "OEMBackground", 0, reg.REG_DWORD, 1)
                    OEM = reg.QueryValueEx(logon, "OEMBackground")[0]
            except:
                warn('Failed to set OEMBackground registry key')
            #Set UseOEMBackground reg key
            try:
                keyVal = 'SOFTWARE\\Policies\\Microsoft\\Windows\\System'
                with reg.OpenKey(y, keyVal, 0, ALL_ACCESS) as use_oem:
                    reg.SetValueEx(use_oem,\
                                  'UseOEMBackground',\
                                   0,\
                                   reg.REG_DWORD,\
                                   1)
                    useOEM = reg.QueryValueEx(use_oem, "UseOEMBackground")[0]
            except:
                warn('Failed to set UseOEMBackground registry key')
        return not not (OEM and useOEM) 
       
    def set_logon_screen(self, imageData):
        '''
        ***Must be run with administrator privileges***
        '''
        success = False
        screen = Monitor()
        savePath = self.pictures.logon_screen
        if not path.exists(savePath):
            os.makedirs(savePath)
        if path.exists(imageData):
            img = Image.open(imageData)
        try:
            if Image.isImageType(imageData):
                img = imageData
        except:
            pass
        try:
            if Image.isImageType(img_loads(imageData)):
                img = img_loads(imageData)
        except:
            pass
        try:
            logon_img = img.resize(screen.logon_dimensions(), Image.ANTIALIAS)
        except:
            warn('Failed to resize image.')
        try:
            if self.enable_custom_logon_screen():
                logon_img.save(savePath + '\\backgroundDefault.jpg')
                success = True
            else:
                warn('Failed to enable custom logon screen')
        except:
            warn('Failed to customize logon screen')
        return success
        
    def get_trim(self):
        '''Return [Prefetch, Superfetch].'''
        trimPrefetch = ''
        trimSuperfetch = ''
        try:
            trimKeys = values(self.keys.TRIM)
            trimPrefetch = trimKeys['Values']['EnablePrefetcher']
            trimSuperfetch = trimKeys['Values']['EnableSuperfetch']
        except:
            warn('Failed to retrieve TRIM registry keys.')
        return [trimPrefetch, trimSuperfetch]
    
    def set_trim(self, enable=True):
        '''
        Enable TRIM ==> [Prefetch, Superfetch] = [0, 0]

        >>> computer = System()
        >>> computer.set_trim()
        True
        >>> computer.get_trim()
        [0, 0]
        >>> computer.set_trim(enable=False)
        True
        >>> computer.get_trim()
        [1, 1]
        >>> computer.set_trim()
        True
        '''
        Prefetch = ''
        Superfetch = ''
        val = 0 if enable else 1
        keys = self.keys
        try:
            with reg.OpenKey(LOCAL_MACHINE, keys.TRIM, 0, ALL_ACCESS) as trim:
                reg.SetValueEx(trim, 'EnablePrefetcher', 0, reg.REG_DWORD, val)
                reg.SetValueEx(trim, 'EnableSuperfetch', 0, reg.REG_DWORD, val) 
                Prefetch = reg.QueryValueEx(trim, 'EnablePrefetcher')[0]
                Superfetch = reg.QueryValueEx(trim, 'EnableSuperfetch')[0]
        except:
            warn('Failed to set TRIM registry keys.')
            warn('(Prefetch, Superfetch) = (%s, %s)'%(Prefetch, Superfetch))
        return Prefetch == Superfetch == val   
        
    def add_extension(self, *args):
        '''Add file extension(s) to %PATHEXT% system variable'''
        for ext in args:
            pathext = ''
            x = ''
            if isinstance(ext, list):
                for i in ext:
                    System.add_extension(i)
            if isinstance(ext, str):
                try:
                    x = reg.ConnectRegistry(None, LOCAL_MACHINE)
                    pathKey = reg.OpenKey(x,\
                                          self.keys.environment,\
                                          0,\
                                          ALL_ACCESS)
                    pathext = reg.QueryValueEx(pathKey, "pathext")[0]
                except:
                    warn('Failed to open pathext reg key.')
                if '.' in ext:
                    fileExt = ';' + ext.upper()
                else:
                    fileExt = ';.' + ext.upper()
                if fileExt not in pathext:
                    reg.SetValueEx(pathKey,\
                                  'pathext',\
                                   0,\
                                   reg.REG_EXPAND_SZ,\
                                   pathext + fileExt)
                    pathext = reg.QueryValueEx(pathKey, "pathext")[0]    
                    if fileExt in pathext:
                        pass
                    else:
                        warn('FAILED to add %s to \%PATHEXT\%.'%fileExt)
                else:
                    pass
                if x:
                    reg.CloseKey(pathKey)
                    reg.CloseKey(x)   

    def add_path(self, *args):
        '''
        Add path(s) to %PATH% system variable.
        Uses os module for current session.
        Permanent key change requires reboot.
        '''
        val = self.keys.environment
        for path in args:
            if isinstance(path, list):
                for i in path:
                    System.add_path(i)
            if isinstance(path, str):
                newEnvPath = ''
                if os.path.exists(path):
                    with reg.ConnectRegistry(None, LOCAL_MACHINE) as x:
                        with reg.OpenKey(x, val, 0, ALL_ACCESS) as pathKey:
                            Env = reg.QueryValueEx(pathKey,"path")[0]
                            if path not in Env and path.upper() not in Env:
                                newEnvPath = re.sub(';;',\
                                                    ';',\
                                                    Env + ';' + path)
                                if newEnvPath:
                                    os.environ['path'] = newEnvPath
                                reg.SetValueEx(pathKey,\
                                              'path',\
                                               0,\
                                               reg.REG_EXPAND_SZ,\
                                               newEnvPath)
                                Env = reg.QueryValueEx(pathKey,"path")[0]
                                if path in Env:
                                    pass
                                else:
                                    warn('Failed to add %s to %%PATH%%'%path)
                            else:
                                pass
                            
    def set_py_runas_admin(self):
        value = '\"C:\\Python27\\python.exe\" \"%1\" %*'      
        try:
            with reg.CreateKey(CLASSES_ROOT, self.keys.py_runas_admin) as key:
                reg.SetValue(key, 'command', reg.REG_SZ, value)
        except:
            warn('Failed to set key for running python file type as admin.')
        py_runas_admin = self.keys.py_runas_admin
        return values(py_runas_admin, HKCR=True)[0]['Values'][''] == value
        
    def set_drive_icon(self, drive, iconPath):
        '''
        Sets the registry keys for CURRENT_USER and LOCAL_MACHINE.
        '''
        #Set registry key for LOCAL_MACHINE
        DRIVE = string.capitalize(drive)
        keyStr = '%s\\%s\\DefaultIcon'%(self.keys.drives, DRIVE)
        try:
            with reg.ConnectRegistry(None, LOCAL_MACHINE) as x:
                with reg.CreateKey(x, keyStr) as driveIconKey:
                    reg.SetValueEx(driveIconKey, '', 0, reg.REG_SZ, iconPath)
                    reg.CloseKey(driveIconKey)
                    driveIconKey = reg.CreateKey(x, keyStr)
                    if reg.QueryValueEx(driveIconKey, '')[0] == iconPath:
                        pass
                    else:
                        log('FAILED to change icon for %s drive. (HKLM)'%DRIVE)
        except:
            log('Failed to open registry keys.(HKLM)')
        #Set registry key for CURRENT_USER
        try:
            with reg.ConnectRegistry(None, CURRENT_USER) as x:
                with reg.CreateKey(x, keyStr) as driveIconKey:
                    reg.SetValueEx(driveIconKey,"",0,reg.REG_SZ,iconPath)
                    reg.CloseKey(driveIconKey)
                    driveIconKey = reg.CreateKey(x, keyStr)
                    if reg.QueryValueEx(driveIconKey,"")[0] == iconPath:
                        pass
                    else:
                        log('FAILED to change icon for %s drive. (HKCU)'%DRIVE)
        except:
            log('Failed to open registry keys.(HKCU)') 
     
    def save(self, path=''):
        '''
        Class method to save IPL object to sqlite3 database 
        in System package directory.
        
        Only most recent System object is retained.
        
        ex:
            >> comp = System()
            >> comp.save()
        '''
        if path:
            save_path = path
        else:
            save_path = os.path.dirname(__file__)
        save_path = '%s\\%s.%s'%(save_path, self.name, self.ipl.date)
        if os.path.exists(save_path):
            os.remove(save_path)
        f = open(save_path, 'w')
        cPickle.dump(self, f)
        f.close()
        return save_path

    @classmethod
    def load(cls, path):
        '''
        Class method to load System object from sqlite3 database.
        
        ex: 
            >> loaded_System = System.load('path\\to\\System\\db')
        '''
        f = open(path, 'r')
        system = cPickle.load(f)
        f.close()
        return system  
  
    @classmethod
    def lock(cls, wait=0):
        '''
        Locks computer after a certain time, wait, measured in seconds.
        '''
        time.sleep(wait)
        locked = ctypes.windll.user32.LockWorkStation()
        if not locked:
            warn('Failed to lock system.')
            
    @classmethod
    def volume(cls, letter):
        '''
        Returns dictionay with volume name and file system:
        {'name': <volume name>, 'file system': <NTFS, FAT, etc...>}
        
        >>> System.volume('c')['file system']
        'NTFS'
        '''
        kernel32 = ctypes.windll.kernel32
        volumeName = ctypes.create_unicode_buffer(1024)
        serial_number = None
        max_component_length = None
        file_system_flags = None
        fileSystem = ctypes.create_unicode_buffer(1024)
        kernel32.GetVolumeInformationW(ctypes.c_wchar_p("%s:\\"%letter),
                                       volumeName,
                                       ctypes.sizeof(volumeName),
                                       serial_number,
                                       max_component_length,
                                       file_system_flags,
                                       fileSystem,
                                       ctypes.sizeof(fileSystem))
        return {'name': str(volumeName.value),\
                'file system': str(fileSystem.value)}
     
    @classmethod
    def drive_letter(cls, name):
        '''
        Returns letter of volume name.
        '''
        name = name.upper()
        drive_letter = ''
        letters = ['A',\
                   'B',\
                   'C',\
                   'D',\
                   'E',\
                   'F',\
                   'G',\
                   'H',\
                   'I',\
                   'J',\
                   'K',\
                   'L',\
                   'M',\
                   'N',\
                   'O',\
                   'P',\
                   'Q',\
                   'R',\
                   'S',\
                   'T',\
                   'U',\
                   'V',\
                   'W',\
                   'X',\
                   'Y',\
                   'Z']
        for letter in letters:
            if System.volume(letter)['name'] == name:
                drive_letter = letter
                break
        return drive_letter

def get_cmd(cmd):
    proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)   
    for line in proc.stdout:
        print string.split(line.lower())

def get_process_info(process):
    '''Returns list of process info ==> [name, PID, path]'''
    program = process.lower()
    proc_name = ''
    pid = ''
    proc_path = ''
    cmd = 'WMIC PROCESS get Caption, CommandLine, ProcessId'
    proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    for line in proc.stdout:
        tempList = string.split(line.lower())
        if len(tempList) > 1:
            name = tempList[0]
            if program in name:
                proc_name = name
                pid = tempList[len(tempList) - 1]
                m = re.search('C:\\\\.*exe',line)
                if m:
                    proc_path = m.group(0)
                    proc_path = '\"%s\"'%proc_path
                break
    return [proc_name, pid, proc_path]
    
def running(process):
    '''Returns True if running, False otherwise.'''
    pid = get_process_info(process)[1]
    return not not pid

def kill(process):
    '''
    call('taskkill /F /PID pid /T', shell = True)

    /F specifies to forcefully terminate process
    /T terminates process and child processes
    One can terminate processes with /IM <name> instead of /PID <pid> 
    '''
    pid = get_process_info(process)[1]
    if pid:
        cmd = 'taskkill /F /PID ' + pid
        proc = subprocess.Popen(cmd, shell = True, stdout = subprocess.PIPE)
        proc.wait() 
    return not running(process)

def messagebox(text,\
               title='',\
               buttons=0,\
               error=False,\
               question=False,\
               warn=False,\
               information=False,\
               top=True):
    '''
    Default functionality is an "OK" messagebox that opens on top.    
    
    buttons:
        OK = 0x00000000 (default)
        OK_CANCEL = 0x00000001 (1)
        ABORT_RETRY_IGNORE = 0x00000002 (2)
        YES_NO_CANCEL = 0x00000003 (3)
        YES_NO = 0x00000004 (4)
        RETRY_CANCEL = 0x00000005 (5)
        CANCEL_TRY_CONTINUE = 0x00000006 (6) 
    icons:
        ERROR (STOP) = 0x00000010 (16)
        QUESTION = 0x00000020 (32)
        EXCLAMATION (WARNING) = 0x00000030 (48)
        INFORMATION = 0x00000040 (64)
    modality:
        SYSTEM_MODAL (TOPMOST, INTERACTION ALLOWED) = 0x00001000 (4096)
    '''
    return_codes = [0,\
                    'OK',\
                    'CANCEL',\
                    'ABORT',\
                    'RETRY',\
                    'IGNORE',\
                    'YES',\
                    'NO',\
                    0,\
                    0,\
                    'TRY AGAIN',\
                    'CONTINUE']
    proto = ctypes.WINFUNCTYPE(ctypes.c_int,\
                               ctypes.c_void_p,\
                               ctypes.c_char_p,\
                               ctypes.c_char_p,\
                               ctypes.c_ulong)                          
    flags = '%s | %s | %s | %s | %s | %s'%(buttons,\
                                            16 if error else 0,\
                                            32 if question else 0,\
                                            48 if warn else 0,\
                                            64 if information else 0,\
                                            4096 if top else 0)
    handle = (1, "handle", 0)
    text = (1, "text", text)
    title = (1, "title", title)
    flags = (1, "flags", eval(flags))
    params = handle, text, title, flags
    MessageBox = proto(("MessageBoxA", ctypes.windll.user32), params)
    return return_codes[MessageBox()]
    
def get_clipboard(strip=True):
    OpenClipboard = ctypes.windll.user32.OpenClipboard
    CloseClipboard = ctypes.windll.user32.CloseClipboard
    GetClipboardData = ctypes.windll.user32.GetClipboardData
    GetClipboardData.restype = ctypes.c_char_p
    if OpenClipboard(0):
        text = GetClipboardData(1)
        CloseClipboard()
        text = re.sub(' \r\n$', '', text) if strip else text
        return str(text)
    else:
        return ''

def set_clipboard(txt):
    '''Copies txt to clipboard which can then be retrieved with CRTL+v
    
    >>> set_clipboard("hello world")    
    >>> get_clipboard()
    'hello world'
    '''
    call('echo ' + txt + ' |clip', shell=True)
    
def set_monitor_power(state='off'):
    '''
    http://msdn.microsoft.com/en-us/library/windows/desktop/ms644950(v=vs.85).aspx
    '''
    SendMessage = ctypes.windll.user32.SendMessageA
    state = 1 if state.lower() == 'on' else 2
    SendMessage(65535, 274, 61808, state)

def get_rainbow_table(dirpath, hashfunc):
    '''
    Return dictionary with:
    --key: absolute path to file
    --value: hash of file using hashfunc
    '''
    def get_rainbow(dirpath):
        if path.exists(dirpath):
            for i in os.listdir(dirpath):
                pathName = '%s\\%s'%(dirpath, i)
                if path.isfile(pathName):
                    yield {pathName: hashfunc(pathName)}
                else:
                    for i in get_rainbow(pathName):
                        yield i
    table = {}
    for i in get_rainbow(dirpath):
        table.update(i)
    return table

def get_duplicate_files(dirpath, hashfunc=md5):
    '''
    Return nested list of absolute paths of duplicate files in dirpath.
    '''
    duplicates = []
    inverse = defaultdict(list) 
    data = get_rainbow_table(dirpath, hashfunc) 
    for k, v in data.items():
        inverse[v].append(k)
    for i in inverse.items():
        if len(i[1]) > 1:
            duplicates.append(i[1])
    return duplicates
    
if __name__ == '__main__':
    import doctest
    doctest.testmod(verbose=True)