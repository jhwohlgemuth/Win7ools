# -*- coding: utf-8 -*-
'''
ipl.py -- Class and methods for loading, editing, and saving
installed program lists (IPL) on Windows OS.
'''

import cPickle
import os
import re
import string
from distutils.log import warn
from pkgutil import iter_modules
import sqlite3
from win7ools.lib import get_closest_match as closest
from win7ools.lib import timestamp
from win7ools.reg import RegistryKeys, get_values
  
class IPL(object):
    '''
    Create, edit, and save Installed Program Lists (IPL)

    There are a few ways to work with the IPL class:
    
    You can create an IPL as a stand-alone dictionary,
    ex: 
        >> IPL = IPL.get()
    
    Or create an IPL object that gets the IPL automatically,
    ex: 
        >> info = IPL()
    
    Or create an IPL object with no IPL data... 
    (ex: info = IPL(data=[]) ==> info.IPL = {'Programs':[],'Modules': []})
    
    ...and add data separately,
    ex: 
        >> info.IPL = IPL
    
    You can create an IPL object and pass IPL data:
    ex: 
        >> info = IPL(data=VIPL)
    
    Or create an IPL object with no data and import data from an XML,
    ex: 
        >> info = IPL(data=[]).add_from_xml(xmlpath)
    
    Almost all IPL methods fully support chaining similar to jQuery.
    '''
    def __init__(self, data=None, load=False, keys=RegistryKeys()):
        '''
        Create an IPL instance.
        Data should be of the format: {'Programs':[],'Modules': []}.
        If data is excluded, IPL.get() will populate self.IPL,
        if data is not in the IPL format, self.IPL will be set to
        {'Programs':[],'Modules': []}
        '''
        self.name = os.environ['computername']
        self.date = timestamp()
        self.modified = self.date
        if data is not None:
            try:
                self.IPL = {'Programs': data['Programs'],\
                            'Modules': data['Modules']}
            except:
                self.IPL = {'Programs':[],'Modules': []}
        else:
            if load:
                self.IPL = IPL.load().IPL
            else:
                self.IPL = IPL.get(keys)
        
    def __repr__(self):
        '''Prints sentence describing self.'''
        name = self.name
        program_num = len(self.IPL['Programs'])
        iplStr = 'IPL for %s with %s programs installed.'%(name, program_num)
        return iplStr

    def __getitem__(self, item):
        '''Allows self['Programs'] and self['Modules']'''
        return self.IPL[item]
        
    def __len__(self):
        '''Returns number of items in 'Programs'.'''
        return len(self.IPL['Programs'])

    def __add__(self, other):
        '''Create an IPL object from the union of self.IPL and other.IPL.'''
        iplSum = IPL(data=[])
        for i in self.IPL['Programs']:
            iplSum.IPL['Programs'].append(i)
        for i in self.IPL['Modules']:
            iplSum.IPL['Modules'].append(i)
        for i in other.IPL['Programs']:
            if i['name'] not in [x['name'] for x in iplSum.IPL['Programs']]:
                iplSum.IPL['Programs'].append(i)
        for i in other.IPL['Modules']:
            if i not in iplSum.IPL['Modules']:
                iplSum.IPL['Modules'].append(i)
        return iplSum

    def __sub__(self, other):
        '''Return IPL object with items in self.IPL and not in other.IPL.'''
        diff = IPL(data=[])
        programs = [x for x in self['Programs'] if x not in other['Programs']]
        modules = [x for x in self['Modules'] if x not in other['Modules']]
        diff.IPL['Programs'] = programs
        diff.IPL['Modules'] = modules
        return diff

    def __eq__(self, l):
        '''
        Two IPL objects are equal if:
        --self['Programs'] == l['Programs'] 
        --self['Modules'] == l['Modules']
        
        >>> v = IPL()
        >>> w = IPL()
        >>> v == w
        True
        '''
        return self['Programs'] == l['Programs'] and self['Modules'] == l['Modules']

    def __contains__(self, item):
        '''
        Easily check if program is installed by checking if it is in self.IPL.
        
        >>> 'python' in IPL()
        True
        '''
        installed = False
        if isinstance(item, str):
            for program in self.IPL['Programs']:
                if item.lower() in program['name'].lower():
                    installed = True
                    break
            if not installed:
                for module in self.IPL['Modules']:
                    if item.lower() in module.lower():
                        installed = True
                        break
        return installed

    @classmethod
    def get(cls, keys=RegistryKeys()):
        '''
        Creates and returns an IPL from currently installed software.

        An IPL is a python dictionary with structure:
        {
            'Programs':
                [
                    {
                    'category': Category,
                    'optional': True | False,
                    'silent_install': WPI | Ninite | Other | None,
                    'backup_required': True | False,
                    'name': Name (= name of REG key or DisplayName),
                    'version': DisplayVersion,
                    'date': InstallDate,
                    'location': InstallLocation,
                    'GUID': GUID,
                    'data': Data
                    },...
                ]
            'Modules': [list of installed python modules]
        }
        '''
        def get_canonical(name):
            namel = []
            temp = re.sub('_', ' ', name)
            temp = re.sub('[(].*[)]', '', temp)
            temp = re.sub('[vV]ersion', '', temp)
            for i in re.split(' ', temp):
                n = re.match('[\D7][\d\D]+', i)
                version = re.search('[^\D][\d.]+', i)
                date = re.match('[12][90][019][\d]', i)
                if n and not version:
                    temp = string.strip(n.group(0))
                    if temp.lower() != 'v.':
                        namel.append(temp)
                if date:
                    temp = string.strip(date.group(0))
                    namel.append(temp)
            if string.join(namel) in name:
                rename = string.join(namel)
            else:
                rename = name
            if not rename:
                rename = name
            return rename

        def get_version_from(name):
            version = None
            temp = re.sub('_', ' ', name)
            temp = re.sub('[(].*[)]', '', temp)
            temp = re.sub('[vV]ersion', '', temp)
            temp = re.sub('[vV][.]','', temp)
            for i in re.split(' ',temp):
                v = re.search('[\d.]+\Z', i)
                if v:
                    if not re.match('[12][90][019][\d]', v.group(0)):
                        version = v.group(0)
                        break
            try:       
                v = re.match('[\d.]*\d', version)
                v.group(0)
            except:
                version = None
            return str(version)        
        
        IPL = {'Programs':[],'Modules': []}
        temp = []
        try:
            temp = get_values(keys.installer, HKCR=True) +\
                   get_values(keys.uninstall) +\
                   get_values(keys.uninstall, HKCU=True) +\
                   get_values(keys.uninstall_wow6432)                                 
            
            guid = '[{][\d\D]{8}-[\d\D]{4}-[\d\D]{4}-[\d\D]{4}-[\d\D]{12}[}].*'
            GUID_A = re.compile(guid)
            GUID_B = re.compile('[\d\D]{32}')
                                    
            for i in temp:
                registryName = i['RegistryName']
                Category = ''
                Optional = False
                SilentInstall = None
                SilentInstallTag = ''
                BackupRequired = False
                Name = ''
                Version = ''
                Date = ''
                Location = ''
                GUID = ''
                Data = ''
                if GUID_A.search(registryName) or GUID_B.search(registryName):
                    GUID = re.sub('[{}-]','', registryName)
                    try:
                        # First choice
                        Name = i['Values']['DisplayName']
                    except:
                        try:
                            # Second choice
                            Name = i['Values']['ProductName']
                        except:
                            Name = 'Not Found' 
                    finally:
                        Name = Name.encode('ascii','ignore')
                else:
                    try:
                        Name = i['Values']['DisplayName']
                    except:
                        Name = registryName
                    finally:
                        Name = Name.encode('ascii','ignore')
                try:
                    Version = i['Values']['DisplayVersion']
                    Version = Version.encode('ascii','ignore')
                except:
                    Version = get_version_from(Name)
                try:
                    Date = i['Values']['InstallDate'].encode('ascii','ignore')
                except:
                    Date = ''
                try:
                    Location = i['Values']['InstallLocation']
                    Location = Location.encode('ascii','ignore')
                except:
                    Location = ''
                    
                # Logitech SetPoint shows up as "eReg".
                if Name == 'eReg':
                    Name = 'Logitech SetPoint'
                
                Name = get_canonical(Name)
                
                data = {'category': Category,
                        'optional': Optional,
                        'silent_install': SilentInstall,
                        'silent_install_tag': SilentInstallTag,
                        'backup_required': BackupRequired,
                        'name': Name,
                        'version': Version,
                        'date': Date,
                        'location': Location,
                        'GUID': GUID,
                        'data': Data
                        }
                if Name not in [x['name'] for x in IPL['Programs']]:
                    IPL['Programs'].append(data)                                                 
        except:
            warn('Failed to retrieve IPL')        
        programs = sorted(IPL['Programs'], key=lambda p: p['name'].lower())
        modules = set([re.split('[.]',x[1])[0] for x in iter_modules()])
        IPL['Programs'] = programs
        IPL['Modules'] = sorted(modules)
        return IPL

    def clear(self):
        '''
        >>> ipl = IPL()
        >>> len(ipl) >= 0
        True
        >>> len(ipl.clear()) == 0
        True
        '''
        self.IPL = {'Programs':[],'Modules': []}
        return self
    
    def get_names(self, refine=''):
        '''
        Return list of names of programs from self.IPL
        Refine can be 'manual', 'silent', 'Ninite', or 'tag'        
        '''
        names = []
        for i in self['Programs']:
            names.append(i['name'])
        return sorted(names, key=lambda name: name.lower())
   
    def print_names(self):
        '''Print the names of installed programs.'''
        for name in self.get_names():
            print(name)
            
    def get_program(self, program):
        '''
        Return program information as dictionary:
            {
            'category': ...,
            'optional': ...,
            'silent_install': ...,
            'silent_install_tag': ...,
            'backup_required': ...,
            'name': program,
            'version': ...,
            'date': ...,
            'location': ...,
            'GUID': ...,
            'data': ...
            }
        '''
        info = []
        names = self.get_names()
        names = [name.lower() for name in names]
        name = closest(program.lower(), names)
        for i in self.IPL['Programs']:
            if name in i['name'].lower():
                info = i
                break
        return info

    def add_program(self,\
                    name='',\
                    category='',\
                    optional=False,\
                    silent_install=None,\
                    silent_install_tag=None,\
                    backup_required=False,\
                    version='',\
                    date='',\
                    location='',\
                    GUID='',\
                    data='',\
                    replace=False,\
                    update=False):
        '''
        Add a program to self['Programs'].
        
        >>> ipl = IPL(data=[])
        >>> ipl.add_program(name='Test Program')
        True
        >>> 'Test Program' in ipl
        True
        '''
        modified = False
        if name:
            data = {'category': category,\
                    'optional': optional,\
                    'silent_install': silent_install,\
                    'silent_install_tag': silent_install_tag,\
                    'backup_required': backup_required,\
                    'name': name,\
                    'version': version,\
                    'date': date,\
                    'location': location,\
                    'GUID': GUID,\
                    'data': data}
            if name not in self.get_names():
                self.IPL['Programs'].append(data)
                modified = True
            else:
                if replace:
                    self.remove_program(name)
                    self.IPL['Programs'].append(data)
                    modified = True
                if update:
                    index = self['Programs'].index(self.get_program(name))
                    old = self['Programs'][index]
                    new = data
                    for param in old.items():
                        if not param[1]:
                            old[param[0]] = new[param[0]]
                    modified = True                          
            if modified:
                self.modified = timestamp() 
        return name in self
        
    def add_modules(self, modName):
        '''
        Add a python module to self['Modules'].

        >>> ipl = IPL().clear()
        >>> ipl.add_modules('a')
        True
        >>> ipl['Modules']
        ['a']
        >>> ipl.add_modules(['b', 'c'])
        True
        >>> ipl['Modules']
        ['a', 'b', 'c']
        '''
        if isinstance(modName, str):
            if modName not in self.IPL['Modules']:
                self.IPL['Modules'].append(modName)
                self.modified = timestamp()
            return modName in self['Modules']
        if isinstance(modName, list):
            for mod in modName:
                if mod not in self.IPL['Modules']:
                    self.IPL['Modules'].append(mod)
            self.modified = timestamp()
            return set(self['Modules']).issuperset(set(modName))
                        
        
    def remove_program(self, name):
        '''
        Remove program (name=name) from self['Programs']
        
        >>> ipl = IPL(data=[])
        >>> ipl.add_program(name='Test')
        True
        >>> 'Test' in ipl
        True
        >>> ipl.remove_program('Test')
        True
        >>> 'Test' not in ipl
        True
        '''
        if name in self:
            index = self['Programs'].index(self.get_program(name))
            self['Programs'].pop(index)
            self.modified = timestamp()
        return name not in self

    def remove_module(self, modName):
        '''
        Remove module (name=modName) from self['Modules']
        
        >>> ipl = IPL().clear()
        >>> ipl.add_modules('a')
        True
        >>> 'a' in ipl['Modules']
        True
        >>> ipl.remove_module('a')
        True
        >>> 'a' not in ipl['Modules']
        True
        '''        
        if modName in self['Modules']:
            index = self['Modules'].index(modName)
            self['Modules'].pop(index)
            self.modified = timestamp()
        return modName not in self['Modules']

    @classmethod
    def installed(cls, module):
        '''
        >>> IPL.installed('os')
        True
        >>> IPL.installed('some weird module')
        False
        '''
        try:
            return bool(__import__(module))
        except:
            return False

    def save(self, path=None):
        '''
        Class method to save IPL object to sqlite3 database 
        in base module directory.
            ex: comp_ipl = IPL()
            ex: comp_ipl.save()
        '''
        if path:
            save_path = path
        else:
            save_path = os.path.dirname(__file__)
        save_path = '%s\\IPL.db'%save_path
        conn = sqlite3.connect(save_path)
        cursor = conn.cursor()
        cursor.execute('CREATE TABLE IF NOT EXISTS ipl(name, date, mod, data)')
        values = (self.name,
                  self.date,
                  self.modified,
                  cPickle.dumps(self))
        cursor.execute('INSERT INTO ipl VALUES (?, ?, ?, ?)', values)
        conn.commit()
        conn.close()
       
    @classmethod
    def load(cls, date='', path=''):
        '''
        Class method to load IPL object from sqlite3 database.
        
        Default load location is in the base package directory.
        
            ex: loaded_ipl = IPL.load()
        OR
            ex: loaded_ipl = IPL(load=True)
        OR
            ex: loaded_ipl = IPL(data=IPL.load(date='...').IPL)
        '''
        if path:
            load_path = path
        else:
            load_path = os.path.dirname(__file__)
        load_path = '%s\\IPL.db'%load_path
        conn = sqlite3.connect(load_path) 
        c = conn.cursor()
        ipl_data = ('', '', '', '')
        try:
            dates = c.execute('SELECT date FROM ipl')
            if not date:
                date = max(dates)
            c.execute('SELECT data FROM ipl WHERE date=?',\
                      date)
            ipl_data = c.fetchone()
        except:
            warn('Failed to load IPL data.')
        finally:
            conn.close()
        return cPickle.loads(str(ipl_data[0]))
                    
if __name__ == '__main__':
    import doctest
    doctest.testmod(verbose=True)