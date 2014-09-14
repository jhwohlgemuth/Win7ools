# -*- coding: utf-8 -*-
"""xml.py - XML and networking support for wintools"""

import os
import re
from urllib import urlencode
from urllib import urlopen
from distutils.log import warn
from xml.etree.cElementTree import parse
from xml.etree.cElementTree import Element
from xml.etree.cElementTree import tostring
from .lib import get_most_frequent


class Crawler(object):
    '''
    Class to create web crawlers to gather data from the internet.
    May use some combination of bs4 and pyquery.
    Example uses:
    ** crawl web to develop dictionary of words
    ** crawl links on page to create site map and connections
    ** crawl web to find specified information
    '''
    def __init__(self):
        print('I am spider-man')

def parse_xml(xmlpath):
    doc = ''
    if os.path.exists(xmlpath):
        doc = parse(xmlpath)
    else:
        try:
            doc = parse(urlopen(xmlpath))
        except:
            warn('Failed to parse %s'%xmlpath)
    return doc

def get_xml_type(xmlpath, item='item'):
    '''
    Return whether the XML at xmlpath is meant to be imported as
    <type 'dict'> or <type 'list'> as defined by formats used in 
    get_from_xml and save_to_xml.
    If some but not all items have key assigned, returns None.
    '''
    doc = parse_xml(xmlpath)
    items = doc.findall(item)
    count = 0
    data_type = None
    if items:
        for i in items:
            if i.get('key'):
                count += 1
        if count == len(items):
            data_type = dict
        elif count == 0:
            data_type = list
    return data_type

def get_from_xml(xmlpath, item='item'):
    '''
    Take XML file as input (local or remote) 
    and convert the items to a python list or dict 
    (depending on XML file).
    
    XML input format:
        <?xml...?>
        <items>
            <item key=''>value1</item>
            <item key=''>value2</item>
            .
            .
            .
        </items>
        
    Python list output:
        [value1, value2, ...]
        
    XML input format:
        <?xml...?>
        <items>
            <item key='key1'>value1</item>
            <item key='key2'>value2</item>
            .
            .
            .
        </items>       
        
    Python dictionary output:
        {key1: value1, key2: value2, ...}
    '''
    doc = parse_xml(xmlpath)
    xml_type = get_xml_type(xmlpath, item)
    items = None
    if doc and xml_type:
        if isinstance([], xml_type):
            items = []
            for item in doc.findall(item):
                items.append(item.text)
        if isinstance({}, xml_type):
            items = {}
            for item in doc.findall(item):
                items.update({item.get('key'): item.text})
    return items

def save_to_xml(data, xmlpath='', item='item'):
    '''
    Take python list or dictionary and convert to XML, saved as xmlpath.
    
    Python list input:
        [item1, item2, ...]
        
    XML output format:
        <?xml...?>
        <list>
            <items>
                <item key=''>item1</item>
                <item key=''>item2</item>
                .
                .
                .
            </items>
        </list>
        
    Python dictionary input:
        {item1: value1, item2: value2, ...}
        
    XML output format:
        <?xml...?>
        <list>
            <items>
                <item key='item1'>value1</item>
                <item key='item2'>value2</item>
                .
                .
                .
            </items>
        </list>
    '''            
    def dict_to_xml(d, item):
        el = Element('data')
        for key, val in d.items():
            child = Element(item)
            child.text = str(val)
            child.set('key', key)
            el.append(child)
        print(tostring(el))

    def list_to_xml(l, item):
        el = Element('data')
        for val in l:
            child = Element(item)
            child.text = str(val)
            el.append(child)
        print(tostring(el))
        
    if isinstance(data, dict):
        dict_to_xml(data, item)
    elif isinstance(data, list):
        list_to_xml(data, item)
    else:
        warn('Please, only dicts or lists...')
       
def get_external_ip():  
    '''
    Returns external IP as string.
    
    ***REQUIRES INTERNET CONNECTION***
    '''
    IP = r"\d{1,3}\.\d{1,3}\.\d{1,3}.\d{1,3}"
    canyouseeme = "http://www.canyouseeme.org/"
    jsonip = "http://jsonip.com/"
    icanhazip = "http://icanhazip.com/"
    checkip = "http://checkip.dyndns.org"
    addresses = [canyouseeme, jsonip, icanhazip, checkip]
    extIP = ['']
    for i in addresses:
        try:
            f = urlopen(i)
            html_doc = f.read()
            f.close()
            m = re.search(IP ,html_doc)
            if m:
                extIP.append(m.group(0))
        except:
            pass
    return get_most_frequent(extIP)
    
def get_local_ip():
    '''Returns local IP address as a string.'''
    import socket
    return socket.gethostbyname(socket.gethostname())         

def get_carrier(phoneNumber):
    '''
    Uses fonefinder.net to determine cell phone carrier.
    Will return incorrect carrier for ported cell numbers.
    
    ***REQUIRES INTERNET CONNECTION***
    '''
    npa = phoneNumber[:3]
    nxx = phoneNumber[3:6]
    thoublock = phoneNumber[6:]
    data = urlencode({'npa': npa,\
                      'nxx': nxx,\
                      'thoublock': thoublock})
    f = urlopen('http://www.fonefinder.net/findome.php?' + data).read()
    m = re.search('http://fonefinder[.]net/[abcdefghijklmnopqrstuvwxyz]*[.]php', f)
    carrier = ''
    if m:
        carrier = re.sub('http://fonefinder.net/', '', m.group())
        carrier = re.sub('.php', '', carrier)
    return carrier.upper()
    
    
    