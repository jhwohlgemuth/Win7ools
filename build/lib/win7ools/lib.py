# -*- coding: utf-8 -*-
'''lib.py -- Collection of code snippets for the Windows OS'''

'''
Copyright 2014 Jason Wohlgemuth

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
'''

from ctypes import windll
import types
import os
import pickle
import PIL
import re
import shutil
import string
import subprocess
import time
from collections import Counter
from difflib import get_close_matches
from distutils.log import warn
from functools import wraps
from math import sqrt

class Profiled:
    def __init__(self, func):
        wraps(func)(self)
        self.called = 0
        self.runtimes = []
        self.average_runtime = 0
        self.last_run = 0
        self.wrapped = func
    def __call__(self, *args, **kwargs):
        self.called += 1
        start = time.time()
        result = self.wrapped(*args, **kwargs)
        end = time.time()
        self.runtimes.append(end - start)
        self.average_runtime = sum(self.runtimes) / float(self.called)
        self.last_run = self.runtimes[self.called - 1]
        return result
    def __get__(self, instance, cls):
        if instance is None:
            return self
        else:
            return types.MethodType(self, instance)

def is_hidden(path):
    return bool(windll.kernel32.GetFileAttributesW(unicode(path)) & 2)

def hide(item):
    '''Sets hide attibute of item'''
    p = os.popen('attrib +h ' + item)
    p.read()
    p.close()
    return is_hidden(item)

def unhide(item):
    '''Reverses hide(item)'''
    p = os.popen('attrib -h ' + item)
    p.read()
    p.close()
    return not is_hidden(item)
    
def timestamp():
    return time.strftime("%y%m%d%H%M", time.localtime())

def log(info, filename=''):
    system_drive = os.environ['systemdrive']
    computer_name = os.environ['computername']
    if os.path.exists(filename):
        fname = filename
    else:
        fname = os.path.join(system_drive, '\\' + computer_name + '_log.txt')
    t = time.strftime("%d%b%y %H:%M", time.localtime())
    t = string.upper(t)
    log = open(fname,'a')
    log.write("(%s) %s\n"%(t,info))
    warn("(%s) %s"%(t,info))
    log.close()

def get_day_name(year, month, day):
    '''
    Uses datetime.date to return name of day specified.
        
    >>> get_day_name(2008,12,26)
    'Friday'
    '''
    from datetime import date
    return date(year, month, day).strftime("%A")

def hex_to_str(hexval):
    '''
    >>> hex(46)
    '0x2e'
    >>> hex_to_str(hex(46))
    '2E'
    '''
    if len(hexval) == 3:
        hexstr = re.sub('x', '', hexval)
    else:
        hexstr = re.sub('0x', '', hexval)
    return hexstr.upper() 
    
def truncate(f, n):
    '''
    Truncates/pads a float f to n decimal places without rounding. 
    
    >>> str(truncate(3.1415, 1))
    '3.1'
    '''
    slen = len('%.*f'%(n, f))
    return float(str(f)[:slen])

def segment(items, n):
    '''
    Segment a list of items into a list of lists with length = n
    
    >>> segment([0, 1, 2, 3, 4, 5, 6, 7, 8, 9], 2)
    [[0, 1], [2, 3], [4, 5], [6, 7], [8, 9]]
    '''
    if n >= 1 and isinstance(n, int):
        return [items[i:i + n] for i in range(0, len(items), n)]
    else:
        warn('The segment function only excepts integer values >= 1.')

def flatten(*items):
    '''
    Flatten and combine all elements of items as a list.
    Does not require module import and flattens indefinite list nesting.
    
    >>> flatten([1,2,3,[4,5,[6,7,[8]]]])  
    [1, 2, 3, 4, 5, 6, 7, 8]
    
    >>> flatten(['abc',['def']])
    ['abc', 'def']
    
    >>> flatten('some string')
    'some string'
    '''
    def flattened(x):
        '''Flattens single level of nesting.'''
        flat = []
        for el in x:
            if hasattr(el, "__iter__") and not isinstance(el, str):
                flat.extend(flattened(el))
            else:
                flat.append(el)
        return flat
            
    if len(items) is 1 and isinstance(items[0], str): return items[0]
    result = []
    for item in items:
        result.extend(flattened(item))
    return result

def hamming_distance(strA, strB):
    '''
    Calculate the Hamming distance of two strings.
    
    >>> hamming_distance('hello world', 'hello world')
    0
    >>> hamming_distance('abcde', 'abcdf')
    1
    >>> hamming_distance('Abcde', 'abcdf')
    2
    >>> hamming_distance(12345, 54321)
    4
    '''
    strA = str(strA)
    strB = str(strB)
    if len(strA) is len(strB):
        return sum(charA != charB for charA, charB in zip(strA, strB))
    else:
        warn('Input strings are not the same length.')

def tanimoto():
    return 0

def levenshtein():
    return 0
 
def get_closest_match(a, items, cutoff=0.9):
    '''
    >>> names = ['joe', 'joel', 'joseph']
    >>> get_closest_match('jo', names)
    'joe'
    >>> get_closest_match('jos', names)
    'joseph'
    '''
    match = get_close_matches(a, items, 1, cutoff)
    while not match:
        match = get_closest_match(a, items, cutoff - 0.01)
    match = re.sub("[\[\]\'\"]", "", str(match))
    return match
  
def get_most_frequent(stuff):
    '''
    Return most frequent item the the stuff.
    
    >>> get_most_frequent([1,2,2,3,3,3])
    3
    
    >>> get_most_frequent([1,1,1,2,2,2])
    1

    >>> get_most_frequent([2,2,2,3,3,3])
    2
    
    >>> get_most_frequent({'a':1,'b':2,'b':2})
    'b'
    '''
    freq = Counter(stuff)
    most_freq = sorted(freq.items(), key=lambda x: x[1], reverse=True)[0][0] 
    return most_freq

def readability(text, debug=False):
    '''
    Calculates the readability of input text using:
        - Automated Readability Index (ARI)
        - Coleman-Liau Index (CLI)
        - Simple Measure of Gobbledygook (SMOG)*
        -- This function uses a slightly modified version of SMOG in that
           instead of using the number of "complex" words, it uses the number
           of words with more than 8 letters.  This simplifies the calculation
           while maintaining most of the algorithm's potentcy.  Most words with
           more than 8 letters will also have 3 syllables and therefore be
           'complex'.  There seems to only be one word that starts with 'a',
           has more that 8 letters and has 2 syllables, 'AARDVARKS'.  In any
           case, even among the few words with more than 8 letters and less
           than 3 syllable, only a fraction are actually not 'complex'.
                  
    This function accepts a string of text, or a path to a text file,
    it returns the grade (U.S.) reading level of the text.
    '''
    if os.path.exists(text):
        f = open(text, 'r')
        txt = f.read()
        f.close()
    else:
        txt = text
    letters = len(re.sub('[^a-zA-Z]', '', txt))
    word_list = re.sub('\s+', ' ', txt).split(' ')
    words = len(word_list)
    words_complex = 0
    for word in word_list:
        if len(word) > 9:
            words_complex += 1
    #remove Mr. and Mrs.
    pattern_mr = '[mM][rR]{0,1}[sS]{0,1}[.]'
    txt_formatted = re.sub(pattern_mr, 'mr', txt)
    #replace acronyms
    pattern_acronym = re.compile('([a-zA-Z][.])+[a-zA-Z][.]')
    txt_formatted = re.sub(pattern_acronym, '.', txt_formatted)
    sentence_list = txt_formatted.split('.')
    sentences = len(sentence_list)
    if txt[-1:] is '.':
        sentences -= 1     
    L = (float(letters) / words) * 100
    S = (float(sentences) / words) * 100
    CLI = (0.0588 * L) - (0.296 * S) - 15.8
    CLI = CLI if CLI >= 0 else 0
    ARI = 4.71 * float(letters)/words + 0.5 * float(words)/sentences - 21.43
    ARI = ARI if ARI >= 0 else 0
    SMOG = 1.043 * sqrt(words_complex * (30/sentences)) + 3.1291
    SMOG = SMOG if SMOG >= 0 else 0
    if debug:
        print('ARI: %f'%ARI)
        print('CLI: %f'%CLI)
        print('SMOG: %f'%SMOG)
        print('letters: %i'%letters)
        print('words: %i'%words)
        print('sentences: %i'%sentences)
    return round(float(ARI + CLI + SMOG) / 3, 2)

def copy(src, dst, overwrite_older=True):
    '''
    Copy src (file or dir) to dst (dir). If src exists in dst, it will only
    be copied if it is newer unless overwrite_older=False.
    '''
    if os.path.isdir(src):
        if not os.path.exists(dst):
            os.makedirs(dst)
        src_base = os.path.basename(src)
        destination = os.path.join(dst, src_base)
        if os.path.exists(destination):
            for i in os.listdir(src):
                copy(os.path.join(src, i), os.path.join(dst, src_base))
        else:
            shutil.copytree(src,destination)
    if os.path.isfile(src):
        destination = os.path.join(dst, os.path.basename(src))
        if os.path.exists(destination):
            time_modified_dst = os.path.getmtime(destination)
            time_modified_src = os.path.getmtime(src)           
            if time_modified_dst < time_modified_src and overwrite_older:
                os.remove(destination)
                shutil.copy(src, destination)
            else:
                pass
        else:
            if not os.path.exists(dst):
                os.makedirs(dst)
            shutil.copy(src, destination)

def copy_contents(src, dst):
    '''Copy the contents of directory, src, into the directory, dst.'''
    if os.path.exists(src):
        if not os.path.exists(dst):
            os.makedirs(dst)
        if os.path.isdir(src):
            for i in os.listdir(src):
                if not os.path.exists(dst):
                    os.makedirs(dst)
                copy(os.path.join(src, i), dst)
        else:
            warn('Cannot use copy_contents on file.')
            pass
    else:
        warn(src + ' does not exist.')
    
@Profiled
def find_dir(dirname, searchDir='', find_all=False):
    '''Returns highest level directory named "dirname"'''
    def filterDirs(dirname, directory, files):
        if os.path.basename(directory).upper() == dirname.upper():
            count = 0
            for i in os.path.split(directory.upper()): 
                count += i.count(dirname.upper())
            if count == 1: #Excludes "arg/arg/", etc...
                dirs.append(directory)        
    dirs = []
    startDir = os.getcwd()
    if not searchDir:
        searchDir = os.path.splitdrive(startDir)[0]  + '\\'
        print("Searching for \"%s\" in %s..."%(dirname, searchDir))
        print("This may take several minutes...")
    else:
        print("Searching for \"%s\" in %s..."%(dirname, searchDir))
    os.chdir(searchDir)
    os.path.walk(searchDir, filterDirs, dirname)
    if not dirs:
        dirs.append(dirname + ' Not Found in ' + searchDir)
    os.chdir(startDir)
    return dirs if find_all else dirs[0] 

@Profiled
def find(filename, searchDir='', find_all=False):
    def filterFiles(filename, directory, files):
        if filename in files:
            foundFiles.append(directory + '\\' + files[files.index(filename)])
    foundFiles = []
    startDir = os.getcwd()
    if not searchDir:
        searchDir = os.path.splitdrive(startDir)[0] + '\\'
        print("Searching for \"%s\" in %s..."%(filename, searchDir))
        print("This may take several minutes...")
    else:
        pass
    os.path.walk(searchDir, filterFiles, filename)
    if not foundFiles:
        foundFiles.append(filename + ' Not Found in ' + searchDir)
    return foundFiles if find_all else foundFiles[0]

def mklink(real, link):
    successStatus = ''
    arg = '/H' #Default set to hard link
    if not os.path.isdir(link):
        os.mkdir(link)
    link = link + '\\' + os.path.basename(real)
    realPath = '\"%s\"'%real
    linkPath = '\"%s\"'%link
    if os.path.isdir(real):
        arg = '/J'
        successStatus = "Junction created for " + link
    else:
        successStatus = "Hardlink created for " + link 
    cmd = 'mklink ' + arg + ' ' + linkPath + ' ' + realPath
    if os.path.exists(link):
        if os.path.isfile(link):
            os.remove(link)
        if os.path.isdir(link):
            shutil.rmtree(link)
    proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    for line in proc.stdout:
        if successStatus not in line:
            successStatus = "Failed to execute " + cmd
    unhide(linkPath)
    log(successStatus)
 
@Profiled  
def img_dumps(imgPath):
    '''Serialize image object to string.'''
    image = ''
    result = ''
    name = os.path.splitext(os.path.basename(imgPath))[0]
    img = PIL.Image.open(imgPath)
    if PIL.Image.isImageType(img):
        image = {'name' : name,
                 'pixels': img.tostring(),
                 'size': img.size,
                 'mode': img.mode}
        result = pickle.dumps(image)
    return result

@Profiled
def img_loads(imgStr):
    img = pickle.loads(imgStr)
    image = PIL.Image.fromstring(img['mode'], img['size'], img['pixels'])
    return image
    
def load_dictionary():    
    ospd4 = r"C:\Python27\Lib\site-packages\wintools\doc\OSPD4.txt"    
    f = open(ospd4, 'r')
    lines = f.readlines()
    f.close()   
    definitions = {}   
    for line in lines:
        part_of_speech = re.search('\[.*\]', line).group(0)[1:-1]
        part_of_speech = part_of_speech.split(' ')[0]
        word = line.split(' ')[0]
        word_root = re.search('<.*>', line)
        word_root = word_root.group(0)[1:-1] if word_root else word
        if word_root is not word:
            word_root = word_root.split('=')
        else:
            word_root = [word, part_of_speech]
        definition = re.sub('(<.*>)|(\[.*\])|(%s)'%word, '', line).strip()
        if not re.sub(' ', '', definition):
            definition = word_root
        definitions.update({word.lower(): [definition, part_of_speech, word_root]})  
    return definitions

if __name__ == '__main__':
    import doctest
    doctest.testmod(verbose=True)