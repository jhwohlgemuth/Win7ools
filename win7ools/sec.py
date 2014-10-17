# -*- coding: utf-8 -*-
"""secure.py -- AES Encryption and other security related functions"""

import codecs
from ctypes import windll, cdll, c_buffer, byref, POINTER, Structure, c_char
from ctypes.wintypes import DWORD
import hashlib
import hmac
import os.path as path
import string
from Crypto.Cipher import AES
from Crypto import Random

BLOCK_SIZE = 256*128

#passwords_2011 lists the most common passwords from 2011 as listed on the
#forum of LearnHacking.in
passwords_2011 = ['123456',
                  '12345678',
                  'qwerty',
                  'abc123',
                  'monkey',
                  '1234567',
                  'letmein',
                  'trustno1',
                  'dragon',
                  'baseball',
                  '111111',
                  'iloveyou',
                  'master',
                  'sunshine',
                  'ashley',
                  'bailey',
                  'passw0rd',
                  'shadow',
                  '123123',
                  '654321',
                  'superman',
                  'qazwsx',
                  'michael',
                  'football']

def md5(hashpath, key='', block_size=BLOCK_SIZE):
    '''
    Block size directly depends on the block size of your filesystem
    to avoid performances issues
    Here I have blocks of 4096 octets (Default NTFS)
    
    If key is designated, HMAC will be utilized with MD5 hashing.
    
    >>> md5('abcdefghijklmnopqrstuvwxyz')
    'c3fcd3d76192e4007dfb496cca67e13b'
    
    >>> md5('hello world')
    '5eb63bbbe01eeed093cb22bb8f5acdc3'
    
    >>> md5('hello world', 'key')
    'ae92cf51adf91130130aefc2b39a7595'
    '''
    if key:
        MD5 = hmac.new(str(key), digestmod=hashlib.md5)
    else:
        MD5 = hashlib.md5()
    if path.exists(hashpath):
        with open(hashpath,'rb') as f: 
            for chunk in iter(lambda: f.read(block_size), b''): 
                 MD5.update(chunk)
    else:
        MD5.update(hashpath)
    return MD5.hexdigest()
    
def sha1(hashpath, key='', block_size=BLOCK_SIZE):
    '''
    Block size directly depends on the block size of your filesystem
    to avoid performances issues
    Here I have blocks of 4096 octets (Default NTFS)
    
    If key is designated, HMAC will be utilized with SHA1 hashing.
    
    >>> sha1('abcdefghijklmnopqrstuvwxyz')  
    '32d10c7b8cf96570ca04ce37f2a19d84240d3a89'
    
    >>> sha1('hello world')
    '2aae6c35c94fcfb415dbe95f408b9ce91ee846ed'
    
    >>> sha1('hello world', 'key')
    '34dd234b92683593560528f6193ea68c8005f615'
    '''
    if key:
        SHA1 = hmac.new(str(key), digestmod=hashlib.sha1)
    else:
        SHA1 = hashlib.sha1()
    if path.exists(hashpath):
        with open(hashpath,'rb') as f: 
            for chunk in iter(lambda: f.read(block_size), b''): 
                 SHA1.update(chunk)
    else:
        SHA1.update(hashpath)
    return SHA1.hexdigest()

def sha256(hashpath, key='', block_size=BLOCK_SIZE):
    '''
    Block size directly depends on the block size of your filesystem
    to avoid performances issues
    Here I have blocks of 4096 octets (Default NTFS)
    
    If key is designated, HMAC will be utilized with SHA256 hashing.

    >>> len(sha256('hello world'))
    64
    
    >>> sha256('hello world', 'key')
    '0ba06f1f9a6300461e43454535dc3c4223e47b1d357073d7536eae90ec095be1'
    '''
    if key:
        SHA = hmac.new(str(key), digestmod=hashlib.sha256)
    else:
        SHA = hashlib.sha256()
    if path.exists(hashpath):
        with open(hashpath,'rb') as f: 
            for chunk in iter(lambda: f.read(block_size), b''): 
                 SHA.update(chunk)
    else:
        SHA.update(hashpath)
    return SHA.hexdigest()

def sha512(hashpath, key='', block_size=BLOCK_SIZE):
    '''
    Block size directly depends on the block size of your filesystem
    to avoid performances issues
    Here I have blocks of 4096 octets (Default NTFS)
    
    If key is designated, HMAC will be utilized with SHA512 hashing.
    
    >>> len(sha512('hello world'))
    128
    
    >>> sha512('hello world', 'key')
    'ea0625a5ff1cd1653a327f8a4ae2f478fc51405c73ddac3a8a05a7a810310a6a14d7c8b4d284013493a6016ecadc772cfd98ed6cbe745949c5e6119fafb63b54'
    '''
    if key:
        SHA = hmac.new(str(key), digestmod=hashlib.sha512)
    else:
        SHA = hashlib.sha512()
    if path.exists(hashpath):
        with open(hashpath,'rb') as f: 
            for chunk in iter(lambda: f.read(block_size), b''): 
                 SHA.update(chunk)
    else:
        SHA.update(hashpath)
    return SHA.hexdigest()

def phash(image_path, hash_size = 8):
    '''
    Based on blog post by Silviu Tantos, a back-end developer at iconfinder.com
    http://blog.iconfinder.com/detecting-duplicate-images-using-python/
    Basic idea:
        - convert the image to grayscale with PIL
        - shrink the image to a common size (hash_size)
        - compare adjacent pixels (using difference)
        - convert comparison into bits
        - return hash of bits (hexadecimal string)
    '''
    from PIL import Image
    # Grayscale and shrink the image in one step.
    img = Image.open(image_path)
    img = img.convert('L').resize((hash_size + 1, hash_size), Image.ANTIALIAS,)
    # Compare adjacent pixels.
    difference = []
    for row in xrange(hash_size):
        for col in xrange(hash_size):
            pixel_left = img.getpixel((col, row))
            pixel_right = img.getpixel((col + 1, row))
            difference.append(pixel_left > pixel_right)
    # Convert the binary array to a hexadecimal string.
    decimal_value = 0
    hex_string = []
    for index, value in enumerate(difference):
        if value:
            decimal_value += 2**(index % 8)
        if (index % 8) == 7:
            hex_string.append(hex(decimal_value)[2:].rjust(2, '0'))
            decimal_value = 0
 
    return ''.join(hex_string)
   
def encrypt(key, data, replace=True, hash_key='', prefix='', suffix='_lock'):
    '''
    Encrypt string of data using AES. 
    KEY=md5(str(key)) creates a 32-byte (256-bit) key for encryption.
    data is either a string to be encrypted or the path to file.   
    Tested on strings, multiline text files, and JPEGs.
    
    ***This function writes over original file with decrypted version unless
    replace is set to False***
    
    >>> ptext = "Hello World!"
    >>> ctext = encrypt('password', ptext)
    >>> ptext != ctext
    True
    >>> ptext == decrypt('password', ctext)
    True
    '''
    key = str(key)
    iv = Random.new().read(AES.block_size)
    KEY = md5(str(key), hash_key)
    
    cipher = AES.new(KEY, AES.MODE_CFB, iv)
    if path.exists(data):
        with open(data, 'rb') as f:
            info = f.read()
        data_path = data if replace else '%s%s%s.txt'%(prefix, data[:-4], suffix)
        with open(data_path, 'wb') as f:
            f.write(iv + cipher.encrypt(info))
        return data_path
    else:
        msg = iv + cipher.encrypt(data)
        return msg    
    
def decrypt(key, data, replace=True, hash_key='', prefix='', suffix='_unlock'):
    '''
    Decrypt string of data using AES. 
    KEY=md5(str(key)) creates a 32-byte (256-bit) key for decryption.
    data is the path to the file to be encrypted.    
    Tested on strings, multiline text files, and JPEGs.
    
    ***This function writes over original file with decrypted version unless
    replace is set to False***
    '''
    key = str(key)
    iv = Random.new().read(AES.block_size)
    KEY = md5(str(key), hash_key)
    
    cipher = AES.new(KEY, AES.MODE_CFB, iv)
    if path.exists(data):
        with open(data, 'rb') as f:
            info = f.read()
        data_path = data if replace else '%s%s%s.txt'%(prefix, data[:-4], suffix)
        with open(data_path, 'wb') as f:
            f.write(cipher.decrypt(info)[AES.block_size:])
        return data_path    
    else:
        msg = cipher.decrypt(data)[AES.block_size:]
        return msg     

def get_dictionary(_from, _to):
    english = list(string.ascii_lowercase + string.digits)
    leet = ['4', '8', 'c', 'd', '3', '=', '9', '#', '!', '_', 'k', '1', 'm',\
            '~', '0', 'D', 'O', '2', '5', '7', 'u', 'V', 'w', 'x', "%", 'z']     
    #abcdefghijklmnopqrstuvwxyz0123456789              
    morse = ['._', '_...', '_._.', '_..', '.', '.._.', '__.', '....',\
             '..', '.___', '_._', '._..', '__', '_.', '___', '.__.', '__._',\
             '._.', '...', '_', '.._', '..._', '.__', '_.._', '_.__', '__..',\
             '_____', '.____', '..___', '...__','...._', '.....', '_....',\
             '__...', '___..', '____.']
    language = {"english": english, "leet": leet, "morse": morse}
    return {x:y for (x,y) in zip(language[_from], language[_to])}                     
                     
                     
alphabet_eng = list(string.ascii_lowercase + string.digits)
alphabet_leet = ['4', '8', 'c', 'd', '3', '=', '9', '#', '!', '_', 'k',
                '1', 'm', '~', '0', 'D', 'O', '2', '5', '7', 'u',
                'V', 'w', 'x', "%", 'z']   
                
#abcdefghijklmnopqrstuvwxyz0123456789              
alphabet_morse = ['._', '_...', '_._.', '_..', '.', '.._.', '__.', '....',
                 '..', '.___', '_._', '._..', '__', '_.', '___', '.__.',
                 '__._', '._.', '...', '_', '.._', '..._', '.__', '_.._',
                 '_.__', '__..', '_____', '.____', '..___', '...__', '...._',
                 '.....', '_....', '__...', '___..', '____.']

def encode(plain_text, encoding='leet', separator=' '):
    '''
    Encodes plain_text with input encoding (default 'encoding' is LEET)
    
    >>> encode('hello world')
    '#3110 w021d'
    
    >>> encode('hello world', encoding='morse')
    '.... . ._.. ._.. ___  .__ ___ ._. ._.. _..'
    
    >>> encode('hello world', encoding='rot13')
    'uryyb jbeyq'
    '''  
    ctext = []
    if encoding.lower() == 'leet':
        ptext = plain_text.lower()
        dictionary = get_dictionary('english', encoding)
        for letter in list(ptext):
            try:
                ctext.append(dictionary[letter])
            except:
                ctext.append(letter)               
    elif encoding == 'morse':
        ptext = plain_text.lower()
        dictionary = get_dictionary('english', encoding)
        for letter in list(ptext):
            try:
                ctext.append(dictionary[letter] + separator)
            except:
                ctext.append(letter)              
    elif encoding == 'rot13':
        ctext = codecs.encode(plain_text, 'rot13')                             
    return ''.join(ctext).strip() 

def decode(cipher_text, encoding='leet', separator=' '):
    '''
    Decodes cipher_text with input encoding (default 'encoding' is LEET)
    
    >>> decode('#3110 w021d')
    'hello world'
    
    >>> decode('.... . ._.. ._.. ___  .__ ___ ._. ._.. _..', encoding='morse')
    'helloworld'
    
    >>> decode('uryyb jbeyq', encoding='rot13')
    'hello world'
    ''' 
    ptext = []
    if encoding.lower() == 'leet':
        ctext = cipher_text.lower()
        dictionary = get_dictionary(encoding, 'english')
        for letter in list(ctext):
            try:
                ptext.append(dictionary[letter])
            except:
                ptext.append(letter)               
    elif encoding == 'morse':
        ctext = cipher_text.split(' ')
        dictionary = get_dictionary(encoding, 'english')
        for i in ctext:
            try:
                ptext.append(dictionary[i])
            except:
                ptext.append(i)          
    elif encoding == 'rot13':
        ptext = codecs.encode(cipher_text, 'rot13')                 
    return ''.join(ptext) 

def get_password_strength(password):
    return len(password)

class BLOB(Structure):
    '''
    BLOB class is used in crypt_protect_data and crypt_unprotect_data methods
    '''
    _fields_ = [("cbData", DWORD), ("pbData", POINTER(c_char))]
    
    def data(self):
        LocalFree = windll.kernel32.LocalFree
        memcpy = cdll.msvcrt.memcpy
        cbData = int(self.cbData)
        pbData = self.pbData
        buffer = c_buffer(cbData)
        memcpy(buffer, pbData, cbData)
        LocalFree(pbData);
        return buffer.raw

def crypt_protect_data(ptext, entropy=None, description=u''):
    '''
    This function uses the Windows function, CryptProtectData, to encrypt
    ptext with Triple DES encryption 
    IAW with http://msdn.microsoft.com/en-us/library/ms995355.aspx.
    
    It was heavily influenced by:
    Crusher Joe - http://article.gmane.org/gmane.comp.python.ctypes/420
    
    >>> msg = 'hello world'
    >>> cipher_text = crypt_protect_data(msg)
    >>> cipher_text is not msg
    True
    >>> crypt_unprotect_data(cipher_text)
    'hello world'
    
    >>> msg = 'hello world'
    >>> entropy = '1234567890'
    >>> cipher_text = crypt_protect_data(msg, entropy)
    >>> cipher_text is not msg
    True
    >>> crypt_unprotect_data(cipher_text, entropy)
    'hello world'
    '''
    ctext = ''
    CryptProtectData = windll.crypt32.CryptProtectData
    CRYPTPROTECT_UI_FORBIDDEN = 0x01
    blob_in = BLOB(len(ptext), c_buffer(ptext, len(ptext)))
    if entropy:
        buffer_entropy = c_buffer(entropy, len(entropy))
        blob_entropy = byref(BLOB(len(entropy), buffer_entropy))
    else:
        blob_entropy = entropy
    blob_out = BLOB()
    if CryptProtectData(byref(blob_in), 
                        description, #readable description of data
                        blob_entropy, #additional entropy for encryption
                        None, #reserved for future use
                        None, #used to specify settings for user prompt
                        CRYPTPROTECT_UI_FORBIDDEN, #user UI is not allowed
                        byref(blob_out)):
        ctext = blob_out.data()
    return ctext

def crypt_unprotect_data(ctext, entropy=None):
    '''
    This function uses the Windows function, CryptUnprotectData, to decrypt
    ctext IAW with http://msdn.microsoft.com/en-us/library/ms995355.aspx.
    
    It was heavily influenced by:
    Crusher Joe - http://article.gmane.org/gmane.comp.python.ctypes/420
    '''    
    ptext = ''
    CryptUnprotectData = windll.crypt32.CryptUnprotectData
    CRYPTPROTECT_UI_FORBIDDEN = 0x01
    blob_in = BLOB(len(ctext), c_buffer(ctext, len(ctext)))
    if entropy:
        buffer_entropy = c_buffer(entropy, len(entropy)) 
        blob_entropy = byref(BLOB(len(entropy), buffer_entropy))
    else:
        blob_entropy = entropy
    blob_out = BLOB()
    if CryptUnprotectData(byref(blob_in), 
                          None, #pointer to readable description
                          blob_entropy, 
                          None, 
                          None,
                          CRYPTPROTECT_UI_FORBIDDEN, 
                          byref(blob_out)):
        ptext = blob_out.data()
    return ptext

if __name__ == '__main__':
    import doctest
    doctest.testmod(verbose=True)