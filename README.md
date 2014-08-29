Win7ools Project
================

**Win7ools** (= *Win7 Tools* + [leet][leet]) provides programmatic access to the Windows
Operating System. Win7ools was born out of the desire to list and track installed software 
programmatically and evolved into an attempt at a full-featured window into...Windows.  Win7ools 
is organized into a handful of modules - **lib**, **ipl**, **sec**, **pdf**, **web**, and **reg**, 
and the package, **system**.  The general idea is to be able to do things that normally would require
a specialized piece of installed software that works mysteriously.

modules
-------
**ipl**
  
 <b>I</b>nstalled <b>P</b>rograms <b>L</b>ist - the **ipl** module scans the Windows registry and returns
 information related all installed programs on the parent OS.  Comparatively, it should
 find all the information that the control panel provides(sometimes more).  IPLs can be created
 from an active OS, saved for various purposes, created from scratch, merged, and loaded.

**lib**

> A <b>lib</b>rary of useful code snippets including, but not limited to:

>    - flattening nested lists
>    - truncating floats
>    - segmenting lists
>    - copying files/directories
>    - hiding/unhiding files
>    - creating symbolic links

**pdf**

> Leverages the Python **PDF** Toolkit, [ReportLab][pdf], to create generic PDFs and PDF checklists.

**reg**

> Functions and classes that Win7ools uses to work with the Windows <b>reg</b>istry.  It contains several functions
  for accessing and working with the registry.  It also contains useful registry locations and 
  can retrieve & decode UserAssist data.

**sec**

> Handles <b>sec</b>urity-related tasks.  Among other things, it contains several hashing functions
  and provides access to Windows [DPAPI][dpapi] functions such as `CryptProtectData()` and `CryptUnprotectData()`.

**web**

> Functions and classes for working with the **web**.

packages
--------
**system**

> The **system** package provides access to the System class.  The System class leverages the modules listed above
  to accomplish a variety of tasks including, but not limited to:

>    - setting the logon screen image
>    - getting/setting `TRIM`
>    - changing drive icons
>    - getting process information
>    - working with the Windows clipboard
>    - finding duplicate files
>    - creating Windows message boxes

------------------------------------------------------------------------------------------------------------------

Examples
--------

**Use win7ools.pdf to create a shopping list:**
```python
from wintools.pdf import Checklist
cl = Checklist()
cl.set_title('My Shopping List').set_pretext('Bring a calculator')
items = ['apples', 'pears', 'broccoli', 'bread', 'chicken', 'soda']
cl.add(items)
cl.save()
```
*win7ools.pdf can check, uncheck, and highlight items.  One and two-column format is supported.*

**Use win7ools.ipl to print names of installed software:**
```python
from wintools.ipl import IPL
ipl = IPL()  
ipl.print_names()
```

**Use wint7ools.reg to print names of software run on host computer with the last run date and count:**
```python
from wintools.reg import get_user_assist()
user_assist = get_user_assist()
for item in user_assist:
    print(item['value'], item['lastrun'], item['count'])
```

**Encrypt and decrypt data using the Windows Data Protection API:**
```python    
from win7ools.sec import crypt_protect_data, crypt_unprotect_data
ctext = crypt_protect_data('Hello world')
ptext = crypt_unprotect_data(ctext)
```

*Consult* `help(win7ools.<module>)` *for more information and examples*

------------------------------------------------------------------------------------------------------------------

Version 2 Improvements
----------------------

- Add entry points
 - [ ] `log()`
 - [ ] `get_external_ip()`
 - [ ] `get_local_ip`
 - [ ] `get_trim()`
 - [ ] `set_trim()`
- Add functions to context menu
 - [ ] `hide()` (file/directory)
 - [ ] `unhide()` (file/directory)
 - [ ] various hash functions
 - [ ] `seet_logon_screen()`
- [ ] Improve fuzzy string matching functionality across all modules (start with `lib.get_closest_match()`)
- [ ] Add initial forensics capabilities (`forWINsics`)
- [ ] Generalize data conversion and persistance
- [ ] Implement "stigmergic persistence"
- [ ] Add web crawler class to win7ools.web
- [ ] Make `reg.get_values(...)` get key values recursively
- [ ] Create Windows event handler (evtx log reader/writer)

[leet]: https://en.wikipedia.org/wiki/Leet
[pdf]: http://www.reportlab.com/opensource/
[dpapi]: http://msdn.microsoft.com/en-us/library/ms995355.aspx
