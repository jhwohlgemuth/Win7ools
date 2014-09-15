# -*- coding: utf-8 -*-
"""pdf.py -- Methods to streamline the use of ReportLab.py"""

import os
import re
from reportlab.platypus import Table
from reportlab.platypus import TableStyle
from reportlab.lib.pagesizes import letter, portrait
from reportlab.platypus import Paragraph 
from reportlab.platypus import Spacer
from reportlab.platypus import BaseDocTemplate
from reportlab.platypus import Frame, FrameBreak
from reportlab.platypus import PageTemplate, NextPageTemplate
from reportlab.lib.styles import getSampleStyleSheet 
from reportlab.lib.colors import black, red, green, blue, gray, purple
from win7ools.lib import flatten
from win7ools.lib import timestamp

FORBIDDEN = re.compile('[/?<>\\:*|"]')

def pdf_color(name):
    '''
    >>> pdf_color('black')
    Color(0,0,0,1)
    '''
    colors = {'black': black,\
              'red': red,\
              'green': green,\
              'blue': blue,\
              'gray': gray,\
              'purple': purple}
    return colors[name]
    
def todo_style(color=pdf_color('black'), align='MIDDLE', size=10):
    style = []
    style.append(('GRID', (0, 0), (0, 0), 1, color))
    style.append(('TEXTCOLOR', (0, 0), (0, 0), color))
    style.append(('TEXTCOLOR', (1, 0), (-1, 0), color))
    style.append(('SPAN', (1, 0), (-1, 0)))
    style.append(('VALIGN', (0, 0), (1, 0), 'MIDDLE'))
    style.append(('SIZE', (0, 0), (1, 0), size))    
    return style
    
class PDF(object):
    '''
    Basic PDF object for formatting data as a PDF.
    Base class for Checklist class.'''
    class Item(object):
        def __init__(self, _type='', value='', order='', data=''):        
            self._type = _type
            self.value = value
            self.order = order
            self.data = data     
    
    def __init__(self, title='', pretext='', two_column=False):
        self.story = ['', '']
        self.pagesize = portrait(letter)
        self.styles = getSampleStyleSheet()
        if title: 
            self.title = PDF.Item('TITLE', title)
        else: 
            self.title = PDF.Item('TITLE', timestamp())   
        self.pretext = PDF.Item('PRETEXT', pretext)
        self.set_title(self.title.value)
        self.set_pretext(self.pretext.value)
        self.two_column = two_column
        self.savedir = '%s\\Documents'%os.path.expanduser('~') 
    
    def __repr__(self):
        return self.title.value
    
    def __str__(self):
        return self.title.value      
    
    def limit(self):
        return 33 if self.two_column else 75
    
    def set_title(self, title):
        self.title.order = 1
        self.title.value = title
        self.title.data = Paragraph(self.title.value, self.styles['Title'])
        self.story[0] = self.title
        return self
        
    def set_pretext(self, pretext, align=1):
        '''
        Alignment:
            0 ==> Left
            1 ==> Center
            2 ==> Right
            3 ==> Justified
        '''
        self.pretext.order = 2
        self.pretext.value = pretext
        pretext_data = Paragraph(self.pretext.value, self.styles['Normal'])
        pretext_data.style.alignment = align
        self.pretext.data = pretext_data
        self.story[1] = self.pretext
        return self

    def add_section(self, text, style='Heading3'):
        section = PDF.Item(_type='SECTION',\
                           value=text,\
                           order=len(self.story) + 1,\
                           data=Paragraph(text, self.styles[style]))
        self.story.append(section)
        return self
    
    def add_paragraph(self, text, align=4):
        paragraph = PDF.Item(_type='PARAGRAPH',\
                             value=text,\
                             order=len(self.story) + 1,\
                             data=Paragraph(text, self.styles['BodyText']))
        self.story.append(paragraph)
        self.add_spacer(10)
        return self
        
    def add_spacer(self, H):
        spacer = PDF.Item(_type='SPACER',\
                          order=len(self.story) + 1,\
                          data=Spacer(0, H))
        self.story.append(spacer)
        return self
       
    def save(self, savedir='', show=True, auto=False, boundary=0):
        '''
        Saves to savedir as PDF and opens if show=True (default).
        '''
        story = [i.data for i in self.story]
        if savedir:
            self.savedir = savedir
        savename = re.sub(FORBIDDEN, '', self.title.value)
        savepath = '%s\\%s.pdf'%(self.savedir, savename)
        
        doc = BaseDocTemplate(savepath,\
                              pagesize=self.pagesize,\
                              showBoundary=boundary) 
                             
        pre = Paragraph(self.pretext.value, self.styles['Normal'])
        w, h = pre.wrap(200, 70)
        preambleHeight = max(h, 70)  
        
        if auto:
            items = [i.value for i in self.story]
            self.two_column = True
            for item in items:
                if len(str(item)) < self.limit():
                    pass
                else:
                    self.two_column = False
                    break
        
        if self.two_column:
            #First Page
            frame_Top = Frame(doc.leftMargin,\
                             doc.bottomMargin + doc.height - preambleHeight,\
                             doc.width,\
                             preambleHeight,\
                             id='top')
            frame_L = Frame(doc.leftMargin,\
                              doc.bottomMargin,\
                              doc.width/2-6,\
                              doc.height-preambleHeight,\
                              id='col1')
            frame_R = Frame(doc.leftMargin+doc.width/2+6,\
                               doc.bottomMargin,\
                               doc.width/2-6,\
                               doc.height-preambleHeight,\
                               id='col2')
            #Later Pages
            column_L = Frame(doc.leftMargin,\
                                doc.bottomMargin,\
                                doc.width/2-6,\
                                doc.height,\
                                id='colLeft')
            column_R = Frame(doc.leftMargin+doc.width/2+6,\
                                 doc.bottomMargin,\
                                 doc.width/2-6,\
                                 doc.height,\
                                 id='colRight')   
            firstTemplate = PageTemplate(id='FirstPage',\
                                         frames=[frame_Top, frame_L,frame_R])
            laterTemplate = PageTemplate(id='LaterPages',\
                                         frames=[column_L, column_R])
        else:
            #First Page
            frame_Top = Frame(doc.leftMargin,\
                             doc.bottomMargin + doc.height - preambleHeight,\
                             doc.width,\
                             preambleHeight,\
                             id='top')
            frame_Bottom = Frame(doc.leftMargin,\
                              doc.bottomMargin,\
                              doc.width*2,\
                              doc.height-preambleHeight,\
                              id='bottom')
            #Later Pages
            column = Frame(doc.leftMargin,\
                                doc.bottomMargin,\
                                doc.width,\
                                doc.height,\
                                id='column')
            firstTemplate = PageTemplate(id='FirstPage',\
                                         frames=[frame_Top, frame_Bottom])
            laterTemplate = PageTemplate(id='LaterPages',\
                                         frames=[column])

        doc.addPageTemplates([firstTemplate, laterTemplate])                                 
        doc.build(story)
        if show: os.startfile(savepath)
        return savepath

class Checklist(PDF):
    '''
    Make a checklist using ReportLab.py

    ***Intended for use with letter size paper***    
    '''  
    class Todo(PDF.Item):
        def __init__(self,\
                     _type='TODO',\
                     value='',\
                     order='',\
                     data='',\
                     complete=False,
                     color='black'):
            super(Checklist.Todo, self).__init__(_type, value, order, data)
            self.complete = complete
            self.color = color
            self.mark = 'X'
            
        def render(self, two_column):
            '''
            '''
            bodytext = getSampleStyleSheet()['BodyText']
            row_color = pdf_color(self.color)
            style = todo_style(color=row_color)
            row = [''] * 11 if two_column else [''] * 23
            limit = 33 if two_column else 75
            if len(self.value) > limit:
                self.value = self.value[:limit] + '...'
            row[0] = self.mark if self.complete else ''
            ptext = '<font color=%s>%s</font>'%(row_color, self.value)
            row[1] = Paragraph(ptext, bodytext)
            todo = Table([row], colWidths=20, rowHeights=20)
            todo.setStyle(TableStyle(style))
            todo.hAlign = 'LEFT'
            self.data = todo
    
    def __init__(self, items=[], title='', pretext=''):
        super(Checklist, self).__init__(title, pretext)
        FRAMEBREAK = PDF.Item(_type='FRAMEBREAK',\
                              order=len(self.story) + 1,\
                              data=FrameBreak())    
        self.story.append(FRAMEBREAK)
        NEXTPAGETEMPLATE = PDF.Item(_type='TEMPLATE',\
                                    order=len(self.story) + 1,\
                                    data=NextPageTemplate('LaterPages'))
        self.story.append(NEXTPAGETEMPLATE)
        if items:
            items = flatten(items)
            self.add(items)
        
    def __len__(self):
        '''
        >>> cl = Checklist([1,2,3,4,5])
        >>> len(cl)
        5
        '''
        return len(self.items()) 

    def __add__(self, other):
        '''
        >>> c1 = Checklist([1,2,3])
        >>> c1.items()
        ['1', '2', '3']
        
        >>> c2 = Checklist([4, 5, 6])
        >>> c2.items()
        ['4', '5', '6']
        
        >>> c3 = c1 + c2
        >>> c3.items()
        ['1', '2', '3', '4', '5', '6']
        '''
        c = Checklist(title=self.title.value, pretext=self.pretext.value)
        for item in self.items():
            c.add(item)
        for item in other.items():
            c.add(item)
        return c

    def __contains__(self, item):
        '''
        >>> 'a' in Checklist(['a','b','c'])
        True
        '''
        return item in self.items()
    
    def clear(self):
        '''
        >>> cl = Checklist([1,2,3,4,5])
        >>> len(cl.clear())
        0
        '''
        return Checklist([], self.title.value, self.pretext.value)
    
    def items(self):
        '''
        >>> cl = Checklist(['a','b','c'])
        >>> cl.items()
        ['a', 'b', 'c']
        '''
        return [i.value for i in self.story if i._type is 'TODO']
        
    def completed(self):
        '''
        >>> cl = Checklist(['a','b','c'])
        >>> cl.check(['a','b']).completed()
        ['a', 'b']
        '''
        todos = [i for i in self.story if i._type is 'TODO']
        return [i.value for i in todos if i.complete]
    
    def add(self, *items, **attr):
        '''
        >>> c1 = Checklist()
        >>> c1.add('a').items()
        ['a']
        >>> c1.add(['b','c',1,2,3]).items()
        ['a', 'b', 'c', '1', '2', '3']
        
        >>> c2 = Checklist()
        >>> c2.add(1,2,3,'a','b','c').items()
        ['1', '2', '3', 'a', 'b', 'c']
        
        >>> c3 = Checklist()
        >>> c3.add('one','two','three').items()
        ['one', 'two', 'three']
        
        >>> c4 = Checklist()
        >>> c4.add('one long task').items()
        ['one long task']
        
        >>> c5 = Checklist('one long task')
        >>> c5.items()
        ['one long task']
        '''
        items = [str(i) for i in flatten(items)]
        row_color = 'black'
        complete = attr['complete'] if 'complete' in attr.keys() else False
        if 'color' in attr.keys() and not complete:
            try:
                row_color = attr['color']
            except:
                row_color = 'black'
        if complete:
            row_color = 'gray'
        for item in items:      
            
            TODO = Checklist.Todo(value=item,\
                                  order=len(self.story) + 1,\
                                  color=row_color)
            TODO.render(self.two_column)
            self.story.append(TODO)
            self.add_spacer(5)
        return self        

    def remove(self, *items, **attr):
        '''
        >>> cl = Checklist(['a','b','c'])
        >>> cl.remove('a','b').items()
        ['c']
        '''
        items = [str(i) for i in flatten(items)]
        todos = [i for i in self.story if i.value in items]
        for todo in todos:
            self.story.remove(todo)
        return self
        
    
    def check(self, *items, **keywords):
        '''
        >>> cl = Checklist(['a','b','c'])
        >>> cl.check('a','b','c').completed()
        ['a', 'b', 'c']
        '''
        items = [str(i) for i in flatten(items)]
        todo_index = [(i.order - 1) for i in self.story if i._type is 'TODO']
        values = [{self.story[i].value: i} for i in todo_index]
        value_dict = {}
        for value in values:
            value_dict.update(value)
        for item in items:
            index = value_dict[item]
            todo = self.story[index]
            todo.complete = True
            todo.color = 'gray'
            todo.render(self.two_column)
            self.story[index] = todo            
        return self

    def uncheck(self, *items, **keywords):
        '''
        >>> cl = Checklist(['a','b','c'])
        >>> cl.check('a','b','c').completed()
        ['a', 'b', 'c']
        >>> cl.uncheck('a','c').completed()
        ['b']
        '''
        items = [str(i) for i in flatten(items)]
        todo_index = [(i.order - 1) for i in self.story if i._type is 'TODO']
        values = [{self.story[i].value: i} for i in todo_index]
        value_dict = {}
        for value in values:
            value_dict.update(value)
        for item in items:
            index = value_dict[item]
            todo = self.story[index]
            todo.complete = False
            todo.color = 'black'
            todo.render(self.two_column)
            self.story[index] = todo            
        return self

    def highlight(self, items, color='red'):
        items = [str(i) for i in flatten(items)]
        todo_index = [(i.order - 1) for i in self.story if i._type is 'TODO']
        values = [{self.story[i].value: i} for i in todo_index]
        value_dict = {}
        for value in values:
            value_dict.update(value)
        for item in items:
            index = value_dict[item]
            todo = self.story[index]
            todo.color = color
            todo.render(self.two_column)
            self.story[index] = todo            
        return self
        
if __name__ == '__main__':
    import doctest
    doctest.testmod(verbose=True)