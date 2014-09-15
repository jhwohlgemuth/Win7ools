# -*- coding: utf-8 -*-
'''wintools -- Collection of code snippets for the Windows OS'''

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

__author__ = 'Jason Wohlgemuth'
__version__ = '1.0'
__versioninfo__ = (1, 0, 0)

import doctest
import unittest
from distutils.log import warn
from win7ools.system import System
from win7ools.ipl import IPL
from win7ools.reg import RegistryKeys as keys

python_paths = ['C:\\Python27',\
                'C:\\Python27\\Lib\site-packages',\
                'C:\\Python27\\Scripts',\
                'C:\\Python27\\Tools\\Scripts']


def test():
    suite = unittest.TestSuite()
    runner = unittest.TextTestRunner(verbosity=2)
    modules = ['ipl', 'lib', 'pdf', 'sec']
    for module in modules:
        warn('importing ' + module + ' ...')
        _temp = __import__('win7ools', globals(), locals(), [module], -1)
        mod_test_suite = doctest.DocTestSuite(eval(module))
        suite.addTest(mod_test_suite)        
    runner.run(suite)

if __name__ == '__main__':
    print('Under construction...')