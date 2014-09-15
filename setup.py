from setuptools import setup, find_packages  # Always prefer setuptools over distutils
from codecs import open  # To use a consistent encoding
from os import path

here = path.abspath(path.dirname(__file__))
 
with open(path.join(here, 'README.rst'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='win7ools',

    version='1.0.0',

    description='Python project that provides programmatic access to the Windows OS',
    long_description=long_description,

    url='https://github.com/nexuslevel/Win7ools',

    author='Jason Wohlgemuth',
    author_email='jason.h.wohlgemuth@gmail.com',

    license='MIT',

    # See https://pypi.python.org/pypi?%3Aaction=list_classifiers
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Environment :: Win32 (MS Windows)',
        'Operating System :: Microsoft :: Windows :: Windows 7',
        'Programming Language :: Python :: Implementation :: CPython',

        'Topic :: Scientific/Engineering :: Artificial Intelligence',
        'Topic :: Security',
        'Topic :: System',
        'Topic :: Utilities',

        'Natural Language :: English',
        'Intended Audience :: Developers',
        'Intended Audience :: Information Technology',
        'Intended Audience :: Science/Research',
        'Intended Audience :: System Administrators',

        'License :: OSI Approved :: Apache Software License',
 
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
    ],

    keywords=['programmatic', 'access', 'automation', 'customization', 'productivity'],

    # You can just specify the packages manually here if your project is
    # simple. Or you can use find_packages().
    packages=['win7ools', 'win7ools.system'],

    # List run-time dependencies here.  These will be installed by pip when your
    # project is installed. For an analysis of "install_requires" vs pip's
    # requirements files see:
    # https://packaging.python.org/en/latest/technical.html#install-requires-vs-requirements-files
    install_requires=[],

    # To provide executable scripts, use entry points in preference to the
    # "scripts" keyword. Entry points provide cross-platform support and allow
    # pip to create the appropriate form of executable for the target platform.
    entry_points={
        'console_scripts': [
            'trim=win7ools.system:System.get_trim',
        ],
    },
)
