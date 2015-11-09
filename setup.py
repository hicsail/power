#!/usr/bin/env python
# -*- coding: utf-8 -*-

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

from setuptools.command.test import test as TestCommand
import power
import sys

with open('README.md') as readme_file:
    readme = readme_file.read()


class PyTest(TestCommand):
    def finalize_options(self):
        TestCommand.finalize_options(self)
        self.test_args = ['--strict', '--verbose', '--tb=long', 'tests']
        self.test_suite = True

    def run_tests(self):
        import pytest
        errno = pytest.main(self.test_args)
        sys.exit(errno)

setup(
    name='power',
    version=power.__version__,
    description="Power Grid project description",
    long_description=readme,
    author="Boston University",
    author_email='fjansen@bu.edu',
    url='https://github.com/Hariri-Institute-SAIL/power',
    packages=[
        'power',
    ],
    include_package_data=True,
    zip_safe=False,
    install_requires=[
    ],
    cmdclass={'test': PyTest},
    license="Commercial",
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Intended Audience :: Developers',
        'License :: Commercial',
        'Natural Language :: English',
        "Programming Language :: Python :: 3"
    ],
    test_suite='tests',
    tests_require=['pytest'],
    extras_require={
        'testing': ['pytest'],
    }
)
