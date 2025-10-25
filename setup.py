import setuptools
import re
import functools
import pathlib

def read_file(name: str) -> str:
    return pathlib.Path(name).read_text()


version = re.search(r"__version__ = '([0-9.]*)'", read_file('dodo/__init__.py')).group(1)

requirements = read_file('requirements.txt').strip().split('\n')

setuptools.setup(
    name='dodo-desktop-switcher',
    version=version,
    author='Ram Rachum',
    author_email='ram@rachum.com',
    description='A Windows virtual desktop switcher with keyboard shortcuts',
    long_description=read_file('README.md'),
    long_description_content_type='text/markdown',
    url='https://github.com/cool-RR/dodo',
    project_urls={
        'Bug Tracker': 'https://github.com/cool-RR/dodo/issues',
        'Source Code': 'https://github.com/cool-RR/dodo',
    },
    packages=setuptools.find_packages(exclude=['tests*']),
    install_requires=requirements,
    include_package_data=True,
    python_requires='>=3.12',
    entry_points={
        'console_scripts': [
            'dodo = dodo:cli'
        ],
    },
    classifiers=[
        'Development Status :: 4 - Beta',
        'Environment :: Win32 (MS Windows)',
        'Intended Audience :: End Users/Desktop',
        'Operating System :: Microsoft :: Windows :: Windows 10',
        'Operating System :: Microsoft :: Windows :: Windows 11',
        'Programming Language :: Python :: 3.12',
        'Programming Language :: Python :: 3.13',
        'Programming Language :: Python :: 3.14',
        'Programming Language :: Python :: Implementation :: CPython',
        'Topic :: Desktop Environment',
    ],
)
