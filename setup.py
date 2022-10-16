import setuptools

setuptools.setup(
    name='eksel',
    version='1.0.0',
    description='EKSEL SPLITTER.',
    url='https://github.com/sorzkode/',
    author='sorzkode',
    author_email='<sorzkode@proton.me>',
    packages=setuptools.find_packages(),
    install_requires=['xlwings', 'PySimpleGUI', 'tkinter'],
    long_description='Quickly copy and save Excel worksheets as seperate workbooks.',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: MIT',
        'Operating System :: OS Independent',
        ],
)