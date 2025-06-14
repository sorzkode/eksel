import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setuptools.setup(
    name='eksel',
    version='2.0.0',
    description='Excel worksheet splitter with modern tkinter GUI and error copying capability.',
    url='https://github.com/sorzkode/eksel',
    author='Mister Riley',
    packages=setuptools.find_packages(),
    py_modules=['eksel'],
    install_requires=[
        'xlwings>=0.27.15',
        'Pillow>=9.0.0',
        'pyperclip>=1.8.2',
    ],
    python_requires='>=3.7',
    entry_points={
        'console_scripts': [
            'eksel=eksel:main',
        ],
    },
    package_data={
        '': ['assets/*.png', 'assets/*.ico'],
    },
    include_package_data=True,
    long_description=long_description,
    long_description_content_type="text/markdown",
    classifiers=[
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Intended Audience :: End Users/Desktop',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
        'Development Status :: 5 - Production/Stable',
        'Environment :: Win32 (MS Windows)',
        'Environment :: MacOS X',
        'Environment :: X11 Applications',
    ],
    keywords='excel, spreadsheet, worksheet, splitter, xlwings, automation',
    license='MIT',
)