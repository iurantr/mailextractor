# mailextractor
Extracts all the email addresses from the .doc , .docx and .pdf files in the folder


Installation


pip packages:
pip install pdfminer.six
pip install docx2txt
pip install chardet


Install <antiword> for Windows:
1) Download at:
http://www.winfield.demon.nl/#Windows
2) Extract to:
C:\antiword
3) Add C:\antiword folder to the PATH variables

Install antiword for Ubuntu:
sudo apt-get install antiword




Creation of the executable:


I) Creation of executable py pyinstaller:

1) pip install pyinstaller
2) In command shell type:
pyinstaller --one-file script_name.py

or:

python -m pyinstaller --one-file script_name.py

The script_name.exe can be found in 'dist' folder.

The executable will work only on the current platform (Windows, Linux, ...)
And, probably, only on the current version of Windows: Windows 7, Windows 10, x64 or x32 bits....

II) Making the executable file smaller:

1) Download upx.exe and put into the folder with the script_name.py
pyinstaller will use it zip all the libraries inside of created script_name.exe

This method reduced the size of .exe in the anaconda virtual environement
from 200 Mb to 140Mb

ATTENTION: Known error on Windows10 w64. Don't use it for win_x64 binaries.
upx.exe breacks the vcruntine140.dll library.
If you get the error that VCRUNTIME140.dll is not q proper dll or contains errors, try to disable it with
--noupx
or, just remove the executable from the folder


2)* Launch pyinstaller in virtual environement with minimal set of packages:
* Most efficient method

-- pip install virtualenv
-- python -m virtualenv envoronement_name # to create a forder with new virtual environement
(Note: don't run this command in conda or other heavy virtual env, 
or lots of heavy default packages will be added automatically)

-- cd envoronement_name\Scripts
-- activate
-- Install all the python packages needed by your script
-- install the pyinstaller
-- create the exe like in I)

3) Exclude even more modules:
https://stackoverflow.com/questions/4890159/python-excluding-modules-pyinstaller/17595149#17595149

III) Adding missing libraries to created script_name.exe:
Lots of advices could be found at:
https://stackoverflow.com/questions/20602727/pyinstaller-generate-exe-file-folder-in-onefile-mode/20677118#20677118
https://pythonhosted.org/PyInstaller/advanced-topics.html#the-toc-and-tree-classes
https://pyinstaller.readthedocs.io/en/v3.3.1/spec-files.html

In short:
1) run pytinstaller one
2) edit created script_name.spec
3) add missing libraries/files to the EXE there. 
To add whole folder with Tree command some programs/toos:

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          Tree('..\\python\\images', prefix='images\\'),
....

'..\\python\\images' - replace with the location of the folder on your local hard drive:
prefix='images\\' - that's the name the folder will have during the execution of the script_name.exe
(place where your script_name.py should be searching for it)
