# EZDC
1) download and unzip the files in your desktop

2) install python in your computer
  - https://www.python.org/downloads/
  - while installing, make sure to tick the box to add python to your env path
3) open your terminal, go to the directory you download, install the required package
  - eg in window: 
        i) open 'CMD'
        ii) type 'cd desktop\EZDC'
        iii) type 'pip install -r requirements.txt'
4) type 'python excelmodifier.py'
  
And you will have the updated EZDC excel file.

THINGs to modify:
If you want to update the product list, just add them under the last row of Sheet1.
This script is coded base on the ASIN, name of product sheet(Sheet1), and corresponding columns(now ASIN sits on column4).
So if some changes are needed, please check the comment within the code.
