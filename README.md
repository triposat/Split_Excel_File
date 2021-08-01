![](http://ForTheBadge.com/images/badges/made-with-python.svg)
![](https://forthebadge.com/images/badges/built-by-developers.svg)</br>
[![Prettier](https://img.shields.io/badge/Code%20Style-Prettier-red.svg)](https://github.com/prettier/prettier)
![Size](https://img.shields.io/github/repo-size/Iamtripathisatyam/Split_Excel_File?color=red&label=Repo%20Size%20)
![](https://img.shields.io/tokei/lines/github/Iamtripathisatyam/Split_Excel_File?color=red&label=Lines%20of%20Code)</br>
![sds](https://profile-counter.glitch.me/{Split_Excel_File}/count.svg)

<p align="center">
<a href="https://github.com/Iamtripathisatyam/Split_Excel_File/blob/main/split_excel_files.py"><img width="15%"src="https://cdn.icon-icons.com/icons2/195/PNG/256/Excel_2013_23480.png" /></a>
</p>

We all know how crucial Excel sheets are today for holding large data packets. We will divide the sheets using the ***xlwings*** package, which is defined in Python. We will first copy all of the sheets and then try to convert them into individual Excel files.

Let's have a look at the process step-by-step:

- Install the module given below

    ```pytho
    pip install xlwings
    ```
- Import the following modules

    ```python
    import xlwings as xw
    import os 
  ```
- Create a new excel instance with xlwings.App, then set visibility to 'False' if you don't want the excel window to appear. Then, using the books.open function, open the worksheet.
- We'll now execute a loop to traverse through each sheet.
- Using the api.copy method, copy each sheet to a new workbook.
- After that, before saving the new worksheet, be sure you activate it.
- We will now use the sheet names as file names.
- At last, we'll shut our workbook.
- Finally: when you exit the excel sheet, you will notice that all of the sheets have been changed to new excel files.
