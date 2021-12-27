![](http://ForTheBadge.com/images/badges/made-with-python.svg)
![](https://forthebadge.com/images/badges/built-by-developers.svg)</br>
[![Prettier](https://img.shields.io/badge/Code%20Style-Prettier-red.svg)](https://github.com/prettier/prettier)
![Size](https://img.shields.io/github/repo-size/Iamtripathisatyam/Split_Excel_File?color=red&label=Repo%20Size%20)
![](https://img.shields.io/tokei/lines/github/Iamtripathisatyam/Split_Excel_File?color=red&label=Lines%20of%20Code)</br>
![sds](https://profile-counter.glitch.me/{Split_Excel_File}/count.svg)

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

## Output:
### Before: 
<p align="center">
<img width="70%"src="https://user-images.githubusercontent.com/69134468/127760576-756e2a48-4b02-47ed-bf68-e9362483833e.jpg"/></a>
</p>

### After:
<p align="center">
<img width="70%"src="https://user-images.githubusercontent.com/69134468/127760550-c111305c-f820-4f92-bc48-04ba6b984570.jpg"/></a>
</p>
