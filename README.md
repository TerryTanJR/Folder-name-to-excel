## Folder names and hyperlinks to excel

A program that extracts folder names and saves both their names and hyperlinks to their actural locations into an Excel file.

Note that :
* Please place all the folders inside another folder so that **folder_path** can find the location of all the folders.

## Requirements
Install necessary dependency
```bash
    pip install openpyxl
```
## Part of codes that you need to notice
There are two places that you need to change.
1. replace the **folder_path** to tour own
2. In this example, all the folders are placed in a folder called "Example"
```
folder_path = r"C:\your own path\Example"
```
2. set where you want to store your excel file
```
excel_file_path = r'C:\your own path\folders_names.xlsx'
```
Finally, if you want to capture the hyperlink of a specific file within a folder, instead of just the folder's hyperlink, you can try to modify the following code.

Note that :
* if you want to capture the hyperlink of a specific file within a folder, all the folders **must have** the same target files

```
folder_full_path = folder_full_path + "replace this quote with the full path to the target file"
```

## Dependency
openpyxl

