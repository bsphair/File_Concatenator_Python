# File Concatenator

Program will transfer data from one Excel file to another Excel file.

## Getting Started

This program uses the following packages:

- pandas
- xlrd
- xlsxwriter
- openpyxl

## Using

To run the program, the general format is as follows:

```
$ python3 index.py <directory location> <new data file> <all data file>
```
- directory_location is the directory of the files
- new_data_file is the file that contains the data to be transfered 
- all_data_file is the file in which the data is being tranfered to


## Acknowledgments/Resources

- https://www.datacamp.com/community/tutorials/python-excel-tutorial
- https://xlsxwriter.readthedocs.io/working_with_pandas.html (remove the header from data)
- https://stackoverflow.com/questions/20219254/how-to-write-to-an-existing-excel-file-without-overwriting-data-using-pandas
