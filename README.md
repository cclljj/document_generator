# Document Generator

Given the data file (.xlsx) and the document template (.docx), the program will automatically generate a number of documents in the specified folder by inserting each row of data into the template document.

```
usage: letter_generator.py [-h] -d DATAFILE -t TEMPLATE [-p PREFIX]
```
**DATAFILE** is the data file in .xlsx format. The file must have the header row containing the name of each column, and the first column contains the file name for the output letter of each row data.

**TEMPLATE** is the template document in .docx format. Note that we use '{{FIELD}}' to mark the text that is going to be replaced by the program.

**PREFIX** is the folder of the output files.
