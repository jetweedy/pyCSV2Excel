# pyCSV2Excel

## Installation

1. Install Python 3.6 or higher
2. Then install necessary modules:
 ```
 python3 -m pip install xlsxwriter openpyxl pandas
 ```

## Command Line Usage

### Convert multiple CSV files to a single XLSX with tabs:

```
 python3 csv2xlsx.py targetfile.xlsx sourcefile1 sourcefile2 --tabnames"Tab 1|Tab B"
```

### Password-protect an XLSX file:

```
cat input.xlsx | secure-spreadsheet --password secret --input-format xlsx > output.xlsx
```

Resource: https://github.com/ankane/secure-spreadsheet
