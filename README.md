# pyCSV2Excel

## Command Line Usage

### Convert multiple CSV files to a single XLSX with tabs:

```
 python3 csv2xlsx.py targetfile.xlsx sourcefile1 sourcefile2 etc
```

### Password-protect an XLSX file:

```
cat input.xlsx | secure-spreadsheet --password secret --input-format xlsx > output.xlsx
```

Resource: https://github.com/ankane/secure-spreadsheet