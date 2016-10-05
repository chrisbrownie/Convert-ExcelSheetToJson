# Convert-ExcelSheetToJson.ps1

## SYNOPSIS
Converts an Excel sheet from a workbook to JSON

## SYNTAX

```
Convert-ExcelSheetToJson.ps1 [-InputFile] <Object> [[-OutputFileName] <String>] [[-SheetName] <String>]
```

## DESCRIPTION
To allow for parsing of Excel Workbooks suitably in PowerShell, this script converts a sheet from a spreadsheet into a JSON file of the same structure as the sheet.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Convert-ExcelSheetToJson -InputFile MyExcelWorkbook.xlsx
```

### -------------------------- EXAMPLE 2 --------------------------
```
Get-Item MyExcelWorkbook.xlsx | Convert-ExcelSheetToJson -OutputFileName MyConvertedFile.json -SheetName Sheet2
```

## PARAMETERS

### -InputFile
The Excel Workbook to be converted.
Can be FileInfo or a String.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: 

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -OutputFileName
The path to a JSON file to be created.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetName
The name of the sheet from the Excel Workbook to convert.
If only one sheet exists, it will convert that one.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

## OUTPUTS

## NOTES
Written by: Chris Brown

Find me on:
* My blog: https://flamingkeys.com/
* Github: https://github.com/chrisbrownie

## RELATED LINKS

[https://flamingkeys.com/convert-excel-sheet-json-powershell](https://flamingkeys.com/convert-excel-sheet-json-powershell)

