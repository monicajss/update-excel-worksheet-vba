# update-excel-worksheet-vba
This macro is responsible for searching some data from worksheet A into worksheet B and update worksheet A with the value from worksheet B

## How to use?

This macro was created to separate data from a cell in a certain sheet (sheet A) and search for each separate data in another sheet (sheet B). After finding it, she selects the column cell next to SheetB and replaces it in SheetA. Therefore, you need two different spreadsheets. You can substitute the correct name in parts of the code:

```vba
ActiveWorkbook.Sheets("SheetA").Activate
ActiveWorkbook.Sheets("SheetA").Columns("A:A").Select
'[...]
With Sheets("SheetB").Range("A:A")
'[...]
SplitCurrentCell(x) = ActiveWorkbook.Sheets("SheetB").Range(FirstOccurrence).Value
'[...]
ActiveWorkbook.Sheets("SheetA").Range("A" & i).Value = TextJoin
```

And replace a column you want to copy the value to replace in sheetA:

```vba
 FirstOccurrence = "B" & SplitFirstOccurrence(2) 'B20
 ```
