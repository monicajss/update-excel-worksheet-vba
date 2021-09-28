Sub UpdateSheet()
'
' This macro is responsible for searching some data from worksheet A into worksheet B and
' update worksheet A with the value from worksheet B
'
    
'---------------- Variables ----------------

    Dim DataRange As Range
    Dim CellTotal As Integer
    Dim i, x As Integer
    Dim CurrentCell As String
    Dim SplitCurrentCell() As String
    Dim FirstOccurrence As String
    Dim SplitFirstOccurrence() As String
    Dim TextJoin As String

'-----------------------------------------
    
    ' Select the column (in this case the column B)  and assign it to the  variable "DataRange"
    ActiveWorkbook.Sheets("SheetA").Activate
    ActiveWorkbook.Sheets("SheetA").Columns("A:A").Select
    Set DataRange = Selection
    
    ' Count how many non-zero cells we have in the variable "DataRange"
    CellTotal = WorksheetFunction.CountA(DataRange)
    
    ' FOR  responsible for passing in all cells within DataRange range
    For i = 2 To CellTotal Step 1
    
        ' ActualCell receives current cell value
        CurrentCell = DataRange.Cells(i, 1)
        
        ' SplitCurrentCell separates CurrentCell values and assigns to an array
        SplitCurrentCell = Split(CurrentCell, ";")
        
        ' FOR  esponsible for passing in all values within the SplitCurrentCell array
        For x = LBound(SplitCurrentCell, 1) To UBound(SplitCurrentCell, 1)
        
            ' Search for the value of SplitCurrentCell(x) in SheetB
            With Sheets("SheetB").Range("A:A")
                Set Interval = .Find(What:=SplitCurrentCell(x), _
                                  After:=.Cells(1), _
                                  LookIn:=xlValues, _
                                  LookAt:=xlWhole, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False)
                
                If Not Interval Is Nothing Then
                    FirstOccurrence = Interval.Address '$A$20
                     SplitFirstOccurrence = Split(FirstOccurrence, "$")
                    FirstOccurrence = "B" & SplitFirstOccurrence(2) 'B20
                    ' Assigns column B value to SplitCurrentCell(x)
                    SplitCurrentCell(x) = ActiveWorkbook.Sheets("SheetB").Range(FirstOccurrence).Value
                End If
            End With
        Next x
        ' TextJoin transforms array to string
        TextJoin = Join(SplitCurrentCell, ";")
        ' Assign the TextJoin to the current cell in column A present in Sheet A
        ActiveWorkbook.Sheets("SheetA").Range("A" & i).Value = TextJoin
    Next i
    
    ' Save a copy from worksheet
    GetBookName = Split(ActiveWorkbook.Name, ".")
    NewBookName = ThisWorkbook.Path & "/" & GetBookName(0) & Format(Now(), "yyyymmdd") & ".xlsx"
    ActiveWorkbook.SaveCopyAs NewBookName
    
End Sub

