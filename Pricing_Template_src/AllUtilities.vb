Option Explicit

'Utilities used in various procedures throughout the document.
'Contained within this module because they serve a general purpose
'And could see re-use in future tasks.
'First Functions, Then Subroutines. Should be Alphabetical.

'FUNCTIONS
'FUNCTIONS
'FUNCTIONS

Function AppendDate(message)
    Dim dateSplit() As String: dateSplit() = Split(CStr(Date), "/", 3)

    For i = 0 To 2
        If (Len(dateSplit(i)) = 1) Then
            message = message & "0"
        ElseIf (Len(dateSplit(i)) = 4) Then
            dateSplit(i) = Right(dateSplit(i), 2)
        End If
        
        message = message & dateSplit(i)
    Next
    
    AppendDate = message
End Function

Function ArrayToCollection(A As Variant) As Collection
    'https://stackoverflow.com/questions/12258732/easy-way-to-convert-an-array-to-a-collection-in-vba
    
    Dim c As New Collection
    
    For Each Item In A
        c.Add Item
    Next Item
    
    Set VariantToCollection = c
End Function

Function CollectColumnNames(sht As Worksheet) As Collection
    Dim cell As Range, columnNames As New Collection
    Dim errorMsg As String, emptyMsg As String, err As Boolean
    
    errorMsg = "Errors For: " & vbCrLf & _
                        "Book: """ & sht.Parent.name & """" & vbCrLf & _
                        "Sheet: """ & sht.name & """" & vbCrLf & vbCrLf
                        
    emptyMsg = "Error: Column Name Missing (Required Value)" & vbCrLf & _
                            "Unnamed Column Numbers: "

    For Each cell In RetrieveColumns(sht)
        If cell.Value <> Empty Then
            columnNames.Add Item:=cell.Value
        Else
            If Not err Then
                err = True
            Else
                emptyMsg = emptyMsg & ", "
            End If
            
            emptyMsg = emptyMsg & cell.Column
        End If
    Next cell
    
    If err Then
        MsgBox errorMsg & emptyMsg
    End If
    
    Set CollectColumnNames = columnNames
End Function

Function CollectColumnIds(namesCollection As Collection, _
                                            namedRange As Range) As Collection
    'Look in namedRange for cells which match values in namesCollection.
    'For each match, assign new cIds key = to matching value from
    'namesCollection, and value = to column # of matching cell from
    'namedRange. Ignore all cIds without matching cNames, cNames without
    'matching cIds will pop msgbox, because all columns should be found.
    
    Dim colName As Variant, colFindResults As Range, cIds As New Collection
    Dim errorMsg As String, empError As Boolean, emptyCount As Long
    Dim duplicateCols As String, dupError As Boolean
    Dim missingCols As String, misError As Boolean
    
    'Columns can only be across one row.
    If namedRange.Rows.Count = 1 Then
        errorMsg = "Errors For: " & vbCrLf & _
                            "Book: """ & namedRange.Parent.Parent.name & """" _
                            & vbCrLf & "Sheet: """ & namedRange.Parent.name _
                            & """" & vbCrLf & vbCrLf
                            
        missingCols = "Error: Expected Column Not Found" & vbCrLf & _
                                "Missing Columns:" & vbCrLf
                                
        duplicateCols = "Error: Column Name Not Unique (Must Be Unique, Only" _
                                    & " 1st Transfers)" & vbCrLf & _
                                    "Duplicate Columns:" & vbCrLf
                                    
        dupError = False
        misError = False
        empError = False
        
        emptyCount = 0
        
        'For each column name in names collection:
        For Each colName In namesCollection
            If colName <> Empty Then
                '   Perform find command to search for the column name
                Set colFindResults = _
                            namedRange.Find(What:=colName, _
                                                           LookIn:=xlValues, _
                                                           LookAt:=xlWhole, _
                                                           MatchCase:=False, _
                                                           SearchFormat:=False)
                                                           
                ' If column is found
                If Not colFindResults Is Nothing Then
                    'add the integer representation of found column to cIdsColl
                    'IF NOT KEY ALREADY EXISTS FOR COLLECTION
                    'Debug.Print colFindResults(1, 1).Column & " | " & colName
                    If Not Contains(cIds, colName) Then
                        cIds.Add Item:=colFindResults(1, 1).Column, _
                                        key:=colName
                    Else
                        If dupError = False Then
                            dupError = True
                            
                            If misError = True Then
                                duplicateCols = vbCrLf & duplicateCols
                            End If
                        End If
                        
                        duplicateCols = duplicateCols & """" & colName _
                                                        & """" & vbCrLf
                    End If
                    'Debug.Print colName & ": Column " & cIds(colName)
                ' If not found, error box report. all columns should be found.
                Else
                    If misError = False Then
                        misError = True
                        
                        If dupError = True Then
                             duplicateCols = vbCrLf & duplicateCols
                        End If
                    End If
                    
                    missingCols = missingCols & """" & colName & """" & vbCrLf
                End If
            Else
                empError = True
                emptyCount = emptyCount + 1
            End If
        Next colName
        
        If dupError Or misError Or empError Then
            If misError Then
                errorMsg = errorMsg & missingCols
            End If
            
            If dupError Then
                errorMsg = errorMsg & duplicateCols
            End If
            
            If empError Then
                errorMsg = errorMsg & "Error: # of Untransferable" & _
                                    " Unnamed Columns: " & CStr(emptyCount)
            End If
            
            MsgBox errorMsg
        End If
        'Return collection of whichever cIds are found.
        Set CollectColumnIds = cIds
    Else
        MsgBox "Error: Column Range Must Be Of Height One."
    End If
End Function

Function CollectionToArray(c As Collection) As Variant()
    'https://brettdotnet.wordpress.com/2012/03/30/convert-a-collection-to-an-array-vba/
    
    Dim A() As Variant
    Dim i As Integer
    
    ReDim A(0 To c.Count - 1)
    
    For i = 1 To c.Count
        A(i - 1) = c.Item(i)
    Next
    
    CollectionToVariant = A
End Function


Function Contains(Coll As Collection, key As Variant) As Boolean
    'https://stackoverflow.com/questions/137845/determining-whether-an-object-is-a-member-of-a-collection-in-vba
    
    On Error GoTo err
    
    Contains = True
    
    Dim obj As Variant
    
    obj = IsObject(Coll(key))
    
    Exit Function

err:
    Contains = False
End Function

Function DoubleContains(outerColl As Collection, _
                                            innerColl As Collection, _
                                            key As Variant) As Boolean
                                            
        If Contains(outerColl, key) Then
            If Contains(innerColl, outerColl(key)) Then
                DoubleContains = True
            Else
                DoubleContains = False
                
                Debug.Print "Error: Could not find key """ & key _
                                        & """ in inner collection."
            End If
        Else
            DoubleContains = False
            
            Debug.Print "Error: Could not find key """ & key _
                                    & """ in outer collection."
        End If
End Function

Function GetFolderPath() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False 'may enable in the future, will need rework
        .Title = "Select an import folder:"
        
        If .Show = -1 Then
           GetFolderPath = .SelectedItems(1)
        Else
            GetFolderPath = "ERROR"
        End If
    End With
End Function

Function GetSheetIndex(sheet As String) As Integer
    Dim i As Long
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Worksheets(i).name = sheet Then
            GetSheetIndex = i
            Exit Function
        End If
    Next i
    
    GetSheetIndex = 999
End Function

Function LastOccupiedColNum(sheet As Worksheet) As Long
    'https://gist.github.com/danwagnerco/040402917376969bf362
    
    Dim lng As Long
    
    If Application.WorksheetFunction.CountA(sheet.Cells) <> 0 Then
        With sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByColumns, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).Column
        End With
    Else
        lng = 1
    End If
    
    'Debug.Print "Last Col for " & Sheet.Name & ": " & lng
    LastOccupiedColNum = lng
End Function

Function LastOccupiedColNumInRow(sheet As Worksheet, RowNum As Long) As Long
    'https://gist.github.com/danwagnerco/f4575415d900d9e4b40e699572bd58da#file-combine_certain_sheets-vb
    
    Dim lng As Long
    
    If RowNum > 0 Then
        With sheet
            lng = .Cells(RowNum, .Columns.Count).End(xlToLeft).Column
        End With
    Else
        lng = 0
    End If
    
    LastOccupiedColNumInRow = lng
End Function

Function LastOccupiedRowNum(sheet As Worksheet) As Long
    'https://gist.github.com/danwagnerco/040402917376969bf362
    'https://danwagner.co/how-to-combine-data-from-multiple-sheets-into-a-single-sheet/

    Dim lng As Long
    
    If Application.WorksheetFunction.CountA(sheet.Cells) <> 0 Then
        With sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).Row
        End With
    Else
        lng = 1
    End If
    
    LastOccupiedRowNum = lng
End Function

Function LastOccupiedRowNumInCol(sheet As Worksheet, ColNum As Long) As Long
    'https://danwagner.co/how-to-combine-data-from-certain-sheets-but-not-others-into-a-single-sheet/
    'https://gist.github.com/danwagnerco/f4575415d900d9e4b40e699572bd58da#file-combine_certain_sheets-vb
    Dim lng
    
    If ColNum > 0 Then
        With sheet
            lng = .Cells(.Rows.Count, ColNum).End(xlUp).Row
        End With
    Else
        lng = 0
    End If
    
    LastOccupiedRowNumInCol = lng
End Function

Function RetrieveColumns(sht As Worksheet) As Range
    Set RetrieveColumns = sht.Range(Cells(1, 1).Address, _
                            Cells(1, LastOccupiedColNumInRow(sht, 1)).Address)
End Function

Function RetrieveExcelFilePaths(Optional multi As Boolean = False) As String()
    Dim i As Integer
    Dim fileArray() As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = multi
        
        If multi Then
            .Title = "Select files to import:"
        Else
            .Title = "Select a file to import:"
        End If
        
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls?"
        .Filters.Add "All Files", "*.*"

        If .Show = -1 Then
            ReDim fileArray(.SelectedItems.Count - 1)
            
            If multi Then
                For i = 1 To .SelectedItems.Count
                  fileArray(i - 1) = .SelectedItems(i)
                Next
            Else
                fileArray(0) = .SelectedItems(1)
            End If
            
            RetrieveExcelFilePaths = fileArray
        End If
    End With
End Function

Function RetrieveExcelPathsInFolderPath(folderPath) As Variant()
    'https://stackoverflow.com/questions/29421223/get-list-of-file-names-in-folder-directory-with-excel-vba
    
    Dim fileName As String
    Dim fileCount As Integer: fileCount = 0
    Dim fileArray() As Variant

    fileName = Dir(folderPath & "\*.xls?")
    
    Do While fileName <> ""
        ReDim Preserve fileArray(fileCount)
        
        fileArray(fileCount) = folderPath & "\" & fileName
        fileCount = fileCount + 1
        fileName = Dir()
    Loop

    If fileCount = 0 Then
        RetrieveExcelPathsInFolderPath = Array("ERROR")
    Else
        RetrieveExcelPathsInFolderPath = fileArray
    End If
End Function

Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    'https://stackoverflow.com/questions/6040164/excel-vba-if-worksheetwsname-exists
    'https://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists
    
    Dim sht As Worksheet
    
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    On Error Resume Next
    
    Set sht = wb.Sheets(shtName)
    
    On Error GoTo 0
    
    SheetExists = Not sht Is Nothing
 End Function
 
Function VariantToString(A As Variant) As Collection
    ReDim S(UBound(A))

    For i = 0 To UBound(A) - 1
      S(i) = CStr(A(i))
    Next

    Set VariantToCollection = S
End Function

'SUBROUTINES
'SUBROUTINES
'SUBROUTINES

Sub ExcelDiet(Optional wb As Workbook)
    'http://www.vbaexpress.com/kb/getarticle.php?kb_id=83
    
    Dim j                   As Long
    Dim k                   As Long
    Dim LastRow        As Long
    Dim LastCol          As Long
    Dim ColFormula    As Range
    Dim RowFormula  As Range
    Dim ColValue        As Range
    Dim RowValue       As Range
    Dim Shp                 As Shape
    Dim ws                  As Worksheet
     
    Dim updating As Boolean
    Dim alerts As Boolean
    
    updating = Application.ScreenUpdating
    alerts = Application.DisplayAlerts
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
   
    On Error Resume Next
     
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
     
    For Each ws In wb.Worksheets
        With ws
            On Error Resume Next
            
             'Find the last used cell with a formula and value
             'Search by Columns and Rows
            Set ColFormula = _
                        .Cells.Find(What:="*", _
                                            After:=Range("A1"), _
                                            LookIn:=xlFormulas, _
                                            LookAt:=xlPart, _
                                            SearchOrder:=xlByColumns, _
                                            SearchDirection:=xlPrevious)
            Set ColValue = _
                        .Cells.Find(What:="*", _
                                            After:=Range("A1"), _
                                            LookIn:=xlValues, _
                                            LookAt:=xlPart, _
                                            SearchOrder:=xlByColumns, _
                                            SearchDirection:=xlPrevious)
            Set RowFormula = _
                        .Cells.Find(What:="*", _
                                            After:=Range("A1"), _
                                            LookIn:=xlFormulas, _
                                            LookAt:=xlPart, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlPrevious)
            Set RowValue = _
                        .Cells.Find(What:="*", _
                                            After:=Range("A1"), _
                                            LookIn:=xlValues, _
                                            LookAt:=xlPart, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlPrevious)
            On Error GoTo 0
             
             'Determine the last column
            If ColFormula Is Nothing Then
                LastCol = 0
            Else
                LastCol = ColFormula.Column
            End If
            
            If Not ColValue Is Nothing Then
                LastCol = Application.WorksheetFunction.max(LastCol, _
                                                            ColValue.Column)
            End If
             
             'Determine the last row
            If RowFormula Is Nothing Then
                LastRow = 0
            Else
                LastRow = RowFormula.Row
            End If
            
            If Not RowValue Is Nothing Then
                LastRow = Application.WorksheetFunction.max(LastRow, _
                                                                RowValue.Row)
            End If
             
             'Determine if any shapes are beyond the last row and last column
            For Each Shp In .Shapes
                j = 0
                k = 0
                
                On Error Resume Next
                
                j = Shp.TopLeftCell.Row
                k = Shp.TopLeftCell.Column
                
                On Error GoTo 0
                
                If j > 0 And k > 0 Then
                    Do Until .Cells(j, k).Top > Shp.Top + Shp.Height
                        j = j + 1
                    Loop
                    
                    If j > LastRow Then
                        LastRow = j
                    End If
                    
                    Do Until .Cells(j, k).Left > Shp.Left + Shp.Width
                        k = k + 1
                    Loop
                    
                    If k > LastCol Then
                        LastCol = k
                    End If
                End If
            Next Shp
             
            .Range(.Cells(1, LastCol + 1), _
                        .Cells(.Rows.Count, .Columns.Count)).EntireColumn.Delete
            .Range("A" & LastRow + 1 & _
                        ":A" & .Rows.Count).EntireRow.Delete
        End With
    Next ws
     
    Application.ScreenUpdating = updating
    Application.DisplayAlerts = alerts
End Sub

Sub ExcelSheetDiet(ws As Worksheet)
    Dim j               As Long
    Dim k               As Long
    Dim LastRow         As Long
    Dim LastCol         As Long
    Dim ColFormula      As Range
    Dim RowFormula      As Range
    Dim ColValue        As Range
    Dim RowValue        As Range
    Dim Shp             As Shape
             
    Dim updating As Boolean: updating = Application.ScreenUpdating
    Dim alerts As Boolean: alerts = Application.DisplayAlerts
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     
    On Error Resume Next
    
    With ws
        On Error Resume Next
        
         'Find the last used cell with a formula and value
         'Search by Columns and Rows
        Set ColFormula = _
                    .Cells.Find(What:="*", _
                                        After:=Range("A1"), _
                                        LookIn:=xlFormulas, _
                                        LookAt:=xlPart, _
                                        SearchOrder:=xlByColumns, _
                                        SearchDirection:=xlPrevious)
        Set ColValue = _
                    .Cells.Find(What:="*", _
                                        After:=Range("A1"), _
                                        LookIn:=xlValues, _
                                        LookAt:=xlPart, _
                                        SearchOrder:=xlByColumns, _
                                        SearchDirection:=xlPrevious)
        Set RowFormula = _
                    .Cells.Find(What:="*", _
                                        After:=Range("A1"), _
                                        LookIn:=xlFormulas, _
                                        LookAt:=xlPart, _
                                        SearchOrder:=xlByRows, _
                                        SearchDirection:=xlPrevious)
        Set RowValue = _
                    .Cells.Find(What:="*", _
                                        After:=Range("A1"), _
                                        LookIn:=xlValues, _
                                        LookAt:=xlPart, _
                                        SearchOrder:=xlByRows, _
                                        SearchDirection:=xlPrevious)
        On Error GoTo 0
         
         'Determine the last column
        If ColFormula Is Nothing Then
            LastCol = 0
        Else
            LastCol = ColFormula.Column
        End If
        
        If Not ColValue Is Nothing Then
            LastCol = Application.WorksheetFunction.max(LastCol, _
                                                        ColValue.Column)
        End If
         
         'Determine the last row
        If RowFormula Is Nothing Then
            LastRow = 0
        Else
            LastRow = RowFormula.Row
        End If
        
        If Not RowValue Is Nothing Then
            LastRow = Application.WorksheetFunction.max(LastRow, _
                                                            RowValue.Row)
        End If
         
         'Determine if any shapes are beyond the last row and last column
        For Each Shp In .Shapes
            j = 0
            k = 0
            
            On Error Resume Next
            
            j = Shp.TopLeftCell.Row
            k = Shp.TopLeftCell.Column
            
            On Error GoTo 0
            
            If j > 0 And k > 0 Then
                Do Until .Cells(j, k).Top > Shp.Top + Shp.Height
                    j = j + 1
                Loop
                
                If j > LastRow Then
                    LastRow = j
                End If
                
                Do Until .Cells(j, k).Left > Shp.Left + Shp.Width
                    k = k + 1
                Loop
                
                If k > LastCol Then
                    LastCol = k
                End If
            End If
        Next Shp
         
        .Range(.Cells(1, LastCol + 1), _
                    .Cells(.Rows.Count, .Columns.Count)).EntireColumn.Delete
        .Range("A" & LastRow + 1 & _
                    ":A" & .Rows.Count).EntireRow.Delete
    End With
     
    Application.ScreenUpdating = updating
    Application.DisplayAlerts = alerts
End Sub

Sub FormatSheet(sht As Worksheet)
    With sht
        .UsedRange.Font.name = "Calibri"
        .UsedRange.Font.Size = 11
        .UsedRange.HorizontalAlignment = xlHAlignLeft
    End With
End Sub

Sub MakeSheet(intoWb As Workbook, fromWb As Workbook, sheetName As String)
    'Unused, left in case format is useful.
    'This would fit nicely in mergesheets()
    'if sheetexists..
    .Worksheets.Add.name = sheetName
    ActiveSheet.Move After:=.Worksheets(.Worksheets.Count)
    .Worksheets(sheetName).Tab.ColorIndex = fromSheet.Tab.ColorIndex
    
    fromWb.Sheets(sheetName).UsedRange.Copy 'Probably really slow should investigate with timer
    
    .Worksheets(sheetName).Paste
    
    Application.CutCopyMode = False
End Sub

Sub ResetSheet(sht As Worksheet)
    With sht
        .Activate
        .Range("A2").Select
    End With
    
    Application.CutCopyMode = False
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Sub SheetCleanup(wb As Workbook, sheetNames() As Variant)
    Dim sName As Variant, sheetName As String
    Dim sht As Worksheet
    
    For Each sName In sheetNames
        sheetName = CStr(sName)
        
        If SheetExists(sheetName, wb) Then
                Set sht = wb.Sheets(sheetName)
                
                'ExcelSheetDiet sht
                'validationOn sht
                UnhideAllRowsAndColumns sht
                FormatSheet sht
                ResetSheet sht
                sht.Visible = True
        End If
    Next sName
End Sub

Sub UnhideAllRowsAndColumns(sht As Worksheet)
    With sht.Cells
        .EntireColumn.Hidden = False
        .EntireRow.Hidden = False
    End With
End Sub

Sub UsedRange(Optional wb As Workbook)
    Dim sht As Worksheet, msg As String
    
    'Above is how to use non-constant variables as default parameter.
    'Just check if the object is nothing or the variable is empty.
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    For Each sht In wb.Worksheets
        With sht
            msg = msg & vbCrLf & _
                        CStr(.UsedRange.Columns.Count & "|" & _
                                .UsedRange.Rows.Count & " ===> " & .name)
        End With
    Next
    Debug.Print msg
End Sub

Sub ValidationOn(sht As Worksheet)
    On Error Resume Next
    Dim rngValidation As Range, cell As Variant
    
    Set rngValidation = sht.UsedRange.SpecialCells(xlCellTypeAllValidation)
    
    If Not rngValidation Is Nothing Then
    
        For Each cell In rngValidation
            On Error Resume Next
            cell.Validation.ShowInput = True
        Next cell
    End If
End Sub

Sub ValidationOff(sht As Worksheet)
    On Error Resume Next
    Dim rngValidation As Range, cell As Variant
    
    Set rngValidation = sht.UsedRange.SpecialCells(xlCellTypeAllValidation)
    
    If Not rngValidation Is Nothing Then
        For Each cell In rngValidation
            On Error Resume Next
            cell.Validation.ShowInput = False
        Next cell
    End If
End Sub

Sub ValidationOffMultiSheet(wb As Workbook, sheetNames() As Variant)
    Dim sName As Variant
    Dim name As String
    
    For Each sName In sheetNames
        name = CStr(sName)
        
        If SheetExists(name, wb) Then
            ValidationOff wb.Sheets(name)
        End If
    Next sName
End Sub

Sub ValidationToggle(sht As Worksheet)
    On Error Resume Next
    Dim rngValidation As Range, cell As Variant
    
    Set rngValidation = sht.UsedRange.SpecialCells(xlCellTypeAllValidation)
    
    If Not rngValidation Is Nothing Then
        For Each cell In rngValidation
            On Error Resume Next
            cell.Validation.ShowInput = Not cell.Validation.ShowInput
        Next cell
    End If
End Sub





