Option Explicit

Sub GetPathsMergeThenReset(original As Worksheet)
    MergeSheetsFromPaths ThisWorkbook, _
                            RetrieveExcelPathsInFolderPath(GetFolderPath), _
                            RetrieveImportSheetsAsVariant
                            
    ResetSheet original
End Sub

Sub MergeSheetsFromPaths(intoWb As Workbook, _
                                                sourceFilePaths() As Variant, _
                                                sheetNames() As Variant)
    'http://www.ozgrid.com/forum/showthread.php?t=169266
    'https://vba-useful.blogspot.fr/2013/12/how-do-i-retrieve-data-from-another.html
    
    Dim fromWb As Excel.Workbook, sht As Worksheet
    Dim sPath As Variant, sName As Variant, sheetName As String
    
    If CStr(sourceFilePaths(0)) <> "ERROR" Then
        ValidationOffMultiSheet intoWb, sheetNames
        
        For Each sPath In sourceFilePaths
            Set fromWb = Application.Workbooks.Add(sPath)
            
            MergeSheets intoWb, fromWb, sheetNames
            
            fromWb.Close SaveChanges:=False
        Next sPath
        
        SheetCleanup intoWb, sheetNames
    Else
        MsgBox "No Excel files found."
    End If
End Sub

Sub MergeSheets(ByRef intoWb As Workbook, _
                                ByRef fromWb As Workbook, _
                                sheetNames() As Variant)
                                
    Dim sName As Variant, sheetName As String
    Dim intoSheet As Worksheet, fromSheet As Worksheet
    
    For Each sName In sheetNames
        sheetName = CStr(sName)
        
        If SheetExists(sheetName, fromWb) Then
            If SheetExists(sheetName, intoWb) Then

                Set fromSheet = fromWb.Sheets(sheetName)
                Set intoSheet = intoWb.Worksheets(sheetName)
                
                UnhideAllRowsAndColumns fromSheet
                ExcelSheetDiet fromSheet 'slow?
                ValidationOff fromSheet
                
                UnhideAllRowsAndColumns intoSheet
                
                MergeSheet intoSheet, fromSheet, sheetName
            End If
        End If
    Next sName
End Sub

Sub MergeSheet(intoSheet As Worksheet, _
                            fromSheet As Worksheet, _
                            sName As String)
                            
    Dim cNames As New Collection, cName As Variant
    Dim intoCIds As Collection, fromCIds As Collection
    Dim sentinel As Variant, intoCol As Long, fromCol As Long
    Dim intoLastRow As Long, fromLastRow As Long
    
    Set cNames = CollectColumnNames(intoSheet)
    
    Set intoCIds = CollectColumnIds(cNames, RetrieveColumns(intoSheet))
    Set fromCIds = CollectColumnIds(cNames, RetrieveColumns(fromSheet))

    sentinel = GetSentinelColName(intoCIds)
    
    If sentinel <> vbNull And Contains(fromCIds, sentinel) Then
        If Contains(intoCIds, sentinel) Then
            intoLastRow = LastOccupiedRowNumInCol(intoSheet, _
                                        intoCIds(sentinel))
            fromLastRow = LastOccupiedRowNumInCol(fromSheet, _
                                        fromCIds(sentinel))
            
            For Each cName In cNames
                If Contains(fromCIds, cName) Then
                    fromCol = fromCIds(cName)
                    intoCol = intoCIds(cName)
                    
                    MergeColumns intoSheet, _
                                                fromSheet, _
                                                intoCol, _
                                                fromCol, _
                                                intoLastRow + 1, _
                                                fromLastRow
                End If
            Next cName
        End If
    End If
End Sub

Sub MergeColumns(intoSheet As Worksheet, _
                                fromSheet As Worksheet, _
                                intoCol As Long, _
                                fromCol As Long, _
                                intoStartRow As Long, _
                                fromLastRow, _
                                Optional fromStartRow = 2)
                                
    Dim into As Range, from As Range
    
    If fromStartRow <= fromLastRow Then
        With intoSheet
            Set from = fromSheet.Range(fromSheet.Cells(fromStartRow, _
                            fromCol), fromSheet.Cells(fromLastRow, fromCol))
                            
            Set into = .Range(.Cells(intoStartRow, intoCol), _
                        .Cells(intoStartRow + fromLastRow - fromStartRow, _
                                    intoCol))
                                    
            from.Copy
                                    
            into.PasteSpecial xlPasteAll
        End With
        
        Application.CutCopyMode = False
    End If
End Sub

