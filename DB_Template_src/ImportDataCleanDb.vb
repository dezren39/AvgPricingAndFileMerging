'Author: Drew'
'Date: 4/26/2017'
'Version: 1.5'
'Description: Script to import and clean
'             data gathered by All_qa.bat'

Option Compare Database

Sub CleanDB()
    'Initialize all variables, record time started.'
    Dim runTime As Single, startTime As Single: startTime = Timer()
    Dim in_trans As Boolean
    Dim cn As ADODB.Connection
    Dim i As Integer, lboundSQL As Integer: lboundSQL = 1 'Disable Import Errors SQL until table proven exists.
    Dim sSQL(7) As String, table As String: table = "All_qa"
    
    RenameTableLike "*ImportErrors*", "All_qa_ImportErrors"
    table = GetTableLike("*All_qa_aggregated_*", table)
    
    sSQL(0) = "DELETE " & table & ".* FROM " & table & " WHERE (" & table & ".ID in (Select All_qa_ImportErrors.Row from All_qa_ImportErrors))"
    sSQL(1) = "DELETE " & table & ".* FROM " & table & " WHERE (" & table & ".[Mapped Description] LIKE '%(BUCKET)')"
    sSQL(2) = "DELETE " & table & ".* FROM " & table & " WHERE (" & table & ".[I/E] = 'Excluded')"
    sSQL(3) = "DELETE " & table & ".* FROM " & table & " WHERE (" & table & ".[Change Quantity] <> " & table & ".[Quantity Link])"
    sSQL(4) = "DELETE " & table & ".* FROM " & table & " WHERE (" & table & ".[Change Quantity] < 1)"
    sSQL(5) = "DELETE " & table & ".* FROM " & table & " WHERE (" & table & ".[Change Quantity] LIKE '%.%')"
    sSQL(6) = "DELETE " & table & ".* FROM " & table & " WHERE (" & table & ".[Average Price_Per] <= 0)"
    sSQL(7) = "DELETE " & table & ".* FROM " & table & " WHERE (" & table & ".[Average Price_Per] Is Null)"
  
    If TableExists("All_qa_ImportErrors") Then lboundSQL = 0

    'Access default workspace and begin transaction.'
    DBEngine.BeginTrans
        'Set bool transaction flag (for errors), and initialize/define/open db connection'
        in_trans = True
        Set cn = New ADODB.Connection
        With cn
            .ConnectionString = Application.CurrentDb.Name
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .Open
        End With
        
        'For each SQL statement, execute the statement without buffering for recordset. (It's a delete, nothing to return)
        For i = lboundSQL To UBound(sSQL)
            cn.Execute sSQL(i), , adExecuteNoRecords
        Next i
    DBEngine.CommitTrans
    
    
    'Unset bool transaction flag, close and erase ADO connection '
    in_trans = False
    cn.Close
    Set cn = Nothing
    
    'Must delete ImportErrors first, so that TableExist works properly. Idk why.
    DeleteTableLike "*ImportErrors*"
    
    If Not TableExists("All_qa") Then
        RenameTableLike "*All_qa*", "All_qa"
    End If
    
    'Record Runtime, Display Confirmation
    runTime = Timer() - startTime
    MsgBox "Completed clean in: " + CStr(runTime) + " seconds." + vbCrLf + "Don't forget to Compact and Repair!" _
            + vbCrLf + "Database Tools -> Compact and Repair Database"
End Sub

Sub ImportData()
    Dim startTime As Single: startTime = Timer()
    Dim filePath As String: filePath = getFileName()
    If filePath <> "" Then
        getImportSpecAndUpdate "C:\Users\dpope\Desktop\Pricing\DB\all_qa_aggregated.txt", filePath
        DoCmd.RunSavedImportExport "Import-all_qa_aggregated"
        getImportSpecAndUpdate filePath, "C:\Users\dpope\Desktop\Pricing\DB\all_qa_aggregated.txt"
        MsgBox "Completed import in: " + CStr(Timer() - startTime) + " seconds." + vbCrLf + "Don't forget to clean the data!" _
                + vbCrLf + "Import and Clean -> Clean DB"
    End If
End Sub

Sub MergeTables()
    Dim runTime As Single, startTime As Single: startTime = Timer()
    Dim SQL As String, inputVar As String: inputVar = InputBox("Table Name?", "Merge Table w/All_qa Table:", appendDate("All_qa_aggregated_"))
    
    If inputVar <> "" Then
        If TableExists(inputVar) Then
            SQL = "INSERT INTO All_qa SELECT " & inputVar & ".* FROM " & inputVar
            DoCmd.RunSQL SQL
            DeleteTableLike inputVar
        
            'Record Runtime, Display Confirmation
            runTime = Timer() - startTime
            MsgBox "Completed merge in: " + CStr(runTime) + " seconds." + vbCrLf + "Don't forget to Compact and Repair!" _
            + vbCrLf + "Database Tools -> Compact and Repair Database"
        Else
            MsgBox "Invalid Table Name: " & inputVar
        End If
    End If
End Sub

Function appendDate(message)
    Dim dateSplit() As String: dateSplit() = Split(CStr(Date), "/", 3)

    For i = 0 To 2
        If (Len(dateSplit(i)) = 1) Then
            message = message & "0"
        ElseIf (Len(dateSplit(i)) = 4) Then
            dateSplit(i) = Right(dateSplit(i), 2)
        End If
        message = message & dateSplit(i)
    Next
    appendDate = message
End Function

Function getFileName() As String
    Dim initName As String: initName = "all_qa_aggregated_"

    initName = appendDate(initName)
    
    With Application.FileDialog(3) ' 3=msoFileDialogFilePicker
        .InitialFileName = (CurrentProject.Path & "\") ' start in this folder
        .AllowMultiSelect = False 'may enable in the future, will need rework
        .Title = "Select an all_qa file to import:"
        .InitialFileName = initName 'all_qa_aggregator_(dateToday)
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt"
        .Filters.Add "All Files", "*.*"
        .Show

        If .SelectedItems.Count > 0 Then
           FileName = .SelectedItems(1)
        End If
    End With
    
    getFileName = FileName
End Function

Function getImportSpecAndUpdate(fileIn As String, fileOut As String)
    Dim allQaImportSpec As ImportExportSpecification
    Set allQaImportSpec = CurrentProject.ImportExportSpecifications.Item("Import-all_qa_aggregated")
    allQaImportSpec.XML = Replace(allQaImportSpec.XML, fileIn, fileOut)
    Set allQaImportSpec = Nothing
End Function

Function GetTableLike(nameLike As String, Optional alternate As String) As String
    Dim iTable As DAO.TableDef
    For Each iTable In CurrentDb.TableDefs
        If iTable.Name Like nameLike Then
            GetTableLike = iTable.Name
            Exit Function
        End If
    Next iTable
    GetTableLike = alternate
End Function

Function TableExists(table) As Boolean
    For Each iTable In CurrentDb.TableDefs
        If iTable.Name = table Then
            TableExists = True
            Exit Function
        End If
    Next iTable
    TableExists = False
End Function

'http://datapigtechnologies.com/blog/index.php/clearing-access-importerror-tables/'
Sub DeleteTableLike(nameLike)
    Dim iTable As DAO.TableDef
     
    For Each iTable In CurrentDb.TableDefs
        If iTable.Name Like nameLike Then
            CurrentDb.TableDefs.Delete iTable.Name
        End If
    Next iTable
End Sub

Sub RenameTableLike(nameLike As String, newName As String)
    Dim iTable As DAO.TableDef
    For Each iTable In CurrentDb.TableDefs
        If iTable.Name Like nameLike Then
            iTable.Name = newName
        End If
    Next iTable
End Sub

