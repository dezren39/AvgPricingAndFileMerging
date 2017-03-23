'Author: Drew'
'Date: 3/23/2017'
'Version: 1.1'
'Description: Script to clean up data gathered by All_qa.txt files for DSN.'


Option Compare Database

Private Sub btnRunQuery_Click()
On Error GoTo Err_btnRunQuery_Click
    'Initialize all variables, record time started.'
    Dim runTime As Single, startTime As Single: startTime = Timer()
    Dim in_trans As Boolean
    Dim cn As ADODB.Connection
    Dim i As Integer
    Dim sSQL(7) As String
    
    'Rename ImportError to match how it's named in sSQL(0): All_qa_ImportErrors
    RenameImportErrorTable
    
    'Define what to delete, based on recommendations.'
    sSQL(0) = "DELETE All_qa.* FROM All_qa WHERE (All_qa.ID in (Select All_qa_ImportErrors.Row from All_qa_ImportErrors))"
    sSQL(1) = "DELETE All_qa.* FROM All_qa WHERE (All_qa.[Mapped Description] LIKE '%(BUCKET)')"
    sSQL(2) = "DELETE All_qa.* FROM All_qa WHERE (All_qa.[I/E] = 'Excluded')"
    sSQL(3) = "DELETE All_qa.* FROM All_qa WHERE (All_qa.[Change Quantity] <> All_qa.[Quantity Link])"
    sSQL(4) = "DELETE All_qa.* FROM All_qa WHERE (All_qa.[Change Quantity] < 1)"
    sSQL(5) = "DELETE All_qa.* FROM All_qa WHERE (All_qa.[Change Quantity] LIKE '%.%')"
    sSQL(6) = "DELETE All_qa.* FROM All_qa WHERE (All_qa.[Average Price_Per] <= 0)"
    sSQL(7) = "DELETE All_qa.* FROM All_qa WHERE (All_qa.[Average Price_Per] Is Null)"
    
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
        For i = LBound(sSQL) To UBound(sSQL)
            'runTime = Timer()
            cn.Execute sSQL(i), , adExecuteNoRecords
            'runTime = Timer() - runTime
            'MsgBox Cstr(i) + ": " + Cstr(runTime) + " seconds."
        Next i
    'Commit transactions.'
    DBEngine.CommitTrans

    'Unset bool transaction flag, close and erase ADO connection '
    in_trans = False
    cn.Close
    Set cn = Nothing
    
    'Delete ImportErrors'
    DeleteImportErrorTables

    'Record Runtime, Display Confirmation
    runTime = Timer() - startTime
    MsgBox "Completed in: " + CStr(runTime) + " seconds." + VbCrLf + "Don't forget to Compact and Repair!" _
            + VbCrLf + "Database Tools -> Compact and Repair Database"

'If Error, try to rollback partial executions, erase ADO connection, display Error.'
Err_btnRunQuery_Click:
    If err.Number <> 0 Then
        If in_trans = True Then
            DBEngine.Rollback
        End If
        Set cn = Nothing
        MsgBox "Error " + Err.Description
        Exit Sub
    End If
End Sub

'Function to delete ImportErrors.'
'http://datapigtechnologies.com/blog/index.php/clearing-access-importerror-tables/'
Sub DeleteImportErrorTables()
    Dim iTable As DAO.TableDef
     
    For Each iTable In CurrentDb.TableDefs
        If iTable.Name Like "*ImportErrors*" Then
            CurrentDb.TableDefs.Delete iTable.Name
        End If
    Next iTable
End Sub

'Function to rename ImportError table for SQL querying.'
Sub RenameImportErrorTable()
    Dim iTable As DAO.TableDef
     
    For Each iTable In CurrentDb.TableDefs
        If iTable.Name Like "*ImportErrors*" Then
           iTable.Name = "All_qa_ImportErrors"
           Exit Sub
        End If
    Next iTable
End Sub
