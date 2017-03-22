Option Compare Database

Private Sub btnRunQuery_Click()
On Error GoTo Err_btnRunQuery_Click
    Dim runTime As Single, startTime As Single: startTime = Timer()
    Dim in_trans As Boolean
    Dim cn As ADODB.Connection
    Dim sSQL As String: sSQL = "DELETE All_qa.* FROM All_qa WHERE (All_qa.ID in (Select All_qa_ImportErrors.Row from All_qa_ImportErrors))"
    
    DBEngine.BeginTrans

        in_trans = True
        Set cn = New ADODB.Connection
        With cn
            .ConnectionString = Application.CurrentDb.Name
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .Open
            .Execute sSQL, , adExecuteNoRecords
        End With

    DBEngine.CommitTrans

    in_trans = False
    cn.Close
    Set cn = Nothing
    runTime = Timer() - startTime
    DeleteImportErrorTables
    MsgBox "Completed in " + CStr(runTime) + " seconds."
    Exit Sub


Err_btnRunQuery_Click:
    If in_trans = True Then
        DBEngine.Rollback
    End If
    If cn.State = adStateOpen Then cn.Close
    Set cn = Nothing
    MsgBox "Error " + Err.Description
    Exit Sub
End Sub


Sub DeleteImportErrorTables()
    Dim iTable As DAO.TableDef
     
    For Each iTable In CurrentDb.TableDefs
        If iTable.Name Like "*ImportErrors*" Then
            CurrentDb.TableDefs.Delete iTable.Name
        End If
    Next iTable
End Sub