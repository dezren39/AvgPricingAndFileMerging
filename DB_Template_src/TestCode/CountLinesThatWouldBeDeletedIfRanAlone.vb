Option Compare Database

Private Sub btnRunQuery_Click()
On Error GoTo Err_btnRunQuery_Click
    Dim runTime As Single, startTime As Single: startTime = Timer()
    Dim in_trans As Boolean
    Dim cn As adodb.Connection
    Dim count As Double, result As String: result = "Lines Counted"
    Dim i As Integer

    Dim sSQL(1 to 6) As String
    'SQL(0) = "Select Count(*) FROM All_qa WHERE (All_qa.ID in (Select All_qa_ImportErrors.Row from All_qa_ImportErrors))"
    sSQL(1) = "Select Count(*) FROM All_qa WHERE (All_qa.[Mapped Description] LIKE '%(BUCKET)')"
    sSQL(2) = "Select Count(*) FROM All_qa WHERE (All_qa.[I/E] = 'Excluded')"
    sSQL(3) = "Select Count(*) FROM All_qa WHERE (All_qa.[Change Quantity] <> All_qa.[Quantity Link])"
    sSQL(4) = "Select Count(*) FROM All_qa WHERE (All_qa.[Change Quantity] < 1)"
    sSQL(5) = "Select Count(*) FROM All_qa WHERE (All_qa.[Average Price_Per] <= 0)"
    sSQL(6) = "Select Count(*) FROM All_qa WHERE (All_qa.[Average Price_Per] Is Null)"

    DBEngine.BeginTrans
        in_trans = True

        Set cn = New adodb.Connection
        With cn
            .ConnectionString = Application.CurrentDb.Name
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .Open
        End With

        For i = LBound(sSQL) To UBound(sSQL)
            Dim rs As adodb.Recordset

            runTime = Timer()
            Set rs = cn.Execute(sSQL(i))
            runTime = Timer() - runTime
            count = rs(0).Value
           '' MsgBox "passed"
            result = result + vbCrLf + CStr(i) + ": " + CStr(count) + " in " + CStr(runTime) + " seconds."
          ''  MsgBox result
            rs.Close
            Set rs = Nothing

        Next i

    DBEngine.CommitTrans
    in_trans = False
    cn.Close
    Set cn = Nothing
    runTime = Timer() - startTime
    MsgBox "Completed in " + CStr(runTime) + " seconds." + vbCrLf + result
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