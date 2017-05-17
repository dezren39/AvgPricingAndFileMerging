'*********************************************************************
'This program was created by Edward Woodford, amendments by Drewry Pope after 03/2017
'*********************************************************************
'This module contains the main logic which is ran against each item entered _
   'by either the single form or through runAllProds called by the dbSelect form.
'EDIT_ACCESS_QUERY prepares the SQL query and runs it against a selected db.
'PLOT_PRICES generates a workign sheet, runs calculations against the results _
'   of the query and generates the graph for the proof_table.
'COPY_VALUES copies the relevant info from the working sheet to the proof_table _
'   and Pricing sheets. If auto-ran, the working sheet is deleted.

Sub Show_Form()
On Error GoTo Err1
    query_form.Show
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub

Sub EDIT_ACCESS_QUERY(db_name As String, prod_name As String, qty As String, max As String, min As String, contains As String, nutra As Boolean, autoRun As Boolean)
On Error GoTo Err1
    Application.ScreenUpdating = False
    
    'Initialize ODB connection, recordset, DB source, and other local variables.
    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim sSQL As String, MyConn As String: MyConn = ThisWorkbook.Path & "\" & db_name
    Dim pr As Worksheet: Set pr = ThisWorkbook.Sheets("Pricing")
    Dim Rw As Long, Col As Long, c As Long
    Dim MyField, Location As Range
    Dim nutTxtChange As String, nutTxtLink As String
    
    'DEPRECATED: Nutrition Products are not priced.
    'Nutra DB uses different column names, these 2 IF blocks compensate for that.
    If nutra = True Then
       nutTxtChange = "[Average Quantity]"
       nutTxtLink = nutTxtChange
    End If
    If nutra = False Then
       nutTxtChange = "[Change Quantity]"
       nutTxtLink = "[Quantity Link]"
    End If
    
    'Create query, connect to DB, execute SQL, create RecordSet.
    sSQL = "SELECT All_qa." & nutTxtChange & ", All_qa.[Avg Total Price], All_qa.[Mapped Description] FROM All_qa WHERE (((All_qa.[Mapped Description]) = '" & prod_name & "')" & qty & max & min & contains & ") ORDER BY All_qa.[Avg Total Price];"
    Set Cn = New ADODB.Connection
    With Cn
             .Provider = "Microsoft.Jet.OLEDB.4.0"
             .Open MyConn
    Set Rs = .Execute(sSQL)
    End With
    
    'If results were found, write RecordSet to new temporary sheet.
    If Not Rs.EOF Then
        ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
        'Set destination
        Set Location = [A2]
        Rw = Location.row
        Col = Location.Column
        c = Col
        
        Do Until Rs.EOF
            For Each MyField In Rs.Fields
                Cells(Rw, c) = MyField
                c = c + 1
            Next MyField
            Rs.MoveNext
            Rw = Rw + 1
            c = Col
        Loop
        
        'Close unused dbconnection and range variables.
        Set Location = Nothing
        Set Cn = Nothing
        
        'Unescape prod_name now that SQL string has completed.
        prod_name = Replace(prod_name, "''", "'")
        
        'Send to next subroutine.
        Call PLOT_PRICES(prod_name, autoRun)
        
    'If no records for product, copy prod_name to Price Table
    ElseIf Rs.EOF Then
        'Starting at 7, While myRow not blank, add one to myRow.'
        Dim myRow As Long: myRow = 7
        While pr.Cells(myRow, 1).Value <> ""
            myRow = myRow + 1
        Wend
        
        'Paste copied prod_name to myRow'
        pr.Range("A" & myRow).Value = Replace(prod_name, "''", "'")
    End If
       
Done:
    On Error Resume Next
    Set Cn = Nothing
    Exit Sub
    
Err1:
    Application.ScreenUpdating = True
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.Number & vbCrLf & "Error Source: RedefQry" & vbCrLf & "Error Description: " & _
    Err.Description, vbCritical, "An Error has Occured!"
    Resume Done
End Sub


Sub PLOT_PRICES(prod_name As String, autoRun As Boolean)
On Error GoTo Err1
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(ActiveSheet.Name)
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    
    Set CollA = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
    Set CollB = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
    Set CollC = ws.Range("C2:C" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
    Set CollD = ws.Range("D2:D" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
    
    ws.Range("A1").Value = "Change Quantity"
    ws.Range("B1").Value = "Avg Total Price"
    ws.Range("C1").Value = "Average Price"
    ws.Range("D1").Value = "Distribution"
    ws.Range("F1").Value = "Product"
    ws.Range("G1").Value = "Median"
    ws.Range("H1").Value = "Mean"
    ws.Range("I1").Value = "STD"
    ws.Range("J1").Value = "Total Count"

    ws.Range("G2").Font.Bold = True
    
    CollC.Formula = "=ABS(B" & CollB.Cells(1, 1).row & "/ A" & CollA.Cells(1, 1).row & ")"
    
    ws.Range("G2").Formula = "=QUARTILE.EXC(C:C,2)"
    ws.Range("H2").Formula = "=Subtotal(101, C:C)"
    ws.Range("I2").Formula = "=Subtotal(107, C:C)"
    ws.Range("J2").Formula = "=SUBTOTAL(102, C:C)"
    
    ws.Range("F2").Value = prod_name
    
    CollD.Formula = "=NORM.DIST(C2,$H$2,$I$2,false)"
    
'Convert to Text
    ws.Range("G2").Value = ws.Range("G2").Value
    ws.Range("H2").Value = ws.Range("H2").Value
    ws.Range("I2").Value = ws.Range("I2").Value
    ws.Range("J2").Value = ws.Range("J2").Value
    
'* Possibly needs Try Catch
    ws.Range("G2").Value = Round(ws.Range("G2").Value, 2)
    ws.Range("H2").Value = Round(ws.Range("H2").Value, 2)
    
'Sort
    ws.Range("A1:D" & lastRow).Sort key1:=ws.Range("C2:C" & lastRow), _
    order1:=xlAscending, Header:=xlYes
    
'Make Graph
    ws.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=Range( _
     "$C:$C,$D:$D")
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = prod_name
    ActiveChart.SetElement (msoElementLegendBottom)
    
    ActiveChart.Parent.Height = 210
    
    ActiveChart.Parent.Cut
    ws.Range("I4").PasteSpecial
    
'Check Addins & Run Stats
    If AddIns("Analysis Toolpak").Installed Then
        Application.Run "ATPVBAEN.XLAM!Descr", ActiveSheet.Range("$C:$C"), _
        ActiveSheet.Range("$F$4"), "C", True, True, , , 95
    Else
        MsgBox "Analysis Toolpak is NOT installed: Please activate the 'Analysis Toolpak' add-in. You can do this by going to File, Options, Add-ins, click 'Go' and check both Analysis Toolpak add-ins."
    End If
    
'Send to next subroutine.
    Call COPY_VALUES(prod_name, autoRun)

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub

Sub COPY_VALUES(prod_name As String, autoRun As Boolean)
On Error GoTo Err1
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(ActiveSheet.Name)
    Dim pr As Worksheet: Set pr = ThisWorkbook.Sheets("Pricing")
    Dim myRow As Long: myRow = 7

'Copy Values to Price Table
    'While A:myRow not blank, add one to myRow
    While pr.Cells(myRow, 1).Value <> ""
        myRow = myRow + 1
    Wend
    
    'Paste Pricing Summary to myRow'
    pr.Range("A" & myRow & ":E" & myRow).Value = ws.Range("F2:J2").Value
    pr.Range("B" & myRow).Font.Bold = True

'Copy Values to Proof Table
    'Initialize exists flag
    Dim exists As Boolean
    
    'If Proof_Table exists, set exists flag
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "Proof_Table" Then
            exists = True
        End If
    Next i
    
    'If exists is Unset, add Proof_Table
    If Not exists Then
        Worksheets.Add.Name = "Proof_Table"
        ThisWorkbook.Sheets("Proof_Table").Columns("C").ColumnWidth = 4
        ThisWorkbook.Sheets("Proof_Table").Columns("K").ColumnWidth = 8
        ThisWorkbook.Sheets("Proof_Table").Rows(1).Interior.Color = 6299648
    End If
    
    'Must declare/init Proof_table after validating it exists
    Dim proof As Worksheet: Set proof = ThisWorkbook.Sheets("Proof_Table")
    
    'Check every 17th line, starting at 2, for emptiness. While not empty, continue.'
    myRow = 2
    While proof.Cells(myRow, 1).Value <> ""
        myRow = myRow + 17
    Wend
    
    'Paste relevant information to Proof_Table.'
    proof.Range("A" & myRow & ":B" & myRow + 16).Value = ws.Range("F4:G20").Value
    ws.Pictures.Copy
    proof.Range("D" & myRow).PasteSpecial
    proof.Range("L" & myRow + 1).Value = prod_name
    'Format
    proof.Range("L" & myRow + 1).Font.Underline = True
    proof.Range("A" & myRow & ":B" & myRow).HorizontalAlignment = xlHAlignCenterAcrossSelection
    proof.Range("A" & myRow & ":B" & myRow).Borders(xlEdgeBottom).Weight = xlMedium
    proof.Rows(myRow + 16).Interior.Color = 6299648

    'If multi-pricing, disable alerts, delete Calculation sheet, activate Pricing sheet.'
    If autoRun = True Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    'If not multi-pricing, activate Calculation sheet.'
    Else
        ws.Activate
    End If

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub

