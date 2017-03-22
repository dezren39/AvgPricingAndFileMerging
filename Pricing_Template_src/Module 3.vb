'*********************************************************************
'This program was created by Edward Woodford, amendments by Drewry Pope 03/2017
'*********************************************************************
Sub Show_Form()
    query_form.Show
End Sub

Sub Show_DB_Form()
    dbSelect.Show
End Sub

Sub EDIT_ACCESS_QUERY(db_name As String, prod_name As String, qty As String, max As String, min As String, contains As String, nurtia As Boolean, autoRun As Boolean)
    
    'Initialize ODB connection, recordset, DB source, and other local variables.
    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim sSQL As String, MyConn As String: MyConn = ThisWorkbook.Path & "\" & db_name
    Dim Rw As Long, Col As Long, c As Long
    Dim MyField, Location As Range
    Dim nutTxtChange As String, nutTxtLink As String
    
    'Nutra DB uses different column names, these 2 IF blocks compensate for that.
    If nurtia = True Then
       nutTxtChange = "[Average Quantity]"
       nutTxtLink = nutTxtChange
    End If
    
    If nurtia = False Then
       nutTxtChange = "[Change Quantity]"
       nutTxtLink = "[Quantity Link]"
    End If
    
    'Create query, default recordset parameters are hardcoded here.
     sSQL = "SELECT All_qa." & nutTxtChange & ", All_qa.[Avg Total Price], All_qa.[Mapped Description] FROM All_qa WHERE (((All_qa.[Mapped Description]) = '" & prod_name & "') AND ((All_qa.[Average Price_Per]) <  500) AND (All_qa." & nutTxtLink & ") = (All_qa." & nutTxtChange & ") AND ((All_qa." & nutTxtLink & ") >= 1) AND ((All_qa." & nutTxtLink & ") = Int(All_qa." & nutTxtLink & ")) " & qty & max & min & contains & ") ORDER BY All_qa.[Avg Total Price];"
    
    'Connect to DB, execute SQL, create RecordSet.
    Set Cn = New ADODB.Connection
    With Cn
       .Provider = "Microsoft.Jet.OLEDB.4.0"
       .Open MyConn
    Set Rs = .Execute(sSQL)
    End With
    
    'Write RecordSet to results area
    If Not Rs.EOF Then
        ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Range("A2").Select
        
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
        
        Set Location = Nothing
        Set Cn = Nothing
        
        'Send to next subroutine.
        Call PLOT_PRICES(prod_name, autoRun)
    Else
        'MsgBox ("No records for " & prod_name & " found!")
    End If
       
Error_Handler_Exit:
    On Error Resume Next
       Set qdf = Nothing
       Exit Sub
    
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.Number & vbCrLf & "Error Source: RedefQry" & vbCrLf & "Error Description: " & _
    Err.Description, vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Sub

Sub PLOT_PRICES(prod_name As String, autoRun As Boolean)
    lastRow = Cells(Rows.Count, 2).End(xlUp).row
    
    Set CollA = Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).row)
    Set CollB = Range("B2:B" & Cells(Rows.Count, "A").End(xlUp).row)
    Set CollC = Range("C2:C" & Cells(Rows.Count, "A").End(xlUp).row)
    Set CollD = Range("D2:D" & Cells(Rows.Count, "A").End(xlUp).row)
    
    Range("A1").Value = "Change Quantity"
    Range("B1").Value = "Avg Total Price"
    Range("C1").Value = "Average Price"
    Range("D1").Value = "Distribution"
    Range("F1").Value = "Product"
    Range("G1").Value = "Median"
    Range("H1").Value = "Mean"
    Range("I1").Value = "STD"
    Range("J1").Value = "Total Count"
    
    Range("G2").Font.Bold = True
    
    CollC.Formula = "=ABS(B" & CollB.Cells(1, 1).row & "/ A" & CollA.Cells(1, 1).row & ")"
    
    Range("G2").Formula = "=QUARTILE.EXC(C:C,2)"
    Range("H2").Formula = "=Subtotal(101, C:C)"
    Range("I2").Formula = "=Subtotal(107, C:C)"
    Range("J2").Formula = "=SUBTOTAL(102, C:C)"
    
    Range("F2").Value = prod_name
    
    CollD.Formula = "=NORM.DIST(C2,$H$2,$I$2,false)"
    
'Convert to Text
    Range("G2").Value = Range("G2").Value
    Range("H2").Value = Range("H2").Value
    Range("I2").Value = Range("I2").Value
    Range("J2").Value = Range("J2").Value
    
'* Possibly needs Try Catch
    Range("G2").Value = Round(Range("G2").Value, 2)
    Range("H2").Value = Round(Range("H2").Value, 2)
    
'Sort
    Range("A1:D" & lastRow).Sort key1:=Range("C2:C" & lastRow), _
    order1:=xlAscending, Header:=xlYes
    
'Make Graph
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=Range( _
     "$C:$C,$D:$D")
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = prod_name
    ActiveChart.SetElement (msoElementLegendBottom)
    
    ActiveChart.Parent.Height = 210
    
    ActiveChart.Parent.Cut
    Range("I4").Select
    ActiveSheet.Pictures.Paste.Select
    
'Check Addins & Run Stats
    If AddIns("Analysis Toolpak").Installed Then
        Range("C2:C" & lastRow).Select
        Application.Run "ATPVBAEN.XLAM!Descr", ActiveSheet.Range("$C:$C"), _
        ActiveSheet.Range("$F$4"), "C", True, True, , , 95
    Else
        MsgBox "Analysis Toolpak is NOT installed: Please activate the 'Analysis Toolpak' add-in. You can do this by going to File, Options, Add-ins, click 'Go' and check both Analysis Toolpak add-ins."
    End If
    
'Send to next subroutine.
    Call COPY_VALUES(prod_name, autoRun)
End Sub

Sub COPY_VALUES(prod_name As String, autoRun As Boolean)
'Copy Values
    'Assign active sheet to var LstSht, expected to be Calculation sheet.
    Dim LstSht As String
    LstSht = ActiveSheet.Name

    'Copy Pricing Summary, activate Pricing sheet.'
    Range("F2:J2").Select
    Selection.Copy
    Sheets("Pricing").Select

'Copy Values to Price Table
    'Initialize myRow to 7'
    Dim myRow As Long: myRow = 7
    
    'While A:myRow not blank, add one to myRow.'
    While Cells(myRow, 1).Value <> ""
    myRow = myRow + 1
    Wend
    
    'Paste copied Pricing Summary to myRow'
    Range("A" & myRow).Select
    ActiveSheet.Paste

'Copy Values to Proof Table
    'Initialize exists flag and activate Calculation Sheet.'
    Dim exists As Boolean
    Worksheets(LstSht).Activate

    'Select relevant info and copy.'
    Range("F4:P20").Select
    Selection.Copy
    
    'If Proof_Table exists, set exists flag.'
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "Proof_Table" Then
        exists = True
        End If
    Next i
    
    'If exists is Unset, add Proof_Table'
    If Not exists Then
        Worksheets.Add.Name = "Proof_Table"
    End If
    
    'Select Proof_Table, reset myRow to 1'
    Sheets("Proof_Table").Select
    myRow = 1
    
    'Check every 17th line, starting at 1, for emptiness. While not empty, continue.'
    While Cells(myRow, 1).Value <> ""
        myRow = myRow + 17
    Wend
    
    'Paste relevant information to Proof_Table.'
    Range("A" & myRow).Select
    ActiveSheet.Paste
    Range("C" & myRow).Value = prod_name
    Rows(myRow + 16).Interior.Color = 6299648

    'If multi-pricing, disable alerts, delete Calculation sheet, activate Pricing sheet.'
    If autoRun = True Then
        Application.DisplayAlerts = False
        Worksheets(LstSht).Delete
        Application.DisplayAlerts = True
        Sheets("Pricing").Activate
        
    'If not multi-pricing, activate Calculation sheet.'
    Else
        Worksheets(LstSht).Activate
    End If
End Sub

