Option Explicit

'*********************************************************************
'This program was created by Edward Woodford, amendments by Drewry Pope 03/2017
'*********************************************************************
'This module contains the main logic which is ran against each item entered
'   by either the single form or through runAllProds called by dbSelect_form.
'EDIT_ACCESS_QUERY prepares the SQL query and runs it against a selected db.
'PLOT_PRICES generates a workign sheet, runs calculations against the results
'   of the query and generates the graph for the proof_table.
'COPY_VALUES copies the relevant info from the working sheet to the proof_table
'   and Pricing sheets. If auto-ran, the working sheet is deleted.

Sub EDIT_ACCESS_QUERY(db_name As String, _
                                            prod_name As String, _
                                            qty As String, _
                                            max As String, _
                                            min As String, _
                                            Contains As String, _
                                            nutra As Boolean, _
                                            autoRun As Boolean)
    On Error GoTo Err1
    Application.ScreenUpdating = False
    
    'Initialize ODB connection, recordset, DB source, and other local variables
    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim sSQL As String, MyConn As String
    Dim pr As Worksheet
    Dim Rw As Long, Col As Long, c As Long
    Dim MyField, Location As Range
    Dim nutTxtChange As String, nutTxtLink As String
    
    MyConn = ThisWorkbook.Path & "\" & db_name
    
    Set pr = ThisWorkbook.Sheets(conPricingSheetName)
    
    'DEPRECATED: Nutrition Products are not priced.
    'Nutra DB uses different column names, these IF blocks compensate for that
    If nutra = True Then
       nutTxtChange = "[Average Quantity]"
       nutTxtLink = nutTxtChange
    End If
    
    If nutra = False Then
       nutTxtChange = "[Change Quantity]"
       nutTxtLink = "[Quantity Link]"
    End If
    
    'Create query, connect to DB, execute SQL, create RecordSet
    sSQL = "SELECT All_qa." & nutTxtChange & ", " _
                        & "All_qa.[Avg Total Price], " _
                        & "All_qa.[Mapped Description] " _
                & "FROM All_qa " _
                & "WHERE " & _
                        "(((All_qa.[Mapped Description]) = '" & _
                                                            prod_name & "')" _
                        & qty _
                        & max _
                        & min _
                        & Contains & ") " _
                & "ORDER BY All_qa.[Avg Total Price];"
    
    Set Cn = New ADODB.Connection
    
    With Cn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Open MyConn
        Set Rs = .Execute(sSQL)
    End With
    
    'If results were found, write RecordSet to new temporary sheet
    If Not Rs.EOF Then
        With ThisWorkbook
            .Sheets.Add After:=.Sheets(.Sheets.Count)
        End With
        
        'Set destination
        Set Location = [A2]
        Rw = Location.Row
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
        
        'Close unused dbconnection and range variables
        Set Location = Nothing
        Set Cn = Nothing
        
        'Unescape prod_name now that SQL string has completed
        prod_name = Replace(prod_name, "''", "'")
        
        'Send to next subroutine.
        Call PLOT_PRICES(prod_name, autoRun)
        
    'If no records for product, copy prod_name to Price Table
    ElseIf Rs.EOF Then
        'Starting at 7, While myRow not blank, add one to myRow
        Dim myRow As Long: myRow = conPricingSheetStartRow
        
        While pr.Cells(myRow, conPricingSheetMDescColumn).Value <> ""
            myRow = myRow + 1
        Wend
        
        'Paste copied prod_name to myRow'
        pr.Cells(myRow, conPricingSheetMDescColumn).Value = _
                                                Replace(prod_name, "''", "'")
    End If
       
Done:
    On Error Resume Next
    Set Cn = Nothing
    Exit Sub
    
Err1:
    Application.ScreenUpdating = True
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & _
                        "Error Number: " & err.Number & vbCrLf & _
                        "Error Source: RedefQry" & vbCrLf & _
                        "Error Description: " & err.Description, _
                    vbCritical, _
                    "An Error has Occured!"
    Resume Done
End Sub

Sub PLOT_PRICES(prod_name As String, autoRun As Boolean)
    On Error GoTo Err1
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet, LastRow As Long
    Dim CollA As Range, CollB As Range, CollC As Range, CollD As Range
    
    Set ws = ThisWorkbook.Sheets(ActiveSheet.name)
    
    With ws
        LastRow = .Cells(.Rows.Count, 2).End(xlUp).Row
           
        Set CollA = .Range("A2:A" & .Cells(.Rows.Count, "A").End(xlUp).Row)
        Set CollB = .Range("B2:B" & .Cells(.Rows.Count, "A").End(xlUp).Row)
        Set CollC = .Range("C2:C" & .Cells(.Rows.Count, "A").End(xlUp).Row)
        Set CollD = .Range("D2:D" & .Cells(.Rows.Count, "A").End(xlUp).Row)
        
        .Range("A1").Value = "Change Quantity"
        .Range("B1").Value = "Avg Total Price"
        .Range("C1").Value = "Average Price"
        .Range("D1").Value = "Distribution"
        .Range("F1").Value = "Product"
        .Range("G1").Value = "Median"
        .Range("H1").Value = "Mean"
        .Range("I1").Value = "STD"
        .Range("J1").Value = "Total Count"
    
        .Range("G2").Font.Bold = True
        
        CollC.Formula = "=ABS(B" & CollB.Cells(1, 1).Row & _
                                            "/ A" & CollA.Cells(1, 1).Row & ")"
        
        .Range("G2").Formula = "=QUARTILE.EXC(C:C,2)"
        .Range("H2").Formula = "=Subtotal(101, C:C)"
        .Range("I2").Formula = "=Subtotal(107, C:C)"
        .Range("J2").Formula = "=SUBTOTAL(102, C:C)"
        
        .Range("F2").Value = prod_name
        
        CollD.Formula = "=NORM.DIST(C2,$H$2,$I$2,false)"
        
    'Convert to Text
        .Range("G2").Value = .Range("G2").Value
        .Range("H2").Value = .Range("H2").Value
        .Range("I2").Value = .Range("I2").Value
        .Range("J2").Value = .Range("J2").Value
        
    '* Possibly needs Try Catch
        .Range("G2").Value = Round(.Range("G2").Value, 2)
        .Range("H2").Value = Round(.Range("H2").Value, 2)
        
    'Sort
        .Range("A1:D" & LastRow).Sort _
                                            key1:=.Range("C2:C" & LastRow), _
                                            order1:=xlAscending, _
                                            Header:=xlYes
        
    'Make Graph
        .Shapes.AddChart2(240, xlXYScatterSmooth).Select
        ActiveChart.SetSourceData Source:=Range( _
         "$C:$C,$D:$D")
        ActiveChart.HasTitle = True
        ActiveChart.ChartTitle.Text = prod_name
        ActiveChart.SetElement (msoElementLegendBottom)
        
        ActiveChart.Parent.Height = 210
        
        ActiveChart.Parent.Cut
        .Range("I4").PasteSpecial
        
    End With
    
    'Check Addins & Run Stats
    If AddIns("Analysis Toolpak").Installed Then
        Application.Run "ATPVBAEN.XLAM!Descr", ActiveSheet.Range("$C:$C"), _
        ActiveSheet.Range("$F$4"), "C", True, True, , , 95
    Else
        MsgBox "Analysis Toolpak is NOT installed: Please activate the " & _
                        "'Analysis Toolpak' add-in. You can do this by " & _
                        "going to File, Options, Add-ins, clicking 'Go' " & _
                        "and checking both Analysis Toolpak checkboxes."
    End If
    
    'Send to next subroutine.
    Call COPY_VALUES(prod_name, autoRun)

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub

Sub COPY_VALUES(prod_name As String, autoRun As Boolean)
    On Error GoTo Err1
    Application.ScreenUpdating = False
    
    Dim pr As Worksheet, ws As Worksheet
    Dim myRow As Long, i As Long, exists As Boolean
    
    Set pr = ThisWorkbook.Sheets(conPricingSheetName)
    Set ws = ThisWorkbook.Sheets(ActiveSheet.name)

    myRow = conPricingSheetStartRow

    'Copy Values to Price Table
    With pr
        'While A:myRow not blank, add one to myRow
        While .Cells(myRow, 1).Value <> ""
            myRow = myRow + 1
        Wend
        
        'Paste Pricing Summary to myRow'
        .Range(.Cells(myRow, conPricingSheetMDescColumn), _
                        .Cells(myRow, conPricingSheetCountColumn)).Value = _
                                                            ws.Range("F2:J2").Value
        .Range("B" & myRow).Font.Bold = True
    End With

    'Copy Values to Proof Table
    
    'If Proof_Table exists, set exists flag
    For i = 1 To Worksheets.Count
        If Worksheets(i).name = "Proof_Table" Then exists = True
    Next i
    
    With ThisWorkbook
        'If exists is Unset, add Proof_Table
        If Not exists Then
            Worksheets.Add.name = "Proof_Table"
            .Sheets("Proof_Table").Columns("C").ColumnWidth = 4
            .Sheets("Proof_Table").Columns("K").ColumnWidth = 8
            .Sheets("Proof_Table").Rows(1).Interior.Color = 6299648
        End If
        
        'Must declare/init Proof_table after validating it exists
        Dim proof As Worksheet
        
        Set proof = .Sheets("Proof_Table")
    End With
    
    With proof
        'While not empty, Check every 17th line, starting at 2, for emptiness
        myRow = 2
        
        While .Cells(myRow, 1).Value <> ""
            myRow = myRow + 17
        Wend
        
        'Paste relevant information to Proof_Table
        .Range("A" & myRow & ":B" & myRow + 16).Value = _
                                                ws.Range("F4:G20").Value
                                                
        ws.Pictures.Copy
        
        .Range("D" & myRow).PasteSpecial
        
        .Range("L" & myRow + 1).Value = prod_name
        
        'Format
        .Range("L" & myRow + 1).Font.Underline = True
        
        .Range("A" & myRow & ":B" & myRow).HorizontalAlignment = _
                                                xlHAlignCenterAcrossSelection
        .Range("A" & myRow & ":B" & myRow).Borders(xlEdgeBottom).Weight = _
                                                xlMedium
        
        .Rows(myRow + 16).Interior.Color = 6299648
    End With
    
    'If multi: disable alerts, delete Calculation sht, activate Pricing sht
    If autoRun = True Then
        Application.DisplayAlerts = False
        
        ws.Delete
        
        Application.DisplayAlerts = True
    'If not multi-pricing, activate Calculation sheet
    Else
        ws.Activate
    End If

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub




