Option Explicit

Sub Show_MultiPricing_Form()
    On Error GoTo Err1
    dbSelect_form.Show

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub

Sub Show_SinglePricing_Form()
    On Error GoTo Err1
    singleQuery_form.Show

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub

Sub Run_Finalize_Pricing_Button()
    'This sub copies final prices from the pricing sheet to the revised sheet.
    On Error GoTo Err1
    
    Dim startTime As Single, i As Long, grey As Long
    Dim pricingRow As Long, revisedRow As Long
    Dim rd As Worksheet, pr As Worksheet
    Dim revisedCell As Range, pricingCell As Range
    Dim numErrRevised As String, numErrRevFlag As Boolean
    Dim numErrPricing As String, numErrPricFlag As Boolean
    Dim pricHold As String, revHold As String, hold() As String
    
    If (vbYes = MsgBox("CAUTION: Are you sure you wish to finalize?" & vbCrLf _
                        & "This will overwrite ""Suggested Price"" values " & _
                            "in the ""Revised Description"" sheet." & vbCrLf & _
                        "Non-empty prices are copied to " & _
                            "the first matching ""Master Description"" row.", _
                    vbYesNo, _
                    "PLEASE CONFIRM TO CONTINUE:")) Then
        
        startTime = Timer()
        
        pricingRow = conPricingSheetStartRow
        revisedRow = conRevisedSheetStartRow
        
        Set pr = ThisWorkbook.Sheets(conPricingSheetName)
        Set rd = ThisWorkbook.Sheets(conRevisedSheetName)
        
        numErrPricing = "Error(s) Finalizing at " & pr.name & ": " & vbCrLf & _
                        "Invalid New Suggested Price. (Non-Numeric)" _
                        & vbCrLf & "Row Numbers Not Copied From: "
        numErrRevised = "Error(s) Finalizing at " & rd.name & ": " & vbCrLf & _
                        "Invalid Original Suggested Price. (Non-Numeric)" _
                        & vbCrLf & "Row Numbers Not Copied Into: "
                        
        While pr.Cells(pricingRow, conPricingSheetMDescColumn).Value <> ""
            
            Set pricingCell = pr.Cells(pricingRow, _
                        conPricingSheetSuggPriceColumn)
            pricHold = Trim(pricingCell.Value)
            
            If pricHold <> "" And IsNumeric(pricingCell.Value) _
                                                And pricingCell.Value > 0 Then
                
                While rd.Cells(revisedRow, _
                                        conRevisedSheetMDescColumn).Value <> _
                            pr.Cells(pricingRow, conPricingSheetMDescColumn) _
                    And rd.Cells(revisedRow, _
                                        conRevisedSheetMDescColumn).Value <> ""
                                        
                    revisedRow = revisedRow + 1
                Wend
                
                If rd.Cells(revisedRow, _
                            conRevisedSheetMDescColumn).Value <> "" Then
                            
                    Set revisedCell = rd.Cells(revisedRow, _
                                    conRevisedSheetSuggPriceColumn)
                                    
                    With revisedCell
                        revHold = Trim(.Value)
                        
                        If revHold = "" Or (revHold <> "" And _
                                                        IsNumeric(.Value)) Then
                            .Value = revHold
                            
                            If .Value <> "" Then
                                .Value = Round(.Value, 2)
                            End If
                            
                            .Value = Trim(.Value & " " & Round(pricHold, 2))
                            
                            .Font.ColorIndex = 1 'black
                            .ClearFormats
                            
                            hold = Split(.Value, " ")
                                
                                'if 2 words, make 1st word gray.
                            If UBound(hold) = 1 Then
                                    .Characters(0, _
                                        Len(hold(0))).Font.ColorIndex = 16 '50Grey
                            End If
                        
                            'background med blue, font bold, align left
                            .Interior.Color = 15773696
                            .Font.Bold = True
                            .HorizontalAlignment = xlHAlignLeft
                            
                        Else
                            If Not numErrRevFlag Then
                                numErrRevFlag = True
                            Else
                                numErrRevised = numErrRevised & ", "
                            End If
                            
                            numErrRevised = numErrRevised & .Row
                        End If
                        
                    End With
                End If
                
            Else
                If pricHold <> "" Then
                    If Not numErrPricFlag Then
                        numErrPricFlag = True
                    Else
                        numErrPricing = numErrPricing & ", "
                    End If
                    
                    numErrPricing = numErrPricing & pricingRow
                End If
            End If
            
            pricingRow = pricingRow + 1
        Wend
        
        FormatSheet rd
        ResetSheet rd
        
        pr.Activate
        pr.Cells(2, 2).Select
                
        MsgBox "Finalized in: " & CStr(Timer() - startTime) & " seconds." & _
                        vbCrLf & "Be sure to double-check before submitting!"
                        
        If numErrRevFlag Then
            MsgBox numErrRevised
        End If
        
        If numErrPricFlag Then
            MsgBox numErrPricing
        End If
    End If
    
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub



