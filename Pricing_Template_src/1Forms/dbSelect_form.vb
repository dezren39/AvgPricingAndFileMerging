'This form contains the 2 functions:
'UserForm_Initialize configures the combo box on formload.
'get_data_btn_click is ran when the run query button is pressed.
'PrepThenRunAllProdsThenCleanUp is executed. This code is ran to prepare for
'running MTRF_PricingLogics's runAllProds sub. runAllProds is executed,
'averaging pricing for each item on the pricing sheet. code is ran after
'runAllProds to copy specific values if they meet certain requirements.

Private Sub UserForm_Initialize()
    On Error GoTo Err1
    
    'Populate Combo Box.'
    DB_Combo_Box.AddItem "CC.mdb"
    DB_Combo_Box.AddItem "Nutritional.mdb"
    DB_Combo_Box.AddItem "Pharma.mdb"
    DB_Combo_Box.Value = "CC.mdb"

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub

Private Sub Get_data_btn_Click()
    On Error GoTo Err1
    dbSelect_form.Hide
    Application.ScreenUpdating = False
    
    Dim startTime As Single
    Dim db_name As String
    
    startTime = Timer()
    
    'Match Combo Box items to DB names. Assign one to db_name'
    If DB_Combo_Box.Value = "CC.mdb" Then
        db_name = "CC.mdb"
    End If
    
    If DB_Combo_Box.Value = "Pharma.mdb" Then
        db_name = "Pharma.mdb"
    End If
    
    If DB_Combo_Box.Value = "Nutritional.mdb" Then
        db_name = "Nutritional.mdb"
        nutria = True
    End If

    PrepThenRunAllProdsThenCleanUp db_name

    Sheets(conPricingSheetName).Activate
    Sheets(conPricingSheetName).Cells(2, 2).Select
    
    Application.ScreenUpdating = True
    
    MsgBox "Completed pricing in: " + CStr(Timer() - startTime) + " seconds." _
                + vbCrLf + "Now it's time to evaluate the results." _
                + vbCrLf + "Use good judgement and don't forget to finalize!"
                
    Unload Me

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub



