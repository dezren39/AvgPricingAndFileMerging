'This form contains the text boxes and triggers to manually _
run a single product for pricing.
'Clear all values in Textboxes and the Combobox.'

Private Sub UserForm_Initialize()
    On Error GoTo Err1
    
    'Populate Combo Box
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

Public Sub DB_Combo_Box_Change()
    On Error GoTo Err1
    
    'Init db_name var and match Combo Box items to DB names.'
    Dim db_name As String
    
    If DB_Combo_Box.Value = "CC.mdb" Then
        db_name = "CC.mdb"
    End If
    
    If DB_Combo_Box.Value = "Nutritional.mdb" Then
        db_name = "Nutritional.mdb"
    End If
    
    If DB_Combo_Box.Value = "Pharma.mdb" Then
        db_name = "Pharma.mdb"
    End If
    
    'If DB Name is unavailable, inform user.'
    If Dir(ThisWorkbook.Path & "\" & db_name) = "" Then
        MsgBox (db_name & " - This database is currently unavailable.")
    End If
    
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub

Private Sub Clear_Btn_Click()
    Name_Text.Value = ""
    Max_Price_Text.Value = ""
    Min_Price_Text.Value = ""
    DB_Combo_Box.Value = ""
End Sub

Private Sub Get_data_btn_Click()
    On Error GoTo Err1
    
    Dim db_name As String, prod_name As String
    Dim qty As String, max As String, min As String
    Dim Contains As String, nutria As Boolean

    prod_name = LTrim(RTrim(Name_Text.Value))
    nutria = False
    
    'Either select qty any, only 1, or any value not = to 1.'
    If qty_one_btn.Value = True Then
        qty = " AND ((All_qa.[Change Quantity]) = 1) "
    End If
    
    If qty_more_btn.Value = True Then
        qty = " AND ((All_qa.[Change Quantity]) <> 1) "
    End If

    'If there is a min or max value in respective textbox, add that to the query.'
    If Max_Price_Text.Value = True Then
        max = " AND ((All_qa.[Average Price_Per])<" & Max_Price_Text.Value & ") "
    End If
    
    If Min_Price_Text.Value = True Then
        min = " AND ((All_qa.[Average Price_Per])>" & Min_Price_Text.Value & ") "
    End If

    'If there is a specific clinic description in respective textbox,'
    'Commented method: Description equals the text of the respective textbox.
    'New method: Description contains the text of the respective textbox.
    If contains_Text.Value = True Then
        Contains = " AND ((All_qa.[Clinic Description]) LIKE '%" & LTrim(RTrim(contains_Text.Value)) & "%')"
        'contains = (InStr(1,All_qa.[Clinic Description],'" & LTrim(RTrim(contains_Text.Value)) & "') > 0) "
    End If
    
    'Match Combo Box items to DB names.'
    If DB_Combo_Box.Value = "Client Connection" Then
        db_name = "CC.mdb"
    End If
    
    If DB_Combo_Box.Value = "GFK Pharma" Then
        db_name = "Pharma.mdb"
    End If
    
    If DB_Combo_Box.Value = "GFK Nutritional" Then
       db_name = "Nutritional.mdb"
       nutria = True
    End If
    
    'Run Calculations.'
    Call EDIT_ACCESS_QUERY(db_name, _
                                                prod_name, _
                                                qty, _
                                                max, _
                                                min, _
                                                Contains, _
                                                nutria, _
                                                False)
    
    .Sheets(conPricingSheetName).Cells(2, 2).Select
    
    'Unload form.
    Unload Me
    
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub




