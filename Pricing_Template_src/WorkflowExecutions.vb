Option Explicit

Sub Run_ImportAndMergeResearch_Button()
    Dim startTime As Single
    Dim oSheet As Worksheet
    
    startTime = Timer()
    
    Set oSheet = ThisWorkbook.ActiveSheet
    
    With Application
        .Calculation = xlCalculationManual
        .DisplayCommentIndicator = xlNoIndicator
        .DisplayNoteIndicator = False
        .EnableEvents = False
        .DisplayAlerts = False
        .ScreenUpdating = False
        
        GetPathsMergeThenReset oSheet
        
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .DisplayNoteIndicator = True
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .Calculation = xlCalculationAutomatic
    End With
    
    MsgBox "Operation took " & CStr(Round(Timer() - startTime, 2)) & " seconds."
End Sub

Sub Test_AllUsedRngToImmediate()
    UsedRange
End Sub

Sub Test_EnableScreenUpdating()
    Application.ScreenUpdating = True
End Sub

Sub Test_PrintCol()
    Dim test As String: test = "q8"
    
    Debug.Print ThisWorkbook.Sheets("revised description").Range(test).Font.Bold
    Debug.Print ThisWorkbook.Sheets("revised description").Range(test).Font.Color
    Debug.Print ThisWorkbook.Sheets("revised description").Range(test).Font.ColorIndex
    Debug.Print ThisWorkbook.Sheets("revised description").Range(test).Interior.Color
    Debug.Print ThisWorkbook.Sheets("revised description").Range(test).Interior.ColorIndex
End Sub

Sub Test_SubStringCounter()
    'https://www.mrexcel.com/forum/excel-questions/234028-visual-basic-applications-count-substrings-string.html
    
    Dim T As Double, S As String
    Dim X As Long, L As Long, LngPos As Long
    Const SubString As String = "A"
    
    S = "ABBBBACCC"
    T = Timer
    
    For L = 0 To 10000000
        LngPos = 1
        X = 0
        
        Do
            LngPos = InStrB(LngPos, S, SubString, vbBinaryCompare)
            
            If LngPos > 0 Then
                X = X + 1
                LngPos = LngPos + 1
            End If
        Loop Until LngPos = 0
    Next L
    
    MsgBox "InStrB Function: " & Timer - T '3ish seconds
    'tested already ubound and len, lenreplace. both were about 16 seconds
    'https://stackoverflow.com/questions/5193893/count-specific-character-occurrences-in-string
End Sub

Sub Test_WorkbookCleanUp()
    ExcelDiet
End Sub

