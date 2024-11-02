Option Explicit

'Global Constants
Const FirstColumnAreaToColor    As String = "B"
Const LastColumnAreaToColor     As String = "F"
Const GradeColumn               As String = "D"
Const DateLastReviewColumn      As String = "F"
Const DataFirstRow              As Integer = 7
Const HistoryStartColumn        As String = "G"
Const HistorySecondColmn        As String = "H"
Const HistoryBeforeLastColumn   As String = "J"
Const HistoryEndColumn          As String = "K"


Sub QAE_SetGradeGreat()
    Call QAE_UpdateGrade("Good")
End Sub

Sub QAE_SetGradeGreatOk()
    Call QAE_UpdateGrade("Ok")
End Sub

Sub QAE_SetGradeGreatBad()
    Call QAE_UpdateGrade("Bad")
End Sub

Sub QAE_UpdateGrade(Grade As String)
        
    Dim today               As Date
    Dim selectedRow         As Integer
    Dim selectedRowGrade    As String
    Dim promptAnswer        As Integer
    
    selectedRow = ActiveCell.row
    today = Date
    selectedRowGrade = Cells(selectedRow, GradeColumn)
    
    'Prevent out of boundaries select
    If (selectedRow < DataFirstRow) Then Exit Sub
    
    'Warning if already revised today to add option to cancel action
    If (Cells(selectedRow, DateLastReviewColumn) = Date) Then
        promptAnswer = MsgBox("You already set a grade today. You want rewrite?", vbQuestion + vbYesNo + vbDefaultButton2, "Already set")
    End If
    If (promptAnswer = 7) Then Exit Sub
    
    'Calculate days passed from last revise
    Dim lastDateRevised     As Date
    lastDateRevised = Cells(selectedRow, DateLastReviewColumn)
    Dim daysPassed          As Integer
    daysPassed = DateDiff("d", lastDateRevised, Now)
    
    'Shift history area
    Range(HistoryStartColumn & selectedRow & ":" & HistoryBeforeLastColumn & selectedRow).Copy
    Range(HistorySecondColmn & selectedRow & ":" & HistoryEndColumn & selectedRow).PasteSpecial _
                Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Update data in cells
    Cells(selectedRow, GradeColumn) = Grade
    Cells(selectedRow, DateLastReviewColumn) = Date
    Cells(selectedRow, HistoryStartColumn) = daysPassed
    
    'Color current History refer to answer
    If (Grade = "Good") Then Cells(selectedRow, HistoryStartColumn).Interior.Color = RGB(226, 239, 218)
    If (Grade = "Ok") Then Cells(selectedRow, HistoryStartColumn).Interior.Color = RGB(255, 242, 204)
    If (Grade = "Bad") Then Cells(selectedRow, HistoryStartColumn).Interior.Color = RGB(255, 199, 206)

End Sub


Sub QAE_ApplyConditionalFormating()
    '-----------------------------------
    'Important precondition in Excel:
    'You should define named ranges: goodTimeout, okTimeout, badTimeout
    'They should be anywere in excel and it is just 1 cell each with number of days you want to set for color change
    '-----------------------------------
    
    'Set Active color range
    Dim StrongColorRange    As Range
    Dim LightColorRange     As Range
    Set StrongColorRange = Range(FirstColumnAreaToColor & ":" & LastColumnAreaToColor)
    Set LightColorRange = Range(FirstColumnAreaToColor & ":" & FirstColumnAreaToColor)
    
    'Delete Existing Conditional Formatting from Range
    StrongColorRange.FormatConditions.Delete
    LightColorRange.FormatConditions.Delete
    
    'New question answered bad
    LightColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=""Bad"",$" & DateLastReviewColumn & "1+badTimeout>TODAY())"
    With LightColorRange.FormatConditions(1)
            .Interior.Color = RGB(255, 240, 240)
            .Font.Color = RGB(128, 128, 128)
            .StopIfTrue = False
    End With
    
    'New question answered Ok
    LightColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=""Ok"",$" & DateLastReviewColumn & "1+okTimeout>TODAY())"
    With LightColorRange.FormatConditions(2)
            .Interior.Color = RGB(255, 250, 220)
            .Font.Color = RGB(128, 128, 128)
            .StopIfTrue = False
    End With
    
    'New question answered Good
    LightColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=""Good"",$" & DateLastReviewColumn & "1+goodTimeout>TODAY())"
    With LightColorRange.FormatConditions(3)
            .Interior.Color = RGB(230, 250, 230)
            .Font.Color = RGB(128, 128, 128)
            .StopIfTrue = False
    End With
    
    
    
    'Old question answered bad
    StrongColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=""Bad"",$" & DateLastReviewColumn & "1+badTimeout<=TODAY())"
    With StrongColorRange.FormatConditions(4)
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(0, 97, 0)
            .StopIfTrue = False
    End With

    'Old question answered Ok
    StrongColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=""Ok"",$" & DateLastReviewColumn & "1+okTimeout<=TODAY())"
    With StrongColorRange.FormatConditions(5)
            .Interior.Color = RGB(255, 235, 156)
            .Font.Color = RGB(156, 101, 0)
            .StopIfTrue = False
    End With

    'Old question answered Good
    StrongColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=""Good"",$" & DateLastReviewColumn & "1+goodTimeout<=TODAY())"
    With StrongColorRange.FormatConditions(6)
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 97, 0)
            .StopIfTrue = False
    End With


End Sub