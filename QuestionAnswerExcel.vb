Option Explicit

Sub QAE_SetGradeGreat()
    Call QAE_UpdateGrade("G")
End Sub
Sub QAE_SetGradeGreatOk()
    Call QAE_UpdateGrade("O")
End Sub
Sub QAE_SetGradeGreatBad()
    Call QAE_UpdateGrade("B")
End Sub

Sub QAE_UpdateGrade(Grade As String)
    
    Const startingRow As Integer = 7
    Const dateColumn As String = "K"
    Const GradeColumn As String = "J"
    Dim today As Date
    Dim currentRow As Integer
    Dim cellValue As String
    
    currentRow = ActiveCell.row
    today = Date
    cellValue = Cells(currentRow, GradeColumn)
    
    If (currentRow < startingRow) Then Exit Sub
    If (Cells(currentRow, dateColumn) = Date) Then
        promptAnswer = MsgBox("You already set a grade today. You want rewrite?", vbQuestion + vbYesNo + vbDefaultButton2, "Already set")
    End If
    
    If (promptAnswer = 7) Then Exit Sub
    If (promptAnswer = 6) Then cellValue = Left(cellValue, Len(cellValue) - 1)
    
    Cells(currentRow, GradeColumn) = cellValue + Grade
    Cells(currentRow, dateColumn) = Date

End Sub


Sub QAE_ApplyConditionalFormating()
    
    '-----------------------------------
    'Important precondition in Excel:
    'You should define named ranges: goodTimeout, okTimeout, badTimeout
    'They should be anywere in excel and it is just 1 cell each with number of days you want to set for color change
    '-----------------------------------
    
    Const FirstColumnAreaToColor    As String = "B"
    Const LastColumnAreaToColor     As String = "E"
    Const GradeColumn               As String = "D"
    Const DateLastReviewColumn      As String = "E"
    
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
            "=AND($" & GradeColumn & "1=3,$" & DateLastReviewColumn & "1+badTimeout>TODAY())"
    With LightColorRange.FormatConditions(1)
            .Interior.Color = RGB(255, 240, 240)
            .Font.Color = RGB(128, 128, 128)
            .StopIfTrue = False
    End With
    
    'New question answered Ok
    LightColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=2,$" & DateLastReviewColumn & "1+okTimeout>TODAY())"
    With LightColorRange.FormatConditions(2)
            .Interior.Color = RGB(255, 250, 220)
            .Font.Color = RGB(128, 128, 128)
            .StopIfTrue = False
    End With
    
    'New question answered Good
    LightColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=1,$" & DateLastReviewColumn & "1+goodTimeout>TODAY())"
    With LightColorRange.FormatConditions(3)
            .Interior.Color = RGB(230, 250, 230)
            .Font.Color = RGB(128, 128, 128)
            .StopIfTrue = False
    End With
    
    
    
    'Old question answered bad
    StrongColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=3,$" & DateLastReviewColumn & "1+badTimeout<=TODAY())"
    With StrongColorRange.FormatConditions(4)
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(0, 97, 0)
            .StopIfTrue = False
    End With

    'Old question answered Ok
    StrongColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=2,$" & DateLastReviewColumn & "1+okTimeout<=TODAY())"
    With StrongColorRange.FormatConditions(5)
            .Interior.Color = RGB(255, 235, 156)
            .Font.Color = RGB(156, 101, 0)
            .StopIfTrue = False
    End With

    'Old question answered Good
    StrongColorRange.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & GradeColumn & "1=1,$" & DateLastReviewColumn & "1+goodTimeout<=TODAY())"
    With StrongColorRange.FormatConditions(6)
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 97, 0)
            .StopIfTrue = False
    End With


End Sub

