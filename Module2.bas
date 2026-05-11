Attribute VB_Name = "Module2"
Sub 사례점검입력()

    Dim wsForm As Worksheet
    Dim wsTarget As Worksheet
    Dim findValue As Variant
    Dim r As Variant

    Set wsForm = Worksheets("회의서식")
    Set wsTarget = Worksheets("대상")

    findValue = wsForm.Range("C5").Value

    r = Application.Match(findValue, wsTarget.Columns("A"), 0)

    If IsError(r) Then
        MsgBox "대상 시트에서 해당 값을 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    wsTarget.Cells(r, 28).Value = wsForm.Range("C13").Value
    wsTarget.Cells(r, 35).Value = wsForm.Range("C14").Value
    wsTarget.Cells(r, 38).Value = wsForm.Range("C15").Value

    MsgBox "입력이 완료되었습니다.", vbInformation

End Sub

Sub 입력초기화()

    Worksheets("회의서식").Range("C13:C15").ClearContents
    
    MsgBox "입력값이 초기화되었습니다.", vbInformation

End Sub
