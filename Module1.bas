Attribute VB_Name = "Module1"
Sub 이동()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim searchValue As Variant
    Dim foundCell As Range

    ' 현재 시트와 대상 시트 설정
    Set wsSource = ActiveSheet
    Set wsTarget = Worksheets("대상")

    ' C5에 입력된 값 가져오기
    searchValue = wsSource.Range("C5").Value

    ' 대상 시트에서 값 찾기
    Set foundCell = wsTarget.Columns("A").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)

    ' 값을 찾은 경우
    If Not foundCell Is Nothing Then
        ' 찾은 셀의 AB열로 이동
        Application.Goto Reference:=wsTarget.Range("AB" & foundCell.Row), Scroll:=True
    Else
        MsgBox "값을 찾을 수 없습니다.", vbExclamation
    End If
End Sub
