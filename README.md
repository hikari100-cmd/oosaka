# oosaka

' 키워드에 기반한 데이터 처리
If InStr(line, "参加形式:") > 0 Then
    ws.Cells(rowNum, 2).Value = Replace(line, "参加形式:", "")
    ws.Cells(rowNum, 2).Interior.Color = RGB(255, 255, 0) ' 배경색을 노란색으로
ElseIf InStr(line, "名前:") > 0 Then
    ws.Cells(rowNum, 3).Value = Replace(line, "名前:", "")
    ws.Cells(rowNum, 3).Interior.Color = RGB(255, 255, 0)
ElseIf InStr(line, "メールアドレス:") > 0 Then
    Dim email As String
    email = Replace(line, "メールアドレス:", "")
    ws.Cells(rowNum, 4).Value = email
    ws.Cells(rowNum, 4).Interior.Color = RGB(255, 255, 0)
    ws.Cells(rowNum, 5).Value = email ' E열에도 같은 값
    ws.Cells(rowNum, 5).Interior.Color = RGB(255, 255, 0)
ElseIf InStr(line, "電話番号:") > 0 Then
    ws.Cells(rowNum, 6).Value = Replace(line, "電話番号:", "")
    ws.Cells(rowNum, 6).Interior.Color = RGB(255, 255, 0)
ElseIf InStr(line, "メッセージ:") > 0 Then
    isMessage = True ' 메세지 수집 시작
End If
