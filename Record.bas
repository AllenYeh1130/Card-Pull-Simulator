Attribute VB_Name = "Record"
' 紀錄累積完成
Sub recordID(cardID As String)
    Dim result As Worksheet
    Set result = Worksheets("卡片圖鑑")
    
    ' 找到該ID的那一列
    MatchRow = Application.Match(cardID, result.Range("A:A"), 0)
    ' 累積次數
    result.Cells(MatchRow, 4).Value = result.Cells(MatchRow, 4).Value + 1
    ' 迴圈每人完成記錄
    result.Cells(MatchRow, 7).Value = 1
    ' 迴圈每人完成次數
    result.Cells(MatchRow, 8).Value = result.Cells(MatchRow, 8).Value + 1
End Sub


' 紀錄累積SET完成
Sub RecordSet()
    Dim result As Worksheet
    Dim AllSet As Integer
    Set result = Worksheets("卡片圖鑑")
    
    ' 紀錄各SET完成次數
    UpdateSetCompletion result, "G3:G11", "G2"
    UpdateSetCompletion result, "G13:G21", "G12"
    UpdateSetCompletion result, "G23:G31", "G22"
    UpdateSetCompletion result, "G33:G41", "G32"
    UpdateSetCompletion result, "G43:G51", "G42"
    UpdateSetCompletion result, "G53:G61", "G52"
    
    '紀錄全SET完成次數
    AllSet = Application.WorksheetFunction.Sum(result.Range("G2, G12, G22, G32, G42, G52"))
    If AllSet = 6 Then
        result.Range("G62") = 1
    End If
End Sub

' 紀錄各SET完成次數
Private Sub UpdateSetCompletion(ByRef ws As Worksheet, ByVal checkRange As String, ByVal targetCell As String)
    ' 檢查範圍內的總和是否等於9，並更新目標儲存格
    If Application.WorksheetFunction.Sum(ws.Range(checkRange)) = 9 Then
        ws.Range(targetCell).Value = 1
    Else
        ws.Range(targetCell).Value = 0 ' 如果需要，這裡可以設定其他的值
    End If
End Sub


' 紀錄額外心數
Sub RecordStar()
    Dim result As Worksheet
    Dim NumId_rng As Range
    Dim NumCards_rng As Range
    Dim SumStar_rng As Range
    Dim NumCards As Long
    Dim NumStar As Long
    Dim NumID As String
    Dim i As Integer
    
    Set result = Worksheets("卡片圖鑑")

    ' 設定範圍
    Set NumId_rng = result.Range("A2:A61")
    Set NumCards_rng = result.Range("H2:H61")
    Set SumStar_rng = result.Range("J2:J61")
    
    ' 遍歷範圍並相加
    For i = 1 To SumStar_rng.Cells.Count
        ' 如果沒有卡片不減1，減1的意義是表示不列入第一張
        If NumCards_rng.Cells(i).Value = 0 Then
            NumCards = 0
        Else
            NumCards = NumCards_rng.Cells(i).Value - 1
        End If

        If NumCards >= 1 Then
            ' 判斷該卡片ID
            NumID = NumId_rng.Cells(i).Value
            ' 判斷該卡片是幾星
            Stars = IDStar(NumID)
            'MsgBox "卡片ID: " & NumID & ", 星數: " & Stars
            ' 額外心數 (該ID卡片數量 * 該ID卡片心數)
            NumStar = NumCards * Stars
            ' 紀錄累積額外心數
            SumStar_rng.Cells(i).Value = SumStar_rng.Cells(i).Value + NumStar
        End If
    Next i
    
    result.Cells(62, 10).Value = result.Cells(62, 10).Value + Application.WorksheetFunction.Sum(result.Range("J3:J61"))
End Sub


' 根據ID判斷星數
Function IDStar(NumID As String)
    Dim CardStarDict As Object
    Dim CardRange As Range
    Dim ID As Worksheet
    Dim i As Integer, j As Integer
    Dim starLevel As Long
    Dim cardID As String
    Set ID = Worksheets("卡片編號")
    Set CardRange = ID.Range("A15:Z19")
    
    ' 定義字典變數來存儲卡片ID和星數的對應關係
    Set CardStarDict = CreateObject("Scripting.Dictionary")
    
    ' 遍歷表格中的每一個ID並存入字典
    For i = 1 To CardRange.Rows.Count
        ' 取得每一行的星數
        starLevel = Left(CardRange.Cells(i, 1).Value, 1)
        
        ' 遍歷當前行的每一個ID
        For j = 2 To CardRange.Columns.Count
            cardID = CardRange.Cells(i, j).Value
            If cardID <> "" Then
                ' 將卡片ID與星數存入字典
                CardStarDict(cardID) = starLevel
            End If
        Next j
    Next i
    
    ' 用上面做的dictionary判斷書入的ID是幾星
    Stars = CardStarDict(NumID)

    ' 返回結果
    IDStar = Stars
End Function


' 紀錄累積數據
Sub SumResult()
    Dim result As Worksheet
    Dim SumSet As Range
    Dim NumSet As Range
    Dim SumStar As Range
    Dim NumStar As Range
    Dim i As Integer
    
    Set result = Worksheets("卡片圖鑑")
    
    ' 完成Set數據
    Set SumSet = result.Range("E2:E62")
    Set NumSet = result.Range("G2:G62")
    
    ' 遍歷範圍並相加
    For i = 1 To SumSet.Cells.Count
        SumSet.Cells(i).Value = SumSet.Cells(i).Value + NumSet.Cells(i).Value
    Next i
    
    ' 額外心數數據
    Set SumStar = result.Range("I2:I62")
    Set NumStar = result.Range("J2:J62")
    
    ' 遍歷範圍並相加
    For i = 1 To SumStar.Cells.Count
        SumStar.Cells(i).Value = SumStar.Cells(i).Value + NumStar.Cells(i).Value
    Next i
End Sub
