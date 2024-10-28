Attribute VB_Name = "Reward"
' 累積星星兌換卡包，由最好到最差
Sub StarToCard()
    Dim result As Worksheet
    Dim ref As Worksheet
    Dim ws As Worksheet
    Dim Reward_rng As Range
    Dim Stars As Long
    Dim i As Integer
    Dim j As Integer
    
    Set result = Worksheets("卡片圖鑑")
    Set ref = Worksheets("參考資料")
    Set ws = Worksheets("主要運算")
    
    ' 取得額外心數
    Stars = result.Range("J62").Value
    
    'MsgBox "額外心數有: " & Stars
    
    Set Reward_rng = ref.Range("H40:K45")
    
    'MsgBox RewardStardDemend.Columns.Count
    
    ' 審視每個獎勵
    For i = 2 To Reward_rng.Columns.Count
        ' 獎勵心數需求
        RewardDemend = Reward_rng.Cells(1, i).Value
        
        ' 該獎勵可獲得幾個
        Num = Int(Stars / RewardDemend)
        
        'MsgBox "獎勵數: " & Num
        
        ' 該獎勵至少有一個才會計算卡包
        If Num > 0 Then
            ' 獎勵有哪些卡包
            For j = 2 To Reward_rng.Rows.Count
                ws.Range("F" & j).Value = ws.Range("F" & j).Value + Reward_rng.Cells(j, i).Value * Num
                
                'MsgBox Reward_rng.Cells(j, 1).Value & "卡包: " & Reward_rng.Cells(j, i).Value * Num
            Next j
            
            'MsgBox "總星星有: " & Stars & "獲得獎勵: " & Num & "個" & Reward_rng.Cells(1, i).Value & "星獎勵，剩餘星星有: " & Stars - (Num * RewardDemend)
            
            ' 每次領取獎勵後扣除星星
            Stars = Stars - (Num * RewardDemend)

        End If
    Next i
    
    
End Sub
