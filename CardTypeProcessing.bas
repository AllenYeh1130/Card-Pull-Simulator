Attribute VB_Name = "CardTypeProcessing"
' 隨機卡包抽選紀錄
Sub RandomCard(RandomCardNumRng As Range)
    Dim ws As Worksheet
    Dim RandomCardNum As Range
    Dim n As Integer
    Dim CardColor As String
    Dim i As Integer
    Dim StarRnd As String
    Dim cardID As String
    
    ' 設定工作表 ws
    Set ws = Worksheets("主要運算")

    ' 隨機卡包類別初始化
    n = 0
    
    ' 隨著迴圈跑卡包的相對應數量
    For Each RandomCardNum In RandomCardNumRng
        ' n = 卡包類型 (1=綠, 2=藍, 3=粉, 4=紫, 5=金)
        n = n + 1
        
        ' 卡包顏色
        CardColor = ws.Range("A" & (1 + n)).Value
        
        ' 該卡包要抽幾次
        For i = 1 To RandomCardNum.Value
            ' 隨機星數
            StarRnd = GetRandomStar(n)
            
            ' 星數確定隨機ID
            cardID = GetRandomID(StarRnd)
                
            ' 紀錄抽到的ID
            Call Record.recordID(cardID)
        Next i
    Next RandomCardNum
End Sub


' 固定卡包抽選紀錄
Sub ConstantCard(ConstantCardNumRng As Range)
    ' 宣告變數
    Dim ConstantCardNum As Range
    Dim CardColor As String
    Dim StarRnd As String
    Dim cardID As String
    Dim ws As Worksheet
    Dim ID As Worksheet
    Dim n As Integer
    
    ' 設定工作表 ws
    Set ws = Worksheets("主要運算")

    ' 固定卡包類別初始化
    n = 0

    For Each ConstantCardNum In ConstantCardNumRng
        ' n = 卡包類型 (1=綠, 2=藍, 3=粉, 4=紫, 5=金)
        n = n + 1
        
        ' 卡包顏色
        CardColor = ws.Range("A" & (1 + n)).Value
        
        ' 如果 n=1 也就是綠色卡包，就是再隨機抽一次
        If n = 1 Then
            ' 該卡包要抽幾次
            For i = 1 To ConstantCardNum
                ' 隨機星數
                StarRnd = GetRandomStar(n)
                
                ' 星數確定隨機ID
                cardID = GetRandomID(StarRnd)
                    
                ' 紀錄抽到的ID
                Call Record.recordID(cardID)
    
                ' 顯示結果
                'MsgBox "固定抽、CardColor: " & CardColor & ", 抽到第幾包: " & i & "/" & ConstantCardNum.Value & ", 抽到幾星: " & StarRnd & ", 抽到編號: " & CardID
            Next i
        ' 其他卡包就是固定星數
        Else
            ' 該卡包要抽幾次
            For i = 1 To ConstantCardNum
                Set ID = Worksheets("卡片編號")
                StarRnd = ID.Range("A" & (14 + n))
                
                ' 星數確定隨機ID
                cardID = GetRandomID(StarRnd)
                
                ' 紀錄抽到的ID
                Call Record.recordID(cardID)
                
                ' 顯示結果
                'MsgBox "固定抽、CardColor: " & CardColor & ", 抽到第幾包: " & i & "/" & ConstantCardNum.Value & ", 固定幾星: " & StarRnd & ", 抽到編號: " & CardID
            Next i
        End If
    Next ConstantCardNum
End Sub
