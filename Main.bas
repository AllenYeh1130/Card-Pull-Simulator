Attribute VB_Name = "Main"
Sub main()
    ' 關閉運算時紀錄 (效能提升很多)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 定義變數
    Dim ws As Worksheet
    Dim RandomCardNumRng As Range
    Dim ConstantCardNumRng As Range
    Dim RewardCardNumRng As Range
    Dim Num As Long
    Dim n As Integer
    Dim i As Integer
    
    ' 工作表設定
    Set ws = Worksheets("主要運算")
    
    ' 初始化各結果
    Call Reset.Reset
    
    ' 要跑幾個人
    Num = ws.Range("B12").Value
    
    ' 每個人都跑
    For p = 1 To Num
        
        ' 初始化每人結果
        Call Reset.Reset_Num

        ''' 各隨機卡包
        ' 隨機卡包的數量
        Set RandomCardNumRng = ws.Range("E2:E6")
        ' 抽隨機卡包
        Call CardTypeProcessing.RandomCard(RandomCardNumRng)
        
        
        ''' 各固定卡包
        ' 固定卡包的數量
        Set ConstantCardNumRng = ws.Range("C2:C6")
        ' 抽固定卡包
        Call CardTypeProcessing.ConstantCard(ConstantCardNumRng)
        
        ''' 額外心心累積後兌換獎勵
        ' 判斷每個人額外心數
        Call Record.RecordStar
        ' 用額外心數換獎勵
        Call Reward.StarToCard
        ' 獎勵卡包數
        Set RewardCardNumRng = ws.Range("F2:F6")
        ' 抽獎勵的隨機卡包
        Call CardTypeProcessing.RandomCard(RewardCardNumRng)
        ' 抽獎勵的固定卡包
        Call CardTypeProcessing.ConstantCard(RewardCardNumRng)
        
        ''' 判斷是否完成Set (各Set、全部Set)
        Call Record.RecordSet
        ' 紀錄累積數據 (Set、額外心數)
        Call Record.SumResult
        
    Next p
    
    ' 重新啟用運算時紀錄
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
