Attribute VB_Name = "Random"
' 隨機抽星數
' n = 卡包類型 (1=綠, 2=藍, 3=粉, 4=紫, 5=金)
Function GetRandomStar(n As Integer) As String
    ' 定義變數
    Dim CardWeightRange As Range
    Dim CardWeightArray() As Double
    Dim CardWeightTotal As Double
    Dim CardWeightCount As Integer
    Dim CardWeigtRnd As Double
    Dim CardCumWeight As Double
    Dim j As Integer
    Dim StarRnd As String
    Dim CardRatio As Worksheet
    
    Set CardRatio = Worksheets("卡包機率")
    
    ' 每次呼叫都要把總權重歸零
    CardWeightTotal = 0
    
    ' 取得該卡包的所有權重範圍
    Set CardWeightRange = CardRatio.Range(CardRatio.Cells(3, 1 + 2 * n), CardRatio.Cells(100, 1 + 2 * n))
    
    ' 計算有幾個權重
    CardWeightCount = WorksheetFunction.CountA(CardWeightRange)
    
    ' 設定權重array大小
    ReDim CardWeightArray(1 To CardWeightCount)
    
    ' 填入卡片權重並計算總權重
    For w = 1 To CardWeightCount
        CardWeightArray(w) = CardRatio.Cells(2 + w, 1 + 2 * n).Value
        ' 計算總權重
        CardWeightTotal = CardWeightTotal + CardWeightArray(w)
    Next w
    
    ' 初始化隨機數生成器
    Randomize
    
    ' 隨機抽取
    CardWeigtRnd = CardWeightTotal * Rnd
    
    ' 每次訂單重設累積權重
    CardCumWeight = 0
    
    ' 根據累積權重抽星數
    For j = 1 To UBound(CardWeightArray)
        ' 累積權重
        CardCumWeight = CardCumWeight + CardWeightArray(j)
        
        ' 用累積權重判斷是否有抽到此卡片
        If CardWeigtRnd <= CardCumWeight Then
            StarRnd = CardRatio.Cells(2 + j, 1).Value
            Exit For
        End If
    Next j
    
    ' 返回結果
    GetRandomStar = StarRnd
End Function


' 隨機卡包星數隨機抽ID
Function GetRandomID(StarRnd As String)
    ' 隨機星數卡片
    Set ID = Worksheets("卡片編號")
    
    ' 找到星數匹配的行
    MatchRow = Application.Match(StarRnd, ID.Range("A:A"), 0)
    
    'MsgBox "Match第幾列: " & MatchRow
    
    ' 設定範圍，假設數據從 B 欄開始
    Set IDValues = ID.Range(ID.Cells(MatchRow, 2), ID.Cells(MatchRow, ID.Cells(MatchRow, Columns.Count).End(xlToLeft).Column))
    
    ' 隨機選取範圍內的某個值
    RandomIndex = WorksheetFunction.RandBetween(1, IDValues.Cells.Count)
    cardID = IDValues.Cells(RandomIndex).Value
    
    ' 返回結果
    GetRandomID = cardID
End Function
