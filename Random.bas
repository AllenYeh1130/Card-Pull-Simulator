Attribute VB_Name = "Random"
' �H����P��
' n = �d�]���� (1=��, 2=��, 3=��, 4=��, 5=��)
Function GetRandomStar(n As Integer) As String
    ' �w�q�ܼ�
    Dim CardWeightRange As Range
    Dim CardWeightArray() As Double
    Dim CardWeightTotal As Double
    Dim CardWeightCount As Integer
    Dim CardWeigtRnd As Double
    Dim CardCumWeight As Double
    Dim j As Integer
    Dim StarRnd As String
    Dim CardRatio As Worksheet
    
    Set CardRatio = Worksheets("�d�]���v")
    
    ' �C���I�s���n���`�v���k�s
    CardWeightTotal = 0
    
    ' ���o�ӥd�]���Ҧ��v���d��
    Set CardWeightRange = CardRatio.Range(CardRatio.Cells(3, 1 + 2 * n), CardRatio.Cells(100, 1 + 2 * n))
    
    ' �p�⦳�X���v��
    CardWeightCount = WorksheetFunction.CountA(CardWeightRange)
    
    ' �]�w�v��array�j�p
    ReDim CardWeightArray(1 To CardWeightCount)
    
    ' ��J�d���v���íp���`�v��
    For w = 1 To CardWeightCount
        CardWeightArray(w) = CardRatio.Cells(2 + w, 1 + 2 * n).Value
        ' �p���`�v��
        CardWeightTotal = CardWeightTotal + CardWeightArray(w)
    Next w
    
    ' ��l���H���ƥͦ���
    Randomize
    
    ' �H�����
    CardWeigtRnd = CardWeightTotal * Rnd
    
    ' �C���q�歫�]�ֿn�v��
    CardCumWeight = 0
    
    ' �ھڲֿn�v����P��
    For j = 1 To UBound(CardWeightArray)
        ' �ֿn�v��
        CardCumWeight = CardCumWeight + CardWeightArray(j)
        
        ' �βֿn�v���P�_�O�_����즹�d��
        If CardWeigtRnd <= CardCumWeight Then
            StarRnd = CardRatio.Cells(2 + j, 1).Value
            Exit For
        End If
    Next j
    
    ' ��^���G
    GetRandomStar = StarRnd
End Function


' �H���d�]�P���H����ID
Function GetRandomID(StarRnd As String)
    ' �H���P�ƥd��
    Set ID = Worksheets("�d���s��")
    
    ' ���P�Ƥǰt����
    MatchRow = Application.Match(StarRnd, ID.Range("A:A"), 0)
    
    'MsgBox "Match�ĴX�C: " & MatchRow
    
    ' �]�w�d��A���]�ƾڱq B ��}�l
    Set IDValues = ID.Range(ID.Cells(MatchRow, 2), ID.Cells(MatchRow, ID.Cells(MatchRow, Columns.Count).End(xlToLeft).Column))
    
    ' �H������d�򤺪��Y�ӭ�
    RandomIndex = WorksheetFunction.RandBetween(1, IDValues.Cells.Count)
    cardID = IDValues.Cells(RandomIndex).Value
    
    ' ��^���G
    GetRandomID = cardID
End Function
