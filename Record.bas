Attribute VB_Name = "Record"
' �����ֿn����
Sub recordID(cardID As String)
    Dim result As Worksheet
    Set result = Worksheets("�d����Ų")
    
    ' ����ID�����@�C
    MatchRow = Application.Match(cardID, result.Range("A:A"), 0)
    ' �ֿn����
    result.Cells(MatchRow, 4).Value = result.Cells(MatchRow, 4).Value + 1
    ' �j��C�H�����O��
    result.Cells(MatchRow, 7).Value = 1
    ' �j��C�H��������
    result.Cells(MatchRow, 8).Value = result.Cells(MatchRow, 8).Value + 1
End Sub


' �����ֿnSET����
Sub RecordSet()
    Dim result As Worksheet
    Dim AllSet As Integer
    Set result = Worksheets("�d����Ų")
    
    ' �����USET��������
    UpdateSetCompletion result, "G3:G11", "G2"
    UpdateSetCompletion result, "G13:G21", "G12"
    UpdateSetCompletion result, "G23:G31", "G22"
    UpdateSetCompletion result, "G33:G41", "G32"
    UpdateSetCompletion result, "G43:G51", "G42"
    UpdateSetCompletion result, "G53:G61", "G52"
    
    '������SET��������
    AllSet = Application.WorksheetFunction.Sum(result.Range("G2, G12, G22, G32, G42, G52"))
    If AllSet = 6 Then
        result.Range("G62") = 1
    End If
End Sub

' �����USET��������
Private Sub UpdateSetCompletion(ByRef ws As Worksheet, ByVal checkRange As String, ByVal targetCell As String)
    ' �ˬd�d�򤺪��`�M�O�_����9�A�ç�s�ؼ��x�s��
    If Application.WorksheetFunction.Sum(ws.Range(checkRange)) = 9 Then
        ws.Range(targetCell).Value = 1
    Else
        ws.Range(targetCell).Value = 0 ' �p�G�ݭn�A�o�̥i�H�]�w��L����
    End If
End Sub


' �����B�~�߼�
Sub RecordStar()
    Dim result As Worksheet
    Dim NumId_rng As Range
    Dim NumCards_rng As Range
    Dim SumStar_rng As Range
    Dim NumCards As Long
    Dim NumStar As Long
    Dim NumID As String
    Dim i As Integer
    
    Set result = Worksheets("�d����Ų")

    ' �]�w�d��
    Set NumId_rng = result.Range("A2:A61")
    Set NumCards_rng = result.Range("H2:H61")
    Set SumStar_rng = result.Range("J2:J61")
    
    ' �M���d��ìۥ[
    For i = 1 To SumStar_rng.Cells.Count
        ' �p�G�S���d������1�A��1���N�q�O��ܤ��C�J�Ĥ@�i
        If NumCards_rng.Cells(i).Value = 0 Then
            NumCards = 0
        Else
            NumCards = NumCards_rng.Cells(i).Value - 1
        End If

        If NumCards >= 1 Then
            ' �P�_�ӥd��ID
            NumID = NumId_rng.Cells(i).Value
            ' �P�_�ӥd���O�X�P
            Stars = IDStar(NumID)
            'MsgBox "�d��ID: " & NumID & ", �P��: " & Stars
            ' �B�~�߼� (��ID�d���ƶq * ��ID�d���߼�)
            NumStar = NumCards * Stars
            ' �����ֿn�B�~�߼�
            SumStar_rng.Cells(i).Value = SumStar_rng.Cells(i).Value + NumStar
        End If
    Next i
    
    result.Cells(62, 10).Value = result.Cells(62, 10).Value + Application.WorksheetFunction.Sum(result.Range("J3:J61"))
End Sub


' �ھ�ID�P�_�P��
Function IDStar(NumID As String)
    Dim CardStarDict As Object
    Dim CardRange As Range
    Dim ID As Worksheet
    Dim i As Integer, j As Integer
    Dim starLevel As Long
    Dim cardID As String
    Set ID = Worksheets("�d���s��")
    Set CardRange = ID.Range("A15:Z19")
    
    ' �w�q�r���ܼƨӦs�x�d��ID�M�P�ƪ��������Y
    Set CardStarDict = CreateObject("Scripting.Dictionary")
    
    ' �M����椤���C�@��ID�æs�J�r��
    For i = 1 To CardRange.Rows.Count
        ' ���o�C�@�檺�P��
        starLevel = Left(CardRange.Cells(i, 1).Value, 1)
        
        ' �M����e�檺�C�@��ID
        For j = 2 To CardRange.Columns.Count
            cardID = CardRange.Cells(i, j).Value
            If cardID <> "" Then
                ' �N�d��ID�P�P�Ʀs�J�r��
                CardStarDict(cardID) = starLevel
            End If
        Next j
    Next i
    
    ' �ΤW������dictionary�P�_�ѤJ��ID�O�X�P
    Stars = CardStarDict(NumID)

    ' ��^���G
    IDStar = Stars
End Function


' �����ֿn�ƾ�
Sub SumResult()
    Dim result As Worksheet
    Dim SumSet As Range
    Dim NumSet As Range
    Dim SumStar As Range
    Dim NumStar As Range
    Dim i As Integer
    
    Set result = Worksheets("�d����Ų")
    
    ' ����Set�ƾ�
    Set SumSet = result.Range("E2:E62")
    Set NumSet = result.Range("G2:G62")
    
    ' �M���d��ìۥ[
    For i = 1 To SumSet.Cells.Count
        SumSet.Cells(i).Value = SumSet.Cells(i).Value + NumSet.Cells(i).Value
    Next i
    
    ' �B�~�߼Ƽƾ�
    Set SumStar = result.Range("I2:I62")
    Set NumStar = result.Range("J2:J62")
    
    ' �M���d��ìۥ[
    For i = 1 To SumStar.Cells.Count
        SumStar.Cells(i).Value = SumStar.Cells(i).Value + NumStar.Cells(i).Value
    Next i
End Sub
