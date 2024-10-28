Attribute VB_Name = "CardTypeProcessing"
' �H���d�]������
Sub RandomCard(RandomCardNumRng As Range)
    Dim ws As Worksheet
    Dim RandomCardNum As Range
    Dim n As Integer
    Dim CardColor As String
    Dim i As Integer
    Dim StarRnd As String
    Dim cardID As String
    
    ' �]�w�u�@�� ws
    Set ws = Worksheets("�D�n�B��")

    ' �H���d�]���O��l��
    n = 0
    
    ' �H�۰j��]�d�]���۹����ƶq
    For Each RandomCardNum In RandomCardNumRng
        ' n = �d�]���� (1=��, 2=��, 3=��, 4=��, 5=��)
        n = n + 1
        
        ' �d�]�C��
        CardColor = ws.Range("A" & (1 + n)).Value
        
        ' �ӥd�]�n��X��
        For i = 1 To RandomCardNum.Value
            ' �H���P��
            StarRnd = GetRandomStar(n)
            
            ' �P�ƽT�w�H��ID
            cardID = GetRandomID(StarRnd)
                
            ' ������쪺ID
            Call Record.recordID(cardID)
        Next i
    Next RandomCardNum
End Sub


' �T�w�d�]������
Sub ConstantCard(ConstantCardNumRng As Range)
    ' �ŧi�ܼ�
    Dim ConstantCardNum As Range
    Dim CardColor As String
    Dim StarRnd As String
    Dim cardID As String
    Dim ws As Worksheet
    Dim ID As Worksheet
    Dim n As Integer
    
    ' �]�w�u�@�� ws
    Set ws = Worksheets("�D�n�B��")

    ' �T�w�d�]���O��l��
    n = 0

    For Each ConstantCardNum In ConstantCardNumRng
        ' n = �d�]���� (1=��, 2=��, 3=��, 4=��, 5=��)
        n = n + 1
        
        ' �d�]�C��
        CardColor = ws.Range("A" & (1 + n)).Value
        
        ' �p�G n=1 �]�N�O���d�]�A�N�O�A�H����@��
        If n = 1 Then
            ' �ӥd�]�n��X��
            For i = 1 To ConstantCardNum
                ' �H���P��
                StarRnd = GetRandomStar(n)
                
                ' �P�ƽT�w�H��ID
                cardID = GetRandomID(StarRnd)
                    
                ' ������쪺ID
                Call Record.recordID(cardID)
    
                ' ��ܵ��G
                'MsgBox "�T�w��BCardColor: " & CardColor & ", ���ĴX�]: " & i & "/" & ConstantCardNum.Value & ", ���X�P: " & StarRnd & ", ���s��: " & CardID
            Next i
        ' ��L�d�]�N�O�T�w�P��
        Else
            ' �ӥd�]�n��X��
            For i = 1 To ConstantCardNum
                Set ID = Worksheets("�d���s��")
                StarRnd = ID.Range("A" & (14 + n))
                
                ' �P�ƽT�w�H��ID
                cardID = GetRandomID(StarRnd)
                
                ' ������쪺ID
                Call Record.recordID(cardID)
                
                ' ��ܵ��G
                'MsgBox "�T�w��BCardColor: " & CardColor & ", ���ĴX�]: " & i & "/" & ConstantCardNum.Value & ", �T�w�X�P: " & StarRnd & ", ���s��: " & CardID
            Next i
        End If
    Next ConstantCardNum
End Sub
