Attribute VB_Name = "Reward"
' �ֿn�P�P�I���d�]�A�ѳ̦n��̮t
Sub StarToCard()
    Dim result As Worksheet
    Dim ref As Worksheet
    Dim ws As Worksheet
    Dim Reward_rng As Range
    Dim Stars As Long
    Dim i As Integer
    Dim j As Integer
    
    Set result = Worksheets("�d����Ų")
    Set ref = Worksheets("�ѦҸ��")
    Set ws = Worksheets("�D�n�B��")
    
    ' ���o�B�~�߼�
    Stars = result.Range("J62").Value
    
    'MsgBox "�B�~�߼Ʀ�: " & Stars
    
    Set Reward_rng = ref.Range("H40:K45")
    
    'MsgBox RewardStardDemend.Columns.Count
    
    ' �f���C�Ӽ��y
    For i = 2 To Reward_rng.Columns.Count
        ' ���y�߼ƻݨD
        RewardDemend = Reward_rng.Cells(1, i).Value
        
        ' �Ӽ��y�i��o�X��
        Num = Int(Stars / RewardDemend)
        
        'MsgBox "���y��: " & Num
        
        ' �Ӽ��y�ܤ֦��@�Ӥ~�|�p��d�]
        If Num > 0 Then
            ' ���y�����ǥd�]
            For j = 2 To Reward_rng.Rows.Count
                ws.Range("F" & j).Value = ws.Range("F" & j).Value + Reward_rng.Cells(j, i).Value * Num
                
                'MsgBox Reward_rng.Cells(j, 1).Value & "�d�]: " & Reward_rng.Cells(j, i).Value * Num
            Next j
            
            'MsgBox "�`�P�P��: " & Stars & "��o���y: " & Num & "��" & Reward_rng.Cells(1, i).Value & "�P���y�A�Ѿl�P�P��: " & Stars - (Num * RewardDemend)
            
            ' �C��������y�ᦩ���P�P
            Stars = Stars - (Num * RewardDemend)

        End If
    Next i
    
    
End Sub
