Attribute VB_Name = "Main"
Sub main()
    ' �����B��ɬ��� (�įണ�ɫܦh)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' �w�q�ܼ�
    Dim ws As Worksheet
    Dim RandomCardNumRng As Range
    Dim ConstantCardNumRng As Range
    Dim RewardCardNumRng As Range
    Dim Num As Long
    Dim n As Integer
    Dim i As Integer
    
    ' �u�@��]�w
    Set ws = Worksheets("�D�n�B��")
    
    ' ��l�ƦU���G
    Call Reset.Reset
    
    ' �n�]�X�ӤH
    Num = ws.Range("B12").Value
    
    ' �C�ӤH���]
    For p = 1 To Num
        
        ' ��l�ƨC�H���G
        Call Reset.Reset_Num

        ''' �U�H���d�]
        ' �H���d�]���ƶq
        Set RandomCardNumRng = ws.Range("E2:E6")
        ' ���H���d�]
        Call CardTypeProcessing.RandomCard(RandomCardNumRng)
        
        
        ''' �U�T�w�d�]
        ' �T�w�d�]���ƶq
        Set ConstantCardNumRng = ws.Range("C2:C6")
        ' ��T�w�d�]
        Call CardTypeProcessing.ConstantCard(ConstantCardNumRng)
        
        ''' �B�~�ߤֿ߲n��I�����y
        ' �P�_�C�ӤH�B�~�߼�
        Call Record.RecordStar
        ' ���B�~�߼ƴ����y
        Call Reward.StarToCard
        ' ���y�d�]��
        Set RewardCardNumRng = ws.Range("F2:F6")
        ' ����y���H���d�]
        Call CardTypeProcessing.RandomCard(RewardCardNumRng)
        ' ����y���T�w�d�]
        Call CardTypeProcessing.ConstantCard(RewardCardNumRng)
        
        ''' �P�_�O�_����Set (�USet�B����Set)
        Call Record.RecordSet
        ' �����ֿn�ƾ� (Set�B�B�~�߼�)
        Call Record.SumResult
        
    Next p
    
    ' ���s�ҥιB��ɬ���
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
