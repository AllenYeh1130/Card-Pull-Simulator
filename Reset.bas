Attribute VB_Name = "reset"
' ���m�Ҧ��ֿn�ƾ�
Sub Reset()
    Dim result As Worksheet
    Set result = Worksheets("�d����Ų")
    
    ' ��d����
    result.Range("D2:D61") = 0
    ' ������
    result.Range("E2:E62") = 0
    ' �B�~�߼�
    result.Range("I2:I62") = 0
End Sub


' ���m�C�H�p��ƾ�
Sub Reset_Num()
    Dim result As Worksheet
    Dim ws As Worksheet
    Set result = Worksheets("�d����Ų")
    Set ws = Worksheets("�D�n�B��")
    
    ' ��d����
    result.Range("H2:H61") = 0
    ' ������
    result.Range("G2:G62") = 0
    ' �B�~�߼�
    result.Range("J2:J62") = 0
    ' ���y�]��
    ws.Range("F2:F6") = 0
    
 End Sub

