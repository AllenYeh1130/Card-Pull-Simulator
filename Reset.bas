Attribute VB_Name = "reset"
' 竚┮Τ仓縩计沮
Sub Reset()
    Dim result As Worksheet
    Set result = Worksheets("瓜挪")
    
    ' ┾Ω计
    result.Range("D2:D61") = 0
    ' ЧΘ计
    result.Range("E2:E62") = 0
    ' 肂み计
    result.Range("I2:I62") = 0
End Sub


' 竚–璸衡计沮
Sub Reset_Num()
    Dim result As Worksheet
    Dim ws As Worksheet
    Set result = Worksheets("瓜挪")
    Set ws = Worksheets("璶笲衡")
    
    ' ┾Ω计
    result.Range("H2:H61") = 0
    ' ЧΘ计
    result.Range("G2:G62") = 0
    ' 肂み计
    result.Range("J2:J62") = 0
    ' 贱纘计
    ws.Range("F2:F6") = 0
    
 End Sub

