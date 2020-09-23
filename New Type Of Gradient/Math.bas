Attribute VB_Name = "Math"

Public Function Average(ByRef Nums() As Long, NumCount) As Double
    Dim cPlus
    Dim i As Integer
    
    For i = 0 To NumCount - 1
        cPlus = cPlus + Nums(i)
    Next
    Average = cPlus / NumCount
End Function
