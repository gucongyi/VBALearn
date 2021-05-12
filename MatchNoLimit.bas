Attribute VB_Name = "Ä£¿é3"
Function MatchNoLimit(searchCell, A As Range)
    Dim arr
    arr = A
    For i = LBound(arr) To UBound(arr)
        For j = LBound(arr, 2) To UBound(arr, 2)
             If A.Cells(i, j).Value = searchCell.Value Then
                MatchNoLimit = i
                Exit For
             End If
        Next
    Next
End Function
