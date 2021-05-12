Attribute VB_Name = "Ä£¿é2"
Function VlookUpNo256Limit(searchCell, A As Range, coloum_idx)
    Dim arr
    arr = A
    'MsgBox "LBound(arr)" & LBound(arr) & "UBound(arr)" & UBound(arr) & "LBound(arr, 2)" & LBound(arr, 2) & "UBound(arr, 2)" & UBound(arr, 2)
    For i = LBound(arr) To UBound(arr)
        For j = LBound(arr, 2) To UBound(arr, 2)
             If A.Cells(i, j).Value = searchCell.Value Then
                VlookUpNo256Limit = A(i, coloum_idx)
                Exit For
             End If
        Next
    Next
End Function
