Attribute VB_Name = "RangeMod"
Option Explicit

Public Function GetRange() As Range

    Set GetRange = shStaff.Range("A1").CurrentRegion
    Set GetRange = GetRange.Offset(1).Resize(GetRange.Rows.Count - 1)
    
End Function

Public Sub DeleteRow(ByVal row As Long)
    shStaff.Range("A2").Offset(row).EntireRow.Delete
End Sub

